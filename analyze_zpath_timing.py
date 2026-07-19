#!/usr/bin/env python
"""Analyze GO_POSITION and TURN durations within a ZPATH log interval."""

from __future__ import print_function

import argparse
import csv
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path
from statistics import mean, median


DEFAULT_START = "15:57:29.918"
DEFAULT_END = "16:38:05.550"

LOG_PATTERN = re.compile(
    r"^\[[A-Z]\s+(?P<date>\d{2}-\d{2}-\d{2})\s+"
    r"(?P<time>\d{2}:\d{2}:\d{2}\.\d{3})\b.*?\]"
    r"HSM_(?P<level>\d+)_CUTTING_STATE_MACHINE:\s+"
    r"(?P<action>Inner|Sibling|Pop)\s+:\s+(?P<state>[A-Z0-9_]+)\s*$"
)

RAW_TIMESTAMP_PATTERN = re.compile(
    r"^\[[A-Z]\s+(?P<date>\d{2}-\d{2}-\d{2})\s+"
    r"(?P<time>\d{2}:\d{2}:\d{2}\.\d{3})\b"
)

GENERIC_HSM_PATTERN = re.compile(
    r"^\[[A-Z]\s+(?P<date>\d{2}-\d{2}-\d{2})\s+"
    r"(?P<time>\d{2}:\d{2}:\d{2}\.\d{3})\b.*?\]"
    r"HSM_(?P<level>\d+)_(?P<machine>[A-Z]+)_STATE_MACHINE:\s+"
    r"(?P<action>Init|Inner|Sibling|Pop)\s+:\s+(?P<state>[A-Z0-9_]+)\s*$"
)

POINT_TYPE_PATTERN = re.compile(
    r"point type:\s*(?P<type>[012])\(0:\s*astar\s+1:\s*z long\s+2:\s*z short\)"
)

EDGE_TYPE_NAMES = {
    0: "ASTAR",
    1: "LONG_EDGE",
    2: "SHORT_EDGE",
    None: "INCOMPLETE",
}


def parse_args():
    parser = argparse.ArgumentParser(
        description="统计 ZPATH 中 GO_POSITION 及其 TURN 的时长并绘图。"
    )
    parser.add_argument("--log", default="Navigation.log", help="日志文件路径")
    parser.add_argument("--start", default=DEFAULT_START,
                        help="开始时间，HH:MM:SS.mmm 或 YY-MM-DD HH:MM:SS.mmm")
    parser.add_argument("--end", default=DEFAULT_END,
                        help="结束时间，HH:MM:SS.mmm 或 YY-MM-DD HH:MM:SS.mmm")
    parser.add_argument("--save", metavar="PNG", help="将图表保存为 PNG")
    parser.add_argument(
        "--task-save", metavar="PNG",
        help="生成日志内可观测覆盖任务的六项总览图",
    )
    parser.add_argument(
        "--task-summary", action="store_true",
        help="打印日志内可观测覆盖任务的六项汇总",
    )
    parser.add_argument("--csv", metavar="CSV", help="导出逐次统计明细")
    parser.add_argument(
        "--excel", metavar="XLSX", default="zpath_timing.xlsx",
        help="导出 Excel 汇总和明细（默认: zpath_timing.xlsx）",
    )
    parser.add_argument("--no-show", action="store_true", help="不弹出图表窗口")
    return parser.parse_args()


def parse_log(path):
    events = []
    with path.open("r", encoding="utf-8", errors="replace") as log_file:
        for line_number, line in enumerate(log_file, 1):
            match = LOG_PATTERN.match(line.rstrip())
            if not match:
                continue
            timestamp = datetime.strptime(
                match.group("date") + " " + match.group("time"),
                "%y-%m-%d %H:%M:%S.%f",
            )
            events.append({
                "time": timestamp,
                "level": int(match.group("level")),
                "action": match.group("action"),
                "state": match.group("state"),
                "line": line_number,
            })
    if not events:
        raise ValueError("日志中没有找到 CUTTING_STATE_MACHINE 状态事件")
    return events


def parse_point_types(path):
    """Read explicit Z-path destination types from the raw log."""
    point_types = []
    with path.open("r", encoding="utf-8", errors="replace") as log_file:
        for line_number, line in enumerate(log_file, 1):
            match = POINT_TYPE_PATTERN.search(line)
            if match:
                point_types.append({
                    "line": line_number,
                    "point_type": int(match.group("type")),
                })
    return point_types


def parse_raw_context(path):
    """Read the first timestamp plus CUTTING/DOCK HSM events."""
    raw_start = None
    hsm_events = []
    with path.open("r", encoding="utf-8", errors="replace") as log_file:
        for line_number, line in enumerate(log_file, 1):
            if raw_start is None:
                timestamp_match = RAW_TIMESTAMP_PATTERN.match(line)
                if timestamp_match:
                    raw_start = datetime.strptime(
                        timestamp_match.group("date") + " "
                        + timestamp_match.group("time"),
                        "%y-%m-%d %H:%M:%S.%f",
                    )
            match = GENERIC_HSM_PATTERN.match(line.rstrip())
            if not match:
                continue
            hsm_events.append({
                "time": datetime.strptime(
                    match.group("date") + " " + match.group("time"),
                    "%y-%m-%d %H:%M:%S.%f",
                ),
                "level": int(match.group("level")),
                "machine": match.group("machine"),
                "action": match.group("action"),
                "state": match.group("state"),
                "line": line_number,
            })
    if raw_start is None:
        raise ValueError("日志中没有找到有效时间戳")
    return raw_start, hsm_events


def resolve_boundary(value, reference_date):
    for fmt in ("%y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S.%f"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    try:
        clock = datetime.strptime(value, "%H:%M:%S.%f").time()
    except ValueError:
        raise ValueError("时间格式错误: {}".format(value))
    return datetime.combine(reference_date, clock)


def build_intervals(events, level, state, window_start, window_end,
                    infer_initial=False):
    """Build state intervals; a Sibling event replaces the state at its level."""
    intervals = []
    active = None
    saw_level_event = False

    for event in events:
        if event["level"] != level:
            continue
        action = event["action"]
        event_state = event["state"]

        # A log may begin while a state is already active.  If the first event
        # at this level is that state's Pop, retain the observable portion and
        # mark its start line as unknown instead of silently dropping it.
        if (infer_initial and not saw_level_event and action == "Pop"
                and event_state == state):
            synthetic_start = dict(event)
            synthetic_start["time"] = window_start
            synthetic_start["line"] = None
            active = synthetic_start
        saw_level_event = True

        if action == "Sibling":
            if active is not None:
                intervals.append((active, event))
                active = None
            if event_state == state:
                active = event
        elif action == "Inner" and event_state == state:
            if active is not None:
                intervals.append((active, event))
            active = event
        elif action == "Pop" and event_state == state and active is not None:
            intervals.append((active, event))
            active = None

    if active is not None:
        synthetic_end = dict(active)
        synthetic_end["time"] = window_end
        synthetic_end["line"] = None
        intervals.append((active, synthetic_end))

    clipped = []
    for enter, leave in intervals:
        start = max(enter["time"], window_start)
        end = min(leave["time"], window_end)
        if end > start:
            clipped.append({
                "start": start,
                "end": end,
                "duration": (end - start).total_seconds(),
                "start_line": enter["line"],
                "end_line": leave["line"],
            })
    return clipped


def intersection_seconds(first, second):
    start = max(first["start"], second["start"])
    end = min(first["end"], second["end"])
    return max(0.0, (end - start).total_seconds())


def inside_any(interval, containers):
    return any(intersection_seconds(interval, container) > 0 for container in containers)


def clip_to_containers(intervals, containers):
    """Keep only the portions of intervals that are inside parent containers."""
    clipped = []
    for interval in intervals:
        for container in containers:
            start = max(interval["start"], container["start"])
            end = min(interval["end"], container["end"])
            if end <= start:
                continue
            clipped.append({
                "start": start,
                "end": end,
                "duration": (end - start).total_seconds(),
                "start_line": (
                    interval["start_line"]
                    if start == interval["start"] else container["start_line"]
                ),
                "end_line": (
                    interval["end_line"]
                    if end == interval["end"] else container["end_line"]
                ),
            })
    return clipped


def analyze(events, window_start, window_end):
    log_start = events[0]["time"]
    log_end = events[-1]["time"]
    all_coverage = build_intervals(
        events, 1, "COVERAGE", log_start, log_end
    )
    coverage_intervals = [
        item for item in all_coverage
        if item["start"] <= window_start and item["end"] >= window_end
    ]
    all_edge_coverage = build_intervals(
        events, 2, "EDGE_COVERAGE", log_start, log_end
    )
    edge_intervals = [
        item for item in all_edge_coverage
        if inside_any(item, coverage_intervals)
    ]

    zpaths = build_intervals(events, 2, "ZPATH", window_start, window_end)
    if not zpaths:
        raise ValueError("指定时间段内没有有效的 HSM_2 ZPATH 状态")

    go_candidates = build_intervals(events, 3, "GO_POSITION", window_start, window_end)
    # GO_POSITION also exists under other HSM_2 states.  Only count the exact
    # portion where HSM_2 is ZPATH, so a child spanning a parent transition
    # cannot leak time from another parent state into the ZPATH statistics.
    go_intervals = clip_to_containers(go_candidates, zpaths)

    child_states = sorted(set(
        event["state"] for event in events
        if event["level"] == 4
        and window_start <= event["time"] <= window_end
    ))
    child_intervals = {}
    for state in child_states:
        candidates = build_intervals(events, 4, state, window_start, window_end)
        intervals = clip_to_containers(candidates, go_intervals)
        if intervals:
            child_intervals[state] = intervals

    rows = []
    for index, go in enumerate(go_intervals, 1):
        state_durations = {}
        for state, intervals in child_intervals.items():
            state_durations[state] = sum(
                intersection_seconds(go, interval) for interval in intervals
            )
        recorded_duration = sum(state_durations.values())
        rows.append({
            "index": index,
            "start": go["start"],
            "end": go["end"],
            "go_duration": go["duration"],
            "state_durations": state_durations,
            "other_duration": max(0.0, go["duration"] - recorded_duration),
            "start_line": go["start_line"],
            "end_line": go["end_line"],
        })
    return coverage_intervals, edge_intervals, zpaths, child_intervals, rows


def build_edge_traversals(rows, point_types, zpaths):
    """Merge retried GO intervals until the destination point type is logged.

    A GO_POSITION can be interrupted by VISUAL_OBS/STOP and then resumed.  The
    point-type line is emitted only when the logical destination is selected or
    reached at the end of that chain.  Counting each HSM interval separately
    would therefore turn one edge traversal into several misleading samples.
    """
    traversals = []
    pending = []
    pending_zpath = None

    def zpath_index(row):
        return next(
            (index for index, interval in enumerate(zpaths)
             if intersection_seconds(row, interval) > 0),
            None,
        )

    def finish(point_type=None, point_line=None, completed=False):
        if not pending:
            return
        traversal_index = len(traversals) + 1
        turn_duration = sum(
            row["state_durations"].get("TURN", 0.0) for row in pending
        )
        item = {
            "index": traversal_index,
            "point_type": point_type,
            "edge_type": EDGE_TYPE_NAMES[point_type],
            "completed": completed,
            "start": pending[0]["start"],
            "end": pending[-1]["end"],
            "go_duration": sum(row["go_duration"] for row in pending),
            "wall_duration": (
                pending[-1]["end"] - pending[0]["start"]
            ).total_seconds(),
            "turn_duration": turn_duration,
            "other_duration": sum(row["other_duration"] for row in pending),
            "segment_count": len(pending),
            "start_line": pending[0]["start_line"],
            "end_line": pending[-1]["end_line"],
            "point_line": point_line,
        }
        traversals.append(item)
        for row in pending:
            row["traversal_index"] = traversal_index
            row["point_type"] = point_type
            row["edge_type"] = item["edge_type"]
            row["traversal_completed"] = completed
        pending[:] = []

    for row in rows:
        current_zpath = zpath_index(row)
        if pending and current_zpath != pending_zpath:
            finish()
        if not pending:
            pending_zpath = current_zpath
        pending.append(row)

        start_line = row["start_line"]
        end_line = row["end_line"]
        hits = [] if start_line is None or end_line is None else [
            item for item in point_types
            if start_line < item["line"] <= end_line
        ]
        if hits:
            point = hits[-1]
            finish(
                point_type=point["point_type"],
                point_line=point["line"],
                completed=True,
            )

    finish()
    return traversals


def build_go_rows(events, window_start, window_end, zpaths,
                  infer_initial=False):
    """Build GO rows and level-4 child durations inside supplied ZPATHs."""
    go_candidates = build_intervals(
        events, 3, "GO_POSITION", window_start, window_end,
        infer_initial=infer_initial,
    )
    go_intervals = clip_to_containers(go_candidates, zpaths)
    child_states = sorted(set(
        event["state"] for event in events
        if event["level"] == 4
        and window_start <= event["time"] <= window_end
    ))
    child_intervals = {}
    for state in child_states:
        candidates = build_intervals(
            events, 4, state, window_start, window_end,
            infer_initial=infer_initial,
        )
        intervals = clip_to_containers(candidates, go_intervals)
        if intervals:
            child_intervals[state] = intervals

    rows = []
    for index, go in enumerate(go_intervals, 1):
        state_durations = {}
        for state, intervals in child_intervals.items():
            state_durations[state] = sum(
                intersection_seconds(go, interval) for interval in intervals
            )
        recorded_duration = sum(state_durations.values())
        rows.append({
            "index": index,
            "start": go["start"],
            "end": go["end"],
            "go_duration": go["duration"],
            "state_durations": state_durations,
            "other_duration": max(0.0, go["duration"] - recorded_duration),
            "start_line": go["start_line"],
            "end_line": go["end_line"],
        })
    return child_intervals, rows


def analyze_observed_task(events, point_types, raw_start, hsm_events):
    """Summarize every observable COVERAGE session in this log.

    The first COVERAGE/ZPATH/GO state can be left-truncated when logging starts
    mid-task.  Such intervals are retained from the first raw timestamp and the
    summary explicitly reports that the resulting total is a lower bound.
    """
    task_end = events[-1]["time"]
    coverage = build_intervals(
        events, 1, "COVERAGE", raw_start, task_end, infer_initial=True
    )
    if not coverage:
        raise ValueError("日志内没有可观测的 COVERAGE 区间")
    task_end = coverage[-1]["end"]
    coverage = [item for item in coverage if item["start"] < task_end]

    edge = clip_to_containers(
        build_intervals(events, 2, "EDGE_COVERAGE", raw_start, task_end,
                        infer_initial=True),
        coverage,
    )
    goto_road = clip_to_containers(
        build_intervals(events, 2, "GOTO_ROAD", raw_start, task_end,
                        infer_initial=True),
        coverage,
    )
    zpaths = clip_to_containers(
        build_intervals(events, 2, "ZPATH", raw_start, task_end,
                        infer_initial=True),
        coverage,
    )
    child_intervals, rows = build_go_rows(
        events, raw_start, task_end, zpaths, infer_initial=True
    )
    traversals = build_edge_traversals(rows, point_types, zpaths)

    recharge_gaps = []
    for previous, following in zip(coverage, coverage[1:]):
        if following["start"] > previous["end"]:
            recharge_gaps.append({
                "start": previous["end"],
                "end": following["start"],
                "duration": (following["start"] - previous["end"]).total_seconds(),
            })

    dock_charge_inner = next((
        event["time"] for event in hsm_events
        if event["machine"] == "DOCK" and event["level"] == 1
        and event["action"] == "Inner" and event["state"] == "DOCK_CHARGE"
        and recharge_gaps
        and recharge_gaps[0]["start"] <= event["time"] <= recharge_gaps[0]["end"]
    ), None)
    dock_charge_pop = next((
        event["time"] for event in hsm_events
        if event["machine"] == "DOCK" and event["level"] == 1
        and event["action"] == "Pop" and event["state"] == "DOCK_CHARGE"
        and recharge_gaps
        and recharge_gaps[0]["start"] <= event["time"] <= recharge_gaps[0]["end"]
    ), None)
    cut_resume = next((
        event["time"] for event in hsm_events
        if event["machine"] == "CUTTING" and event["level"] == 0
        and event["action"] == "Init" and event["state"] == "CUT"
        and dock_charge_pop is not None
        and dock_charge_pop <= event["time"] <= recharge_gaps[0]["end"]
    ), None)

    charging_intervals = []
    recharge_operational_intervals = list(recharge_gaps)
    if recharge_gaps and dock_charge_pop and cut_resume:
        first_gap = recharge_gaps[0]
        charging_intervals = [{
            "start": dock_charge_pop,
            "end": cut_resume,
            "duration": (cut_resume - dock_charge_pop).total_seconds(),
        }]
        recharge_operational_intervals = [
            {
                "start": first_gap["start"],
                "end": dock_charge_pop,
                "duration": (dock_charge_pop - first_gap["start"]).total_seconds(),
            },
            {
                "start": cut_resume,
                "end": first_gap["end"],
                "duration": (first_gap["end"] - cut_resume).total_seconds(),
            },
        ] + recharge_gaps[1:]

    by_type = {
        edge_type: [item for item in traversals
                    if item["edge_type"] == edge_type]
        for edge_type in EDGE_TYPE_NAMES.values()
    }
    metrics = {
        "coverage": sum(item["duration"] for item in coverage),
        "edge": sum(item["duration"] for item in edge),
        # Recharge timing is movement/control overhead only.  Electrical
        # charging dwell is reported separately and excluded here.
        "recharge": sum(
            item["duration"] for item in recharge_operational_intervals
        ),
        "charging": sum(item["duration"] for item in charging_intervals),
        "recharge_full_gap": sum(item["duration"] for item in recharge_gaps),
        "goto_breakpoint": sum(item["duration"] for item in goto_road),
        # The task-level ZPATH metric counts only completed, explicitly typed
        # logical edges.  Parent-state planning/STOP time and INCOMPLETE rows
        # are excluded.
        "zpath": sum(
            item["go_duration"]
            for edge_type in ("LONG_EDGE", "SHORT_EDGE", "ASTAR")
            for item in by_type[edge_type]
        ),
        "short_edge": sum(item["go_duration"] for item in by_type["SHORT_EDGE"]),
        "long_edge": sum(item["go_duration"] for item in by_type["LONG_EDGE"]),
    }
    recharge_detail = {}
    if (recharge_gaps and dock_charge_inner and dock_charge_pop
            and cut_resume):
        first_gap = recharge_gaps[0]
        recharge_detail = {
            "return_to_charger": (
                dock_charge_pop - first_gap["start"]
            ).total_seconds(),
            "dock_navigation": (
                dock_charge_inner - first_gap["start"]
            ).total_seconds(),
            "dock_alignment": (
                dock_charge_pop - dock_charge_inner
            ).total_seconds(),
            "charging": (
                cut_resume - dock_charge_pop
            ).total_seconds(),
            "resume_after_charge": (
                first_gap["end"] - cut_resume
            ).total_seconds(),
        }
    return {
        "start": coverage[0]["start"],
        "end": coverage[-1]["end"],
        "coverage_intervals": coverage,
        "edge_intervals": edge,
        "goto_intervals": goto_road,
        "zpaths": zpaths,
        "child_intervals": child_intervals,
        "rows": rows,
        "traversals": traversals,
        "recharge_gaps": recharge_gaps,
        "recharge_operational_intervals": recharge_operational_intervals,
        "charging_intervals": charging_intervals,
        "recharge_detail": recharge_detail,
        "metrics": metrics,
        "long_count": len(by_type["LONG_EDGE"]),
        "short_count": len(by_type["SHORT_EDGE"]),
        "astar_count": len(by_type["ASTAR"]),
        "incomplete_count": len(by_type["INCOMPLETE"]),
        "raw_go_count": len(rows),
        "left_truncated": coverage[0]["start_line"] is None,
    }


def duration_text(seconds):
    milliseconds = int(round(seconds * 1000))
    hours, remainder = divmod(milliseconds, 3600000)
    minutes, remainder = divmod(remainder, 60000)
    secs, millis = divmod(remainder, 1000)
    return "{:02d}:{:02d}:{:02d}.{:03d}".format(hours, minutes, secs, millis)


def describe(values):
    return {
        "count": len(values),
        "total": sum(values),
        "average": mean(values) if values else 0.0,
        "median": median(values) if values else 0.0,
        "minimum": min(values) if values else 0.0,
        "maximum": max(values) if values else 0.0,
    }


def print_summary(window_start, window_end, coverage_intervals, edge_intervals,
                  zpaths, child_intervals, rows, traversals):
    coverage_stats = describe([item["duration"] for item in coverage_intervals])
    edge_stats = describe([item["duration"] for item in edge_intervals])
    zpath_stats = describe([item["duration"] for item in zpaths])
    go_stats = describe([row["go_duration"] for row in rows])
    state_stats = {
        state: describe([item["duration"] for item in intervals])
        for state, intervals in child_intervals.items()
    }
    traversal_stats = {
        edge_type: describe([
            item["go_duration"] for item in traversals
            if item["edge_type"] == edge_type
        ])
        for edge_type in ("LONG_EDGE", "SHORT_EDGE", "ASTAR", "INCOMPLETE")
    }

    print("统计时间段: {} - {}".format(
        window_start.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3],
        window_end.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3],
    ))
    print("ZPATH 有效时长: {} ({:.3f}s)".format(
        duration_text(sum(item["duration"] for item in zpaths)),
        sum(item["duration"] for item in zpaths),
    ))
    print("{:<24} {:>6} {:>12} {:>10} {:>10} {:>10} {:>10}".format(
        "项目", "次数", "总时长(s)", "平均(s)", "中位(s)", "最小(s)", "最大(s)"
    ))
    summary_items = [
        ("COVERAGE", coverage_stats),
        ("COVERAGE 中 EDGE", edge_stats),
        ("ZPATH", zpath_stats),
        ("GO_POSITION", go_stats),
    ]
    summary_items.extend(
        ("GO 中 " + state, state_stats[state]) for state in sorted(state_stats)
    )
    summary_items.extend(
        (edge_type, traversal_stats[edge_type])
        for edge_type in ("LONG_EDGE", "SHORT_EDGE", "ASTAR", "INCOMPLETE")
        if traversal_stats[edge_type]["count"]
    )
    for label, stats in summary_items:
        print("{:<24} {:>6} {:>12.3f} {:>10.3f} {:>10.3f} {:>10.3f} {:>10.3f}".format(
            label, stats["count"], stats["total"], stats["average"],
            stats["median"], stats["minimum"], stats["maximum"]
        ))
    coverage_total = coverage_stats["total"]
    edge_ratio = edge_stats["total"] / coverage_total if coverage_total else 0.0
    print("EDGE_COVERAGE / COVERAGE 时长占比: {:.2%}".format(edge_ratio))
    zpath_total = sum(item["duration"] for item in zpaths)
    zpath_ratio = zpath_total / coverage_total if coverage_total else 0.0
    print("ZPATH / COVERAGE 时长占比: {:.2%}".format(zpath_ratio))
    go_ratio = go_stats["total"] / zpath_total if zpath_total else 0.0
    print("GO_POSITION / ZPATH 时长占比: {:.2%}".format(go_ratio))
    for state in sorted(state_stats):
        ratio = state_stats[state]["total"] / go_stats["total"] if go_stats["total"] else 0.0
        print("{} / GO_POSITION 时长占比: {:.2%}".format(state, ratio))
    completed = [item for item in traversals if item["completed"]]
    retried = [item for item in completed if item["segment_count"] > 1]
    print("逻辑边行驶: 完成 {} 次, 其中重试 {} 次, 未完成 {} 次".format(
        len(completed), len(retried), len(traversals) - len(completed)
    ))


def export_csv(path, rows, child_states):
    with path.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.writer(output)
        header = ["序号", "GO开始", "GO结束", "GO时长(s)"]
        for state in child_states:
            header.extend([state + "时长(s)", state + "占比"])
        header.extend([
            "逻辑边序号", "边类型", "point type", "是否完成",
            "开始行", "结束行",
        ])
        writer.writerow(header)
        for row in rows:
            values = [
                row["index"],
                row["start"].strftime("%Y-%m-%d %H:%M:%S.%f")[:-3],
                row["end"].strftime("%Y-%m-%d %H:%M:%S.%f")[:-3],
                "{:.3f}".format(row["go_duration"]),
            ]
            for state in child_states:
                duration = row["state_durations"].get(state, 0.0)
                ratio = duration / row["go_duration"] if row["go_duration"] else 0.0
                values.extend(["{:.3f}".format(duration), "{:.2%}".format(ratio)])
            values.extend([
                row.get("traversal_index"), row.get("edge_type"),
                row.get("point_type"), row.get("traversal_completed"),
                row["start_line"], row["end_line"],
            ])
            writer.writerow(values)


def export_excel(path, window_start, window_end, coverage_intervals,
                 edge_intervals, zpaths, child_intervals, rows, traversals):
    try:
        from openpyxl import Workbook
        from openpyxl.chart import BarChart, Reference
        from openpyxl.styles import Alignment, Font, PatternFill
    except ImportError:
        raise RuntimeError("Excel 导出需要 openpyxl，请执行: pip install openpyxl")

    workbook = Workbook()
    summary = workbook.active
    summary.title = "汇总"
    details = workbook.create_sheet("GO明细")
    edge_details = workbook.create_sheet("边明细")
    state_details = workbook.create_sheet("状态明细")
    coverage_details = workbook.create_sheet("COVERAGE明细")
    child_states = sorted(child_intervals)
    coverage_stats = describe([item["duration"] for item in coverage_intervals])
    edge_stats = describe([item["duration"] for item in edge_intervals])
    zpath_stats = describe([item["duration"] for item in zpaths])
    go_stats = describe([row["go_duration"] for row in rows])
    zpath_total = zpath_stats["total"]

    summary.append(["ZPATH GO_POSITION 状态时长统计"])
    summary.append(["开始时间", window_start])
    summary.append(["结束时间", window_end])
    summary.append(["ZPATH有效时长(s)", zpath_total])
    summary.append([])
    summary.append(["状态", "父状态", "次数", "总时长(s)", "平均(s)", "中位(s)",
                    "最小(s)", "最大(s)", "占父状态比例", "比例说明"])

    summary_rows = [
        ("COVERAGE", "COVERAGE", coverage_stats, coverage_stats["total"]),
        ("EDGE_COVERAGE", "COVERAGE", edge_stats, coverage_stats["total"]),
        ("ZPATH", "COVERAGE", zpath_stats, coverage_stats["total"]),
        ("GO_POSITION", "ZPATH", go_stats, zpath_total),
    ]
    summary_rows.extend(
        (state, "GO_POSITION",
         describe([item["duration"] for item in child_intervals[state]]),
         go_stats["total"])
        for state in child_states
    )
    summary_rows.extend(
        (edge_type, "GO_POSITION",
         describe([item["go_duration"] for item in traversals
                   if item["edge_type"] == edge_type]),
         go_stats["total"])
        for edge_type in ("LONG_EDGE", "SHORT_EDGE", "ASTAR", "INCOMPLETE")
        if any(item["edge_type"] == edge_type for item in traversals)
    )
    for state, parent_state, stats, parent_total in summary_rows:
        ratio = stats["total"] / parent_total if parent_total else 0.0
        summary.append([
            state, parent_state, stats["count"], stats["total"], stats["average"],
            stats["median"], stats["minimum"], stats["maximum"],
            ratio,
            "{} / {} 时长占比：{:.2%}".format(state, parent_state, ratio),
        ])

    detail_header = ["序号", "GO开始", "GO结束", "GO时长(s)"]
    for state in child_states:
        detail_header.extend([state + "时长(s)", state + "占比"])
    detail_header.extend([
        "逻辑边序号", "边类型", "point type", "是否完成", "开始行", "结束行"
    ])
    details.append(detail_header)
    for row in rows:
        values = [row["index"], row["start"], row["end"], row["go_duration"]]
        for state in child_states:
            duration = row["state_durations"].get(state, 0.0)
            ratio = duration / row["go_duration"] if row["go_duration"] else 0.0
            values.extend([duration, ratio])
        values.extend([
            row.get("traversal_index"), row.get("edge_type"),
            row.get("point_type"), row.get("traversal_completed"),
            row["start_line"], row["end_line"],
        ])
        details.append(values)

    edge_details.append([
        "逻辑边序号", "边类型", "point type", "是否完成",
        "开始时间", "结束时间", "GO累计时长(s)", "墙钟时长(s)",
        "TURN累计时长(s)", "非TURN累计时长(s)", "GO分段数",
        "开始行", "结束行", "point type行",
    ])
    for item in traversals:
        edge_details.append([
            item["index"], item["edge_type"], item["point_type"],
            item["completed"], item["start"], item["end"],
            item["go_duration"], item["wall_duration"],
            item["turn_duration"], item["other_duration"],
            item["segment_count"], item["start_line"], item["end_line"],
            item["point_line"],
        ])

    state_details.append(["GO序号", "状态", "开始时间", "结束时间",
                          "时长(s)", "开始行", "结束行"])
    for state in child_states:
        for interval in child_intervals[state]:
            go_index = next(
                (row["index"] for row in rows
                 if intersection_seconds(row, interval) > 0),
                None,
            )
            state_details.append([
                go_index, state, interval["start"], interval["end"],
                interval["duration"], interval["start_line"], interval["end_line"],
            ])

    coverage_details.append([
        "类型", "序号", "开始时间", "结束时间", "时长(s)", "开始行", "结束行"
    ])
    for state, intervals in (
        ("COVERAGE", coverage_intervals),
        ("EDGE_COVERAGE", edge_intervals),
    ):
        for index, interval in enumerate(intervals, 1):
            coverage_details.append([
                state, index, interval["start"], interval["end"],
                interval["duration"], interval["start_line"], interval["end_line"],
            ])

    header_fill = PatternFill("solid", fgColor="176B87")
    header_font = Font(color="FFFFFF", bold=True)
    for sheet, header_row in (
        (summary, 6), (details, 1), (edge_details, 1),
        (state_details, 1), (coverage_details, 1)
    ):
        for cell in sheet[header_row]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
        sheet.freeze_panes = "A{}".format(header_row + 1)
        sheet.auto_filter.ref = sheet.dimensions

    summary["A1"].font = Font(size=14, bold=True)
    summary.merge_cells("A1:J1")
    summary["A1"].alignment = Alignment(horizontal="center")
    for cell in (summary["B2"], summary["B3"]):
        cell.number_format = "yyyy-mm-dd hh:mm:ss.000"
    for row_number in range(7, summary.max_row + 1):
        summary.cell(row_number, 9).number_format = "0.00%"
    for row_number in range(2, details.max_row + 1):
        details.cell(row_number, 2).number_format = "yyyy-mm-dd hh:mm:ss.000"
        details.cell(row_number, 3).number_format = "yyyy-mm-dd hh:mm:ss.000"
        for column in range(6, 4 + len(child_states) * 2, 2):
            details.cell(row_number, column).number_format = "0.00%"
    for row_number in range(2, state_details.max_row + 1):
        state_details.cell(row_number, 3).number_format = "yyyy-mm-dd hh:mm:ss.000"
        state_details.cell(row_number, 4).number_format = "yyyy-mm-dd hh:mm:ss.000"
    for row_number in range(2, edge_details.max_row + 1):
        edge_details.cell(row_number, 5).number_format = "yyyy-mm-dd hh:mm:ss.000"
        edge_details.cell(row_number, 6).number_format = "yyyy-mm-dd hh:mm:ss.000"
    for row_number in range(2, coverage_details.max_row + 1):
        coverage_details.cell(row_number, 3).number_format = "yyyy-mm-dd hh:mm:ss.000"
        coverage_details.cell(row_number, 4).number_format = "yyyy-mm-dd hh:mm:ss.000"

    widths = {
        summary: [20, 20, 12, 16, 14, 14, 14, 14, 16, 48],
        details: [8, 24, 24, 14] + [14, 12] * len(child_states) +
                 [12, 16, 12, 12, 10, 10],
        edge_details: [12, 16, 12, 12, 24, 24, 16, 16, 16, 18, 12, 10, 10, 12],
        state_details: [10, 20, 24, 24, 14, 10, 10],
        coverage_details: [20, 10, 24, 24, 14, 10, 10],
    }
    for sheet, column_widths in widths.items():
        for index, width in enumerate(column_widths, 1):
            sheet.column_dimensions[chr(64 + index)].width = width

    if summary.max_row >= 7:
        chart = BarChart()
        chart.title = "各状态总时长"
        chart.y_axis.title = "时长(s)"
        data = Reference(summary, min_col=4, min_row=6, max_row=summary.max_row)
        categories = Reference(summary, min_col=1, min_row=7, max_row=summary.max_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)
        chart.height = 7
        chart.width = 13
        summary.add_chart(chart, "L2")

    workbook.save(str(path))


def draw_chart(rows, window_start, window_end, coverage_intervals,
               edge_intervals, child_states, traversals,
               save_path=None, show=True):
    try:
        import matplotlib.pyplot as plt
    except ImportError:
        raise RuntimeError("绘图需要 matplotlib，请执行: pip install matplotlib")

    indexes = [row["index"] for row in rows]
    go_values = [row["go_duration"] for row in rows]
    state_values = {
        state: [row["state_durations"].get(state, 0.0) for row in rows]
        for state in child_states
    }
    go_total = sum(go_values)
    state_totals = {
        state: sum(values) for state, values in state_values.items()
    }
    edge_values = {
        edge_type: [item["go_duration"] for item in traversals
                    if item["edge_type"] == edge_type]
        for edge_type in ("LONG_EDGE", "SHORT_EDGE", "ASTAR", "INCOMPLETE")
    }

    fig = plt.figure(figsize=(13, 8))
    grid = fig.add_gridspec(2, 2, height_ratios=[1.25, 1])
    timeline = fig.add_subplot(grid[0, :])
    totals = fig.add_subplot(grid[1, 0])
    distribution = fig.add_subplot(grid[1, 1])

    timeline.plot(indexes, go_values, color="#176B87", linewidth=1.2,
                  marker=".", markersize=3, label="GO_POSITION")
    colors = ["#D95F02", "#5E3C99", "#1B9E77", "#E6AB02", "#7570B3"]
    for index, state in enumerate(child_states):
        timeline.plot(
            indexes, state_values[state], color=colors[index % len(colors)],
            linewidth=1.0, marker=".", markersize=3,
            label="{} in GO_POSITION".format(state),
        )
    timeline.set_xlabel("GO_POSITION sequence")
    timeline.set_ylabel("Duration (seconds)")
    timeline.set_title("Duration by event")
    timeline.grid(axis="y", alpha=0.25)
    timeline.legend()

    labels = ["COVERAGE", "EDGE_COVERAGE", "GO_POSITION"] + child_states + [
        "LONG_EDGE", "SHORT_EDGE"
    ]
    values = [
        sum(item["duration"] for item in coverage_intervals),
        sum(item["duration"] for item in edge_intervals),
        go_total,
    ] + [state_totals[state] for state in child_states] + [
        sum(edge_values["LONG_EDGE"]), sum(edge_values["SHORT_EDGE"])
    ]
    bar_colors = ["#4C4C4C", "#E6AB02", "#176B87"] + [
        colors[index % len(colors)] for index in range(len(child_states))
    ] + ["#C44E52", "#4C72B0"]
    totals.bar(labels, values, color=bar_colors)
    for index, value in enumerate(values):
        totals.text(index, value, "{:.3f}s".format(value), ha="center", va="bottom")
    totals.set_ylabel("Total duration (seconds)")
    totals.set_title("State duration totals")
    totals.grid(axis="y", alpha=0.25)
    totals.margins(y=0.12)
    for label in totals.get_xticklabels():
        label.set_rotation(20)
        label.set_horizontalalignment("right")

    classified_values = edge_values["LONG_EDGE"] + edge_values["SHORT_EDGE"]
    bins = min(30, max(8, int(len(classified_values) ** 0.5 * 2)))
    distribution.hist(edge_values["LONG_EDGE"], bins=bins,
                      color="#C44E52", alpha=0.68, label="LONG_EDGE")
    distribution.hist(edge_values["SHORT_EDGE"], bins=bins,
                      color="#4C72B0", alpha=0.68, label="SHORT_EDGE")
    distribution.set_xlabel("Duration (seconds)")
    distribution.set_ylabel("Count")
    distribution.set_title("Logical edge GO duration distribution")
    distribution.legend()
    distribution.grid(axis="y", alpha=0.25)

    fig.suptitle("ZPATH timing analysis | {} - {}".format(
        window_start.strftime("%H:%M:%S.%f")[:-3],
        window_end.strftime("%H:%M:%S.%f")[:-3],
    ), fontsize=14)
    fig.tight_layout()

    if save_path:
        fig.savefig(str(save_path), dpi=160, bbox_inches="tight")
        print("图表已保存: {}".format(save_path.resolve()))
    if show:
        plt.show()
    else:
        plt.close(fig)


def print_task_summary(task):
    labels = [
        ("覆盖作业（仅运行态）", "coverage"),
        ("沿边切割 EDGE_COVERAGE", "edge"),
        ("回充动作（返航入桩+出桩，不含充电）", "recharge"),
        ("走到断点 GOTO_ROAD", "goto_breakpoint"),
        ("ZPATH", "zpath"),
        ("短边切割（含转弯）", "short_edge"),
        ("长边切割（含转弯）", "long_edge"),
    ]
    print("\n日志内可观测覆盖任务汇总:")
    print("  {} - {}".format(
        task["start"].strftime("%Y-%m-%d %H:%M:%S.%f")[:-3],
        task["end"].strftime("%Y-%m-%d %H:%M:%S.%f")[:-3],
    ))
    for label, key in labels:
        seconds = task["metrics"][key]
        print("  {:<30} {} ({:.3f}s)".format(
            label, duration_text(seconds), seconds
        ))
    print("  长/短/ASTAR/未完成逻辑段: {}/{}/{}/{}；原始 GO 段: {}".format(
        task["long_count"], task["short_count"], task["astar_count"],
        task["incomplete_count"], task["raw_go_count"],
    ))
    if task["recharge_detail"]:
        detail = task["recharge_detail"]
        print("  回充拆分: 返航入桩 {}（其中到充电点 {}、入桩校准 {}），"
              "充电驻留 {}（已排除），出桩恢复 {}".format(
                  duration_text(detail["return_to_charger"]),
                  duration_text(detail["dock_navigation"]),
                  duration_text(detail["dock_alignment"]),
                  duration_text(detail["charging"]),
                  duration_text(detail["resume_after_charge"]),
              ))
    if task["left_truncated"]:
        print("  注意: 日志从任务运行中途开始，覆盖/长短边结果均为可观测下限。")


def draw_task_summary(task, save_path):
    try:
        import matplotlib.pyplot as plt
        from matplotlib.patches import Patch
    except ImportError:
        raise RuntimeError("绘图需要 matplotlib，请执行: pip install matplotlib")

    metric_items = [
        ("Coverage active", "coverage", "#4C4C4C"),
        ("Recharge overhead excl. charging", "recharge", "#8172B2"),
        ("Edge coverage", "edge", "#E6AB02"),
        ("Go to breakpoint", "goto_breakpoint", "#55A868"),
        ("ZPATH", "zpath", "#176B87"),
    ]
    labels = [item[0] for item in metric_items]
    values = [task["metrics"][item[1]] / 60.0 for item in metric_items]
    colors = [item[2] for item in metric_items]

    fig = plt.figure(figsize=(13, 10.2))
    grid = fig.add_gridspec(
        3, 1, height_ratios=[1, 1.65, 1.55], hspace=0.46
    )
    timeline = fig.add_subplot(grid[0])
    bars = fig.add_subplot(grid[1])
    zpath_detail = fig.add_subplot(grid[2])

    state_colors = {
        "EDGE_COVERAGE": "#E6AB02",
        "GOTO_ROAD": "#55A868",
        "ZPATH": "#176B87",
        "OTHER": "#B8B8B8",
    }
    category_intervals = (
        ("EDGE_COVERAGE", task["edge_intervals"]),
        ("GOTO_ROAD", task["goto_intervals"]),
        ("ZPATH", task["zpaths"]),
    )
    task_start = task["start"]
    task_wall_minutes = (task["end"] - task_start).total_seconds() / 60.0
    for coverage_index, coverage_interval in enumerate(
            task["coverage_intervals"], 1):
        left = (coverage_interval["start"] - task_start).total_seconds() / 60.0
        # Paint each COVERAGE span as OTHER, then overlay its HSM_2 states.
        # GO_POSITION remains hidden because it is already contained by ZPATH.
        timeline.barh(
            0, coverage_interval["duration"] / 60.0,
            left=left, height=0.42, color=state_colors["OTHER"],
        )
        for state, intervals in category_intervals:
            for interval in clip_to_containers(
                    intervals, [coverage_interval]):
                timeline.barh(
                    0, interval["duration"] / 60.0,
                    left=(interval["start"] - task_start).total_seconds() / 60.0,
                    height=0.42, color=state_colors[state],
                )
        midpoint = left + coverage_interval["duration"] / 120.0
        truncation = " (left-truncated)" if (
            coverage_index == 1 and task["left_truncated"]
        ) else ""
        timeline.text(
            midpoint, -0.30,
            "Coverage {}{}\n{}".format(
                coverage_index, truncation,
                duration_text(coverage_interval["duration"]),
            ),
            ha="center", va="top", fontsize=8.5,
        )

    for interval in task["recharge_operational_intervals"]:
        timeline.barh(
            0, interval["duration"] / 60.0,
            left=(interval["start"] - task_start).total_seconds() / 60.0,
            height=0.42, color="#8172B2",
        )
    for interval in task["charging_intervals"]:
        timeline.barh(
            0, interval["duration"] / 60.0,
            left=(interval["start"] - task_start).total_seconds() / 60.0,
            height=0.42, color="#E1E1E1",
        )

    timeline.set_xlim(0, task_wall_minutes)
    timeline.set_ylim(-0.62, 0.48)
    timeline.set_yticks([])
    timeline.set_xlabel("Elapsed minutes from observable task start")
    timeline.set_title("Single task bar with two COVERAGE segments")
    timeline.legend(
        handles=[
            Patch(color=state_colors["EDGE_COVERAGE"], label="Edge coverage"),
            Patch(color=state_colors["GOTO_ROAD"],
                  label="GOTO_ROAD / Go to breakpoint"),
            Patch(color=state_colors["ZPATH"],
                  label="ZPATH (includes GO_POSITION)"),
            Patch(color=state_colors["OTHER"], label="Other active time"),
            Patch(color="#8172B2", label="Recharge overhead"),
            Patch(color="#E1E1E1", label="Charging (excluded)"),
        ],
        loc="lower center", bbox_to_anchor=(0.5, 1.25),
        ncol=3, frameon=False, fontsize=8.5,
    )
    timeline.grid(axis="x", alpha=0.22)

    positions = list(range(len(labels)))
    bars.barh(positions, values, color=colors)
    bars.set_yticks(positions, labels)
    bars.invert_yaxis()
    bars.set_xlabel("Duration (minutes)")
    bars.set_title("Task timing metrics")
    bars.grid(axis="x", alpha=0.22)
    bars.margins(x=0.16)
    for position, value in zip(positions, values):
        seconds = value * 60.0
        bars.text(
            value, position,
            "  {:.2f} min  ({})".format(value, duration_text(seconds)),
            va="center", fontsize=10,
        )

    long_items = [
        item for item in task["traversals"]
        if item["edge_type"] == "LONG_EDGE"
    ]
    short_items = [
        item for item in task["traversals"]
        if item["edge_type"] == "SHORT_EDGE"
    ]
    astar_items = [
        item for item in task["traversals"]
        if item["edge_type"] == "ASTAR"
    ]
    for items, color, label, marker in (
        (long_items, "#C44E52", "Long edge", "o"),
        (short_items, "#4C72B0", "Short edge", "s"),
        (astar_items, "#8172B2", "ASTAR", "^"),
    ):
        indexes = [item["index"] for item in items]
        durations = [item["go_duration"] for item in items]
        zpath_detail.vlines(indexes, 0, durations, color=color, alpha=0.20,
                            linewidth=0.7)
        zpath_detail.scatter(
            indexes, durations, color=color, s=15, marker=marker,
            label="{}: {} traversals, {}".format(
                label, len(items),
                duration_text(sum(durations)),
            ),
        )
    zpath_detail.set_xlabel("Logical ZPATH traversal sequence")
    zpath_detail.set_ylabel("Duration (seconds)")
    zpath_detail.set_title(
        "ZPATH detail: long-edge, short-edge, and ASTAR durations"
    )
    zpath_detail.grid(axis="y", alpha=0.22)
    zpath_detail.legend(loc="upper right", frameon=False, ncol=3)

    note = (
        "Logical ZPATH traversals: long {} | short {} | ASTAR {} | incomplete {}"
        " | raw GO segments {}"
    ).format(
        task["long_count"], task["short_count"], task["astar_count"],
        task["incomplete_count"], task["raw_go_count"],
    )
    if task["left_truncated"]:
        note += "\nLog starts mid-task; observable totals are lower bounds."
    fig.text(0.5, 0.015, note, ha="center", fontsize=10)
    fig.savefig(str(save_path), dpi=170, bbox_inches="tight")
    plt.close(fig)
    print("任务总览图已保存: {}".format(save_path.resolve()))


def main():
    args = parse_args()
    log_path = Path(args.log)
    if not log_path.is_file():
        raise FileNotFoundError("日志文件不存在: {}".format(log_path))

    events = parse_log(log_path)
    point_types = parse_point_types(log_path)
    raw_start, hsm_events = parse_raw_context(log_path)
    reference_date = events[0]["time"].date()
    window_start = resolve_boundary(args.start, reference_date)
    window_end = resolve_boundary(args.end, reference_date)
    if window_end <= window_start:
        window_end += timedelta(days=1)

    coverage_intervals, edge_intervals, zpaths, child_intervals, rows = analyze(
        events, window_start, window_end
    )
    if not rows:
        raise ValueError("指定时间段的 ZPATH 中没有找到 GO_POSITION")

    traversals = build_edge_traversals(rows, point_types, zpaths)

    child_states = sorted(child_intervals)
    print_summary(
        window_start, window_end, coverage_intervals, edge_intervals,
        zpaths, child_intervals, rows, traversals,
    )
    if args.csv:
        csv_path = Path(args.csv)
        export_csv(csv_path, rows, child_states)
        print("明细已导出: {}".format(csv_path.resolve()))
    if args.excel:
        excel_path = Path(args.excel)
        export_excel(
            excel_path, window_start, window_end,
            coverage_intervals, edge_intervals, zpaths,
            child_intervals, rows, traversals,
        )
        print("Excel 已导出: {}".format(excel_path.resolve()))
    draw_chart(
        rows, window_start, window_end, coverage_intervals,
        edge_intervals, child_states, traversals,
        save_path=Path(args.save) if args.save else None,
        show=not args.no_show,
    )
    if args.task_summary or args.task_save:
        task = analyze_observed_task(
            events, point_types, raw_start, hsm_events
        )
        print_task_summary(task)
        if args.task_save:
            draw_task_summary(task, Path(args.task_save))


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, RuntimeError) as error:
        print("错误: {}".format(error), file=sys.stderr)
        sys.exit(1)
