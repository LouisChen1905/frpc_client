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


def build_intervals(events, level, state, window_start, window_end):
    """Build state intervals; a Sibling event replaces the state at its level."""
    intervals = []
    active = None

    for event in events:
        if event["level"] != level:
            continue
        action = event["action"]
        event_state = event["state"]

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
    go_intervals = [item for item in go_candidates if inside_any(item, zpaths)]

    child_states = sorted(set(
        event["state"] for event in events
        if event["level"] == 4
        and window_start <= event["time"] <= window_end
    ))
    child_intervals = {}
    for state in child_states:
        candidates = build_intervals(events, 4, state, window_start, window_end)
        intervals = [item for item in candidates if inside_any(item, go_intervals)]
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
                  zpaths, child_intervals, rows):
    coverage_stats = describe([item["duration"] for item in coverage_intervals])
    edge_stats = describe([item["duration"] for item in edge_intervals])
    zpath_stats = describe([item["duration"] for item in zpaths])
    go_stats = describe([row["go_duration"] for row in rows])
    state_stats = {
        state: describe([item["duration"] for item in intervals])
        for state, intervals in child_intervals.items()
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


def export_csv(path, rows, child_states):
    with path.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.writer(output)
        header = ["序号", "GO开始", "GO结束", "GO时长(s)"]
        for state in child_states:
            header.extend([state + "时长(s)", state + "占比"])
        header.extend(["开始行", "结束行"])
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
            values.extend([row["start_line"], row["end_line"]])
            writer.writerow(values)


def export_excel(path, window_start, window_end, coverage_intervals,
                 edge_intervals, zpaths, child_intervals, rows):
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
    detail_header.extend(["开始行", "结束行"])
    details.append(detail_header)
    for row in rows:
        values = [row["index"], row["start"], row["end"], row["go_duration"]]
        for state in child_states:
            duration = row["state_durations"].get(state, 0.0)
            ratio = duration / row["go_duration"] if row["go_duration"] else 0.0
            values.extend([duration, ratio])
        values.extend([row["start_line"], row["end_line"]])
        details.append(values)

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
        (summary, 6), (details, 1), (state_details, 1), (coverage_details, 1)
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
    for row_number in range(2, coverage_details.max_row + 1):
        coverage_details.cell(row_number, 3).number_format = "yyyy-mm-dd hh:mm:ss.000"
        coverage_details.cell(row_number, 4).number_format = "yyyy-mm-dd hh:mm:ss.000"

    widths = {
        summary: [20, 20, 12, 16, 14, 14, 14, 14, 16, 48],
        details: [8, 24, 24, 14] + [14, 12] * len(child_states) + [10, 10],
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
               edge_intervals, child_states,
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

    labels = ["COVERAGE", "EDGE_COVERAGE", "GO_POSITION"] + child_states
    values = [
        sum(item["duration"] for item in coverage_intervals),
        sum(item["duration"] for item in edge_intervals),
        go_total,
    ] + [state_totals[state] for state in child_states]
    bar_colors = ["#4C4C4C", "#E6AB02", "#176B87"] + [
        colors[index % len(colors)] for index in range(len(child_states))
    ]
    totals.bar(labels, values, color=bar_colors)
    for index, value in enumerate(values):
        totals.text(index, value, "{:.3f}s".format(value), ha="center", va="bottom")
    totals.set_ylabel("Total duration (seconds)")
    totals.set_title("State duration totals")
    totals.grid(axis="y", alpha=0.25)

    bins = min(30, max(8, int(len(rows) ** 0.5 * 2)))
    distribution.hist(go_values, bins=bins, color="#176B87", alpha=0.72,
                      label="GO_POSITION")
    for index, state in enumerate(child_states):
        distribution.hist(
            state_values[state], bins=bins, color=colors[index % len(colors)],
            alpha=0.65, label=state,
        )
    distribution.set_xlabel("Duration (seconds)")
    distribution.set_ylabel("Count")
    distribution.set_title("Duration distribution")
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


def main():
    args = parse_args()
    log_path = Path(args.log)
    if not log_path.is_file():
        raise FileNotFoundError("日志文件不存在: {}".format(log_path))

    events = parse_log(log_path)
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

    child_states = sorted(child_intervals)
    print_summary(
        window_start, window_end, coverage_intervals, edge_intervals,
        zpaths, child_intervals, rows,
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
            child_intervals, rows,
        )
        print("Excel 已导出: {}".format(excel_path.resolve()))
    draw_chart(
        rows, window_start, window_end, coverage_intervals,
        edge_intervals, child_states,
        save_path=Path(args.save) if args.save else None,
        show=not args.no_show,
    )


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, RuntimeError) as error:
        print("错误: {}".format(error), file=sys.stderr)
        sys.exit(1)
