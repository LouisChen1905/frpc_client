import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk
import subprocess
import tempfile
import os
import threading
import random
import shutil
import csv
import time
import re
import stat
from datetime import datetime

# List to store multiple instances
running_instances = []  # Each item: {process, port, sk, password, config_file}

# List to store running remote processes
running_remote_processes = []  # Each item: {'process': process, 'command': command_text}

imported_secret_keys = []
batch_running = False

command_file = os.path.join(os.path.dirname(__file__), 'ssh_commands.txt')
sk_history_file = os.path.join(os.path.dirname(__file__), 'sk_history.txt')
sk_history = []
MAX_SK_HISTORY = 100
ansi_escape_pattern = re.compile(r'\x1B\[[0-?]*[ -/]*[@-~]')
SUBPROCESS_TEXT_KWARGS = {
    'text': True,
    'encoding': 'utf-8',
    'errors': 'replace'
}


def append_output(message):
    output_text.insert(tk.END, message)
    output_text.see(tk.END)


def clean_terminal_output(text):
    if not text:
        return text
    return ansi_escape_pattern.sub('', text)

def load_ssh_commands():
    if not os.path.exists(command_file):
        with open(command_file, 'w', encoding='utf-8') as f:
            f.write('uname -a\n')
            f.write('id\n')
            f.write('ls -la\n')
    commands = []
    with open(command_file, 'r', encoding='utf-8') as f:
        for line in f:
            text = line.strip()
            if text:
                commands.append(text)
    return commands

def refresh_command_combo():
    commands = load_ssh_commands()
    command_combo['values'] = commands
    if commands:
        command_combo.set(commands[0])


def set_status(message):
    status_label.config(text=message)


def save_command():
    cmd = command_entry.get().strip()
    if not cmd:
        set_status('Please enter a command to save')
        return
    commands = load_ssh_commands()
    if cmd in commands:
        set_status('Command already exists')
        return
    with open(command_file, 'a', encoding='utf-8') as f:
        f.write(cmd + '\n')
    refresh_command_combo()
    set_status('Command saved')


def build_config_content(sk, bind_port):
    return f"""[common]
server_addr = 52.83.111.247
server_port = 7000

[secret_ssh_cs]
type = stcp
role = visitor
server_name = secret_ssh_{sk}
sk = {sk}
bind_addr = 127.0.0.1
bind_port = {bind_port}
"""


def create_temp_config(sk, bind_port):
    config_content = build_config_content(sk, bind_port)
    with tempfile.NamedTemporaryFile(mode='w', suffix='.ini', delete=False, encoding='utf-8') as f:
        f.write(config_content)
        return f.name


def get_frpc_executable():
    """Return the platform-specific FRPC binary stored beside this script."""
    binary_name = 'frpc.exe' if os.name == 'nt' else 'frpc'
    executable = os.path.join(os.path.dirname(__file__), binary_name)
    if not os.path.isfile(executable):
        raise FileNotFoundError(f'{binary_name} not found in the application directory')

    # Linux downloads may not retain the executable bit after being copied.
    if os.name != 'nt' and not os.access(executable, os.X_OK):
        try:
            os.chmod(executable, os.stat(executable).st_mode | stat.S_IXUSR)
        except OSError as e:
            raise PermissionError(f'Cannot make {binary_name} executable: {e}') from e
    return executable


def start_frpc_instance(sk, password, keep_running=True):
    bind_port = random.randint(10000, 65535)
    port_label.config(text=f"Next port: {bind_port}")
    temp_config = create_temp_config(sk, bind_port)

    try:
        process = subprocess.Popen(
            [get_frpc_executable(), '-c', temp_config],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            **SUBPROCESS_TEXT_KWARGS
        )
    except Exception:
        try:
            os.unlink(temp_config)
        except OSError:
            pass
        raise

    instance = {
        'process': process,
        'port': bind_port,
        'sk': sk,
        'password': password,
        'config_file': temp_config
    }

    if keep_running:
        running_instances.append(instance)
        update_instances_display()
        threading.Thread(target=read_output, args=(process, temp_config), daemon=True).start()

    return instance


def cleanup_instance(inst, remove_from_running=True):
    process = inst['process']
    if process.poll() is None:
        process.terminate()
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            process.kill()
    try:
        os.unlink(inst['config_file'])
    except OSError:
        pass
    if remove_from_running and inst in running_instances:
        running_instances.remove(inst)
        update_instances_display()


def build_ssh_command(bind_port, password, command_text):
    sshpass_name = 'sshpass.exe' if os.name == 'nt' else 'sshpass'
    sshpass_local = os.path.join(os.path.dirname(__file__), sshpass_name)
    if os.path.exists(sshpass_local):
        return [
            sshpass_local, '-p', password,
            'ssh', '-o', 'StrictHostKeyChecking=no', '-l', 'root', '127.0.0.1', '-p', str(bind_port),
            command_text
        ]
    if shutil.which('sshpass'):
        return [
            'sshpass', '-p', password,
            'ssh', '-o', 'StrictHostKeyChecking=no', '-l', 'root', '127.0.0.1', '-p', str(bind_port),
            command_text
        ]
    if shutil.which('ssh'):
        return [
            'ssh', '-o', 'StrictHostKeyChecking=no', '-l', 'root', '127.0.0.1', '-p', str(bind_port),
            command_text
        ]
    raise FileNotFoundError('sshpass / ssh not found')


def execute_remote_command_sync(bind_port, password, command_text, timeout=120):
    cmd = build_ssh_command(bind_port, password, command_text)
    completed = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        **SUBPROCESS_TEXT_KWARGS,
        timeout=timeout
    )
    return {
        'command': ' '.join(cmd),
        'returncode': completed.returncode,
        'stdout': clean_terminal_output(completed.stdout),
        'stderr': clean_terminal_output(completed.stderr)
    }


def normalize_secret_key(raw_key):
    key = raw_key.replace('\ufeff', '').strip()
    if key.upper().startswith('XRM'):
        digits = ''.join(ch for ch in key if ch.isdigit())
        if len(digits) >= 10:
            return digits[-10:]
    return key


def load_sk_history():
    """Load unique SK values, keeping the most recently used first."""
    sk_history.clear()
    if os.path.exists(sk_history_file):
        try:
            with open(sk_history_file, 'r', encoding='utf-8-sig') as f:
                for line in f:
                    key = normalize_secret_key(line)
                    if key and key not in sk_history:
                        sk_history.append(key)
        except OSError as e:
            set_status(f"Could not load SK history: {e}")
    refresh_sk_history_combo()


def save_sk_history():
    try:
        with open(sk_history_file, 'w', encoding='utf-8') as f:
            for key in sk_history[:MAX_SK_HISTORY]:
                f.write(key + '\n')
    except OSError as e:
        set_status(f"Could not save SK history: {e}")
        return False
    return True


def refresh_sk_history_combo():
    sk_entry['values'] = sk_history


def remember_secret_key(raw_key):
    key = normalize_secret_key(raw_key)
    if not key:
        return
    if key in sk_history:
        sk_history.remove(key)
    sk_history.insert(0, key)
    del sk_history[MAX_SK_HISTORY:]
    save_sk_history()
    refresh_sk_history_combo()
    sk_entry.set(key)


def delete_current_sk_history():
    key = normalize_secret_key(sk_entry.get())
    if key not in sk_history:
        set_status("Current SK is not in history")
        return
    sk_history.remove(key)
    if save_sk_history():
        refresh_sk_history_combo()
        sk_entry.set('')
        set_status("SK removed from history")


def clear_sk_history():
    sk_history.clear()
    if save_sk_history():
        refresh_sk_history_combo()
        sk_entry.set('')
        set_status("SK history cleared")


def import_secret_keys():
    file_path = filedialog.askopenfilename(
        title='Select secret key list',
        filetypes=[('Text/CSV Files', '*.txt *.csv'), ('All Files', '*.*')]
    )
    if not file_path:
        return

    keys = []
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        for line in f:
            for part in line.replace(',', '\n').splitlines():
                key = normalize_secret_key(part)
                if key:
                    keys.append(key)

    seen = set()
    unique_keys = []
    for key in keys:
        if key not in seen:
            seen.add(key)
            unique_keys.append(key)

    imported_secret_keys.clear()
    imported_secret_keys.extend(unique_keys)
    key_count_label.config(text=f"Imported keys: {len(imported_secret_keys)}")
    append_output(f"Imported {len(imported_secret_keys)} secret keys from {file_path}\n")
    set_status(f"Imported {len(imported_secret_keys)} keys")


def save_batch_results(results, output_file):
    with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['secret_key', 'return_code', 'stdout', 'stderr', 'executed_at'])
        for item in results:
            writer.writerow([
                item['secret_key'],
                item['return_code'],
                item['stdout'],
                item['stderr'],
                item['executed_at']
            ])


def run_batch_remote_commands():
    global batch_running

    if batch_running:
        set_status('Batch task is already running')
        return
    if not imported_secret_keys:
        set_status('Please import secret keys first')
        return

    password = password_combo.get().strip()
    if not password:
        set_status('Please select a password')
        return

    command_text = command_entry.get().strip() or command_combo.get().strip()
    if not command_text:
        set_status('Please enter or select a command')
        return

    output_file = filedialog.asksaveasfilename(
        title='Save batch result',
        defaultextension='.csv',
        initialfile=f"batch_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        filetypes=[('CSV Files', '*.csv')]
    )
    if not output_file:
        return

    def worker():
        global batch_running
        batch_running = True
        results = []
        append_output(f"Starting batch execution for {len(imported_secret_keys)} keys\n")
        set_status('Batch execution started')

        try:
            for index, raw_sk in enumerate(imported_secret_keys, start=1):
                sk = normalize_secret_key(raw_sk)
                append_output(f"\n[{index}/{len(imported_secret_keys)}] Processing key: {sk}\n")
                set_status(f"Processing {index}/{len(imported_secret_keys)}: {sk}")

                inst = None
                try:
                    inst = start_frpc_instance(sk, password, keep_running=False)
                    time.sleep(3)
                    result = execute_remote_command_sync(inst['port'], password, command_text)
                    append_output(result['stdout'])
                    if result['stderr']:
                        append_output("STDERR: " + result['stderr'])
                    results.append({
                        'secret_key': sk,
                        'return_code': result['returncode'],
                        'stdout': result['stdout'].strip(),
                        'stderr': result['stderr'].strip(),
                        'executed_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    })
                except subprocess.TimeoutExpired:
                    append_output(f"Command timed out for key: {sk}\n")
                    results.append({
                        'secret_key': sk,
                        'return_code': 'timeout',
                        'stdout': '',
                        'stderr': 'Command timed out',
                        'executed_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    })
                except Exception as e:
                    append_output(f"Error for key {sk}: {e}\n")
                    results.append({
                        'secret_key': sk,
                        'return_code': 'error',
                        'stdout': '',
                        'stderr': str(e),
                        'executed_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    })
                finally:
                    if inst is not None:
                        cleanup_instance(inst, remove_from_running=False)
        finally:
            save_batch_results(results, output_file)
            append_output(f"\nBatch results saved to: {output_file}\n")
            set_status(f"Batch completed, saved to {os.path.basename(output_file)}")
            batch_running = False

    threading.Thread(target=worker, daemon=True).start()


def update_instances_display():
    """Update the running instances listbox"""
    instances_listbox.delete(0, tk.END)
    for i, inst in enumerate(running_instances):
        if inst['process'].poll() is None:
            instances_listbox.insert(tk.END, f"Instance {i+1}: Port {inst['port']} - SK: {inst['sk']}")
        else:
            running_instances.remove(inst)

def update_remote_commands_display():
    """Update the running remote commands listbox"""
    remote_commands_listbox.delete(0, tk.END)
    for i, rem in enumerate(running_remote_processes):
        if rem['process'].poll() is None:
            remote_commands_listbox.insert(tk.END, f"Command {i+1}: {rem['command']}")
        else:
            running_remote_processes.remove(rem)

def run_frpc():
    sk = normalize_secret_key(sk_entry.get())
    if not sk:
        set_status("Please enter a secret key (sk)")
        return

    try:
        inst = start_frpc_instance(sk, password_combo.get(), keep_running=True)
        remember_secret_key(sk)
        set_status(f"FRP instance started on port {inst['port']}")
    except FileNotFoundError:
        binary_name = 'frpc.exe' if os.name == 'nt' else 'frpc'
        set_status(f"{binary_name} not found in the application directory")
    except Exception as e:
        set_status(f"Error running frpc: {str(e)}")

def read_output(process, config_file):
    try:
        while True:
            output = process.stdout.readline()
            if output == '' and process.poll() is not None:
                break
            if output:
                append_output(clean_terminal_output(output))
        # Read stderr
        stderr_output = clean_terminal_output(process.stderr.read())
        if stderr_output:
            append_output("STDERR: " + stderr_output)
        # Clean up
        os.unlink(config_file)
        set_status("FRP client stopped")
    except Exception as e:
        append_output(f"Error reading output: {str(e)}")

def read_remote_output(process, command_text):
    try:
        while True:
            output = process.stdout.readline()
            if output == '' and process.poll() is not None:
                break
            if output:
                append_output(clean_terminal_output(output))
        # Read stderr
        stderr_output = clean_terminal_output(process.stderr.read())
        if stderr_output:
            append_output("STDERR: " + stderr_output)
        # Check return code
        returncode = process.poll()
        if returncode == 0:
            set_status(f"Remote command '{command_text}' completed")
        else:
            set_status(f"Remote command '{command_text}' failed ({returncode})")
    except Exception as e:
        append_output(f"Error reading remote output: {str(e)}")

def stop_frpc():
    selection = instances_listbox.curselection()
    if not selection:
        set_status("Please select an instance to stop")
        return
    
    idx = selection[0]
    if idx < len(running_instances):
        inst = running_instances[idx]
        if inst['process'].poll() is None:
            cleanup_instance(inst)
            set_status("Instance stopped")
        else:
            set_status("Instance already stopped")

def stop_all():
    for inst in running_instances:
        if inst['process'].poll() is None:
            cleanup_instance(inst, remove_from_running=False)
    running_instances.clear()
    set_status("All instances stopped")
    update_instances_display()

def stop_remote_command():
    selection = remote_commands_listbox.curselection()
    if not selection:
        set_status("Please select a remote command to stop")
        return
    idx = selection[0]
    if idx < len(running_remote_processes):
        rem = running_remote_processes[idx]
        if rem['process'].poll() is None:
            rem['process'].terminate()
            running_remote_processes.pop(idx)
            set_status("Remote command stopped")
            update_remote_commands_display()
        else:
            set_status("Remote command already stopped")

def open_mobaXterm():
    selection = instances_listbox.curselection()
    if not selection:
        set_status("Please select an instance")
        return
    
    idx = selection[0]
    if idx >= len(running_instances):
        set_status("Invalid instance")
        return
    
    inst = running_instances[idx]
    bind_port = inst['port']
    selected_password = inst['password'].strip()
    
    if not selected_password:
        set_status("Please select a password")
        return
    
    mobaXterm_path = r"C:\Users\175912\Desktop\MobaXterm_Portable_v20.6\MobaXterm_Personal_20.6.exe"
    
    # Try to use sshpass if available, otherwise use plain ssh
    try:
        ssh_cmd = f"sshpass -p '{selected_password}' ssh -l root 127.0.0.1 -p {bind_port}"
        append_output(f"Launching MobaXterm command: {mobaXterm_path} -newtab {ssh_cmd}\n")
        subprocess.Popen([mobaXterm_path, "-newtab", ssh_cmd])
        set_status(f"Opened MobaXterm SSH session on port {bind_port}")
    except Exception as e:
        append_output(f"sshpass launch failed, falling back to plain ssh. Error: {e}\n")
        ssh_cmd = f"ssh -l root 127.0.0.1 -p {bind_port}"
        try:
            append_output(f"Launching MobaXterm command: {mobaXterm_path} -newtab {ssh_cmd}\n")
            subprocess.Popen([mobaXterm_path, "-newtab", ssh_cmd])
            set_status(f"Opened MobaXterm SSH session on port {bind_port}")
        except FileNotFoundError:
            set_status("MobaXterm not found at specified path")
            append_output("Error: MobaXterm not found at specified path\n")
        except Exception as e2:
            set_status(f"Error opening MobaXterm: {str(e2)}")
            append_output(f"Error opening MobaXterm: {e2}\n")


def run_remote_command():
    selection = instances_listbox.curselection()
    if not selection:
        set_status("Please select an instance")
        return
    
    idx = selection[0]
    if idx >= len(running_instances):
        set_status("Invalid instance")
        return
    
    inst = running_instances[idx]
    bind_port = inst['port']
    password = inst['password'].strip()
    if not password:
        set_status("Please select a password")
        return
    
    command_text = command_entry.get().strip()
    if not command_text:
        command_text = command_combo.get().strip()
    if not command_text:
        set_status("Please enter or select a command")
        return

    try:
        cmd = build_ssh_command(bind_port, password, command_text)
        append_output(f"Executing remote command: {' '.join(cmd)}\n")
        try:
            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, **SUBPROCESS_TEXT_KWARGS)
            threading.Thread(target=read_remote_output, args=(process, command_text), daemon=True).start()
            set_status("Remote command started")
            running_remote_processes.append({'process': process, 'command': command_text})
            update_remote_commands_display()
        except Exception as e:
            set_status(f"Error executing remote command: {e}")
            append_output(f"Exception: {e}\n")
    except FileNotFoundError:
        mobaXterm_path = r"C:\Users\175912\Desktop\MobaXterm_Portable_v20.6\MobaXterm_Personal_20.6.exe"
        ssh_cmd = f"ssh -l root 127.0.0.1 -p {bind_port} \"{command_text}\""
        append_output("sshpass and ssh not found on PATH.\n")
        append_output(f"Opening MobaXterm for manual execution: {mobaXterm_path} -newtab {ssh_cmd}\n")
        try:
            subprocess.Popen([mobaXterm_path, "-newtab", ssh_cmd])
            set_status("Opened MobaXterm for manual remote command")
        except FileNotFoundError:
            set_status("MobaXterm not found at specified path")
            append_output("Error: MobaXterm not found at specified path\n")
        except Exception as e:
            set_status(f"Error opening MobaXterm: {e}")
            append_output(f"Error opening MobaXterm: {e}\n")

# Create GUI
root = tk.Tk()
root.title("FRP Client Runner")

# Create frames
left_frame = tk.Frame(root)
left_frame.pack(side=tk.LEFT, fill=tk.Y)

right_frame = tk.Frame(root)
right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

# Output text in left frame
output_text = scrolledtext.ScrolledText(right_frame)
output_text.pack(fill=tk.BOTH, expand=True)

# Controls in right frame
tk.Label(left_frame, text="Secret Key (sk):").pack(pady=5)
sk_entry = ttk.Combobox(left_frame, width=38)
sk_entry.pack(pady=5)

sk_history_button_frame = tk.Frame(left_frame)
sk_history_button_frame.pack(pady=2)
delete_sk_button = tk.Button(
    sk_history_button_frame, text="Delete Current History", command=delete_current_sk_history
)
delete_sk_button.pack(side=tk.LEFT, padx=5)
clear_sk_button = tk.Button(
    sk_history_button_frame, text="Clear SK History", command=clear_sk_history
)
clear_sk_button.pack(side=tk.LEFT, padx=5)

batch_key_frame = tk.Frame(left_frame)
batch_key_frame.pack(pady=5)
import_keys_button = tk.Button(batch_key_frame, text="Import Key List", command=import_secret_keys)
import_keys_button.pack(side=tk.LEFT, padx=5)
batch_run_button = tk.Button(batch_key_frame, text="Batch Run Command", command=run_batch_remote_commands)
batch_run_button.pack(side=tk.LEFT, padx=5)
key_count_label = tk.Label(left_frame, text="Imported keys: 0")
key_count_label.pack(pady=5)

tk.Label(left_frame, text="SSH Password:").pack(pady=5)
password_combo = ttk.Combobox(left_frame, width=40, values=["11", "Xfesd203DSRGiedlxadF", "Psg-vsn.110", "root"], state="readonly")
password_combo.pack(pady=5)
password_combo.set("11")

tk.Label(left_frame, text="Remote SSH Command:").pack(pady=5)
command_entry = tk.Entry(left_frame, width=80)
command_entry.pack(pady=5)
command_combo = ttk.Combobox(left_frame, width=80, state="readonly")
command_combo.pack(pady=5)
command_save_frame = tk.Frame(left_frame)
command_save_frame.pack(pady=5)
command_save_button = tk.Button(command_save_frame, text="Save Command", command=save_command)
command_save_button.pack(side=tk.LEFT, padx=5)
command_refresh_button = tk.Button(command_save_frame, text="Refresh Commands", command=refresh_command_combo)
command_refresh_button.pack(side=tk.LEFT, padx=5)

port_label = tk.Label(left_frame, text="Next port: Not set")
port_label.pack(pady=5)

tk.Label(left_frame, text="Running Instances:").pack(pady=5)
instances_listbox = tk.Listbox(left_frame, width=50, height=6)
instances_listbox.pack(pady=5)

tk.Label(left_frame, text="Running Remote Commands:").pack(pady=5)
remote_commands_listbox = tk.Listbox(left_frame, width=50, height=4)
remote_commands_listbox.pack(pady=5)

button_frame = tk.Frame(left_frame)
button_frame.pack(pady=5)
run_button = tk.Button(button_frame, text="Run FRPC", command=run_frpc)
run_button.pack(side=tk.LEFT, padx=5)
stop_button = tk.Button(button_frame, text="Stop Selected", command=stop_frpc)
stop_button.pack(side=tk.LEFT, padx=5)
stop_all_button = tk.Button(button_frame, text="Stop All", command=stop_all)
stop_all_button.pack(side=tk.LEFT, padx=5)
mobaXterm_button = tk.Button(button_frame, text="Open MobaXterm SSH", command=open_mobaXterm)
mobaXterm_button.pack(side=tk.LEFT, padx=5)
remote_button = tk.Button(button_frame, text="Run Remote Command", command=run_remote_command)
remote_button.pack(side=tk.LEFT, padx=5)
stop_remote_button = tk.Button(button_frame, text="Stop Remote Command", command=stop_remote_command)
stop_remote_button.pack(side=tk.LEFT, padx=5)

status_label = tk.Label(left_frame, text="")
status_label.pack(pady=5)

load_sk_history()
refresh_command_combo()
root.resizable(True, True)
root.mainloop()
