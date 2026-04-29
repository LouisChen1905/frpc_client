import tkinter as tk
from tkinter import scrolledtext, ttk
import subprocess
import tempfile
import os
import threading
import random

# List to store multiple instances
running_instances = []  # Each item: {process, port, sk, password, config_file}

def update_instances_display():
    """Update the running instances listbox"""
    instances_listbox.delete(0, tk.END)
    for i, inst in enumerate(running_instances):
        if inst['process'].poll() is None:
            instances_listbox.insert(tk.END, f"Instance {i+1}: Port {inst['port']} - SK: {inst['sk']}")
        else:
            running_instances.remove(inst)

def run_frpc():
    sk = sk_entry.get().strip()
    if not sk:
        status_label.config(text="Please enter a secret key (sk)")
        return

    # Generate random bind_port
    bind_port = random.randint(10000, 65535)

    # Update port label
    port_label.config(text=f"Next port: {bind_port}")

    # Generate config content
    config_content = f"""[common]
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

    # Create temporary config file
    with tempfile.NamedTemporaryFile(mode='w', suffix='.ini', delete=False) as f:
        f.write(config_content)
        temp_config = f.name

    try:
        # Run frpc.exe with the config
        process = subprocess.Popen(['frpc.exe', '-c', temp_config], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        
        # Add to running instances
        running_instances.append({
            'process': process,
            'port': bind_port,
            'sk': sk,
            'password': password_combo.get(),
            'config_file': temp_config
        })
        
        status_label.config(text=f"FRP instance started on port {bind_port}")
        update_instances_display()
        
        # Start thread to read output
        threading.Thread(target=read_output, args=(process, temp_config), daemon=True).start()
    except FileNotFoundError:
        status_label.config(text="frpc.exe not found in current directory")
    except Exception as e:
        status_label.config(text=f"Error running frpc: {str(e)}")

def read_output(process, config_file):
    try:
        while True:
            output = process.stdout.readline()
            if output == '' and process.poll() is not None:
                break
            if output:
                output_text.insert(tk.END, output)
                output_text.see(tk.END)
        # Read stderr
        stderr_output = process.stderr.read()
        if stderr_output:
            output_text.insert(tk.END, "STDERR: " + stderr_output)
            output_text.see(tk.END)
        # Clean up
        os.unlink(config_file)
        status_label.config(text="FRP client stopped")
    except Exception as e:
        output_text.insert(tk.END, f"Error reading output: {str(e)}")

def stop_frpc():
    selection = instances_listbox.curselection()
    if not selection:
        status_label.config(text="Please select an instance to stop")
        return
    
    idx = selection[0]
    if idx < len(running_instances):
        inst = running_instances[idx]
        if inst['process'].poll() is None:
            inst['process'].terminate()
            try:
                os.unlink(inst['config_file'])
            except:
                pass
            running_instances.pop(idx)
            status_label.config(text="Instance stopped")
            update_instances_display()
        else:
            status_label.config(text="Instance already stopped")

def stop_all():
    for inst in running_instances:
        if inst['process'].poll() is None:
            inst['process'].terminate()
            try:
                os.unlink(inst['config_file'])
            except:
                pass
    running_instances.clear()
    status_label.config(text="All instances stopped")
    update_instances_display()

def open_mobaXterm():
    selection = instances_listbox.curselection()
    if not selection:
        status_label.config(text="Please select an instance")
        return
    
    idx = selection[0]
    if idx >= len(running_instances):
        status_label.config(text="Invalid instance")
        return
    
    inst = running_instances[idx]
    bind_port = inst['port']
    selected_password = inst['password'].strip()
    
    if not selected_password:
        status_label.config(text="Please select a password")
        return
    
    mobaXterm_path = r"C:\Users\175912\Desktop\MobaXterm_Portable_v20.6\MobaXterm_Personal_20.6.exe"
    
    # Try to use sshpass if available, otherwise use plain ssh
    try:
        ssh_cmd = f"sshpass -p '{selected_password}' ssh -l root 127.0.0.1 -p {bind_port}"
        output_text.insert(tk.END, f"Launching MobaXterm command: {mobaXterm_path} -newtab {ssh_cmd}\n")
        output_text.see(tk.END)
        subprocess.Popen([mobaXterm_path, "-newtab", ssh_cmd])
        status_label.config(text=f"Opened MobaXterm SSH session on port {bind_port}")
    except Exception as e:
        output_text.insert(tk.END, f"sshpass launch failed, falling back to plain ssh. Error: {e}\n")
        output_text.see(tk.END)
        ssh_cmd = f"ssh -l root 127.0.0.1 -p {bind_port}"
        try:
            output_text.insert(tk.END, f"Launching MobaXterm command: {mobaXterm_path} -newtab {ssh_cmd}\n")
            output_text.see(tk.END)
            subprocess.Popen([mobaXterm_path, "-newtab", ssh_cmd])
            status_label.config(text=f"Opened MobaXterm SSH session on port {bind_port}")
        except FileNotFoundError:
            status_label.config(text="MobaXterm not found at specified path")
            output_text.insert(tk.END, "Error: MobaXterm not found at specified path\n")
            output_text.see(tk.END)
        except Exception as e2:
            status_label.config(text=f"Error opening MobaXterm: {str(e2)}")
            output_text.insert(tk.END, f"Error opening MobaXterm: {e2}\n")
            output_text.see(tk.END)

# Create GUI
root = tk.Tk()
root.title("FRP Client Runner")

tk.Label(root, text="Secret Key (sk):").pack(pady=5)
sk_entry = tk.Entry(root, width=30)
sk_entry.pack(pady=5)

tk.Label(root, text="SSH Password:").pack(pady=5)
password_combo = ttk.Combobox(root, width=40, values=["11", "Xfesd203DSRGiedlxadF", "Psg-vsn.110"], state="readonly")
password_combo.pack(pady=5)
password_combo.set("11")

port_label = tk.Label(root, text="Next port: Not set")
port_label.pack(pady=5)

tk.Label(root, text="Running Instances:").pack(pady=5)
instances_listbox = tk.Listbox(root, width=50, height=6)
instances_listbox.pack(pady=5)

button_frame = tk.Frame(root)
button_frame.pack(pady=5)
run_button = tk.Button(button_frame, text="Run FRPC", command=run_frpc)
run_button.pack(side=tk.LEFT, padx=5)
stop_button = tk.Button(button_frame, text="Stop Selected", command=stop_frpc)
stop_button.pack(side=tk.LEFT, padx=5)
stop_all_button = tk.Button(button_frame, text="Stop All", command=stop_all)
stop_all_button.pack(side=tk.LEFT, padx=5)
mobaXterm_button = tk.Button(button_frame, text="Open MobaXterm SSH", command=open_mobaXterm)
mobaXterm_button.pack(side=tk.LEFT, padx=5)

status_label = tk.Label(root, text="")
status_label.pack(pady=5)

output_text = scrolledtext.ScrolledText(root, width=80, height=15)
output_text.pack(pady=5)

root.mainloop()