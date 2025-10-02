import ctypes
import json
import time
from datetime import datetime
import tkinter as tk
from tkinter import ttk

# List of work programs
work_programs = ["outlook", "excel", "teams", "powerpoint", "edge", "word", "powerpoint"]

# List to store active window data
active_window_data = []

def get_active_window():
    user32 = ctypes.windll.user32
    window = user32.GetForegroundWindow()
    length = user32.GetWindowTextLengthW(window)
    buff = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(window, buff, length + 1)
    return buff.value

def update_json(new_data, filename="active_window_data.json"):
    try:
        with open(filename, "r") as f:
            existing_data = json.load(f)
    except FileNotFoundError:
        existing_data = []

    existing_data.extend(new_data)

    with open(filename, "w") as f:
        json.dump(existing_data, f, indent=4)

def calculate_totals():
    totals = {prog: 0 for prog in work_programs}
    today = datetime.now().strftime("%Y-%m-%d")
    for entry in active_window_data:
        if entry["Date"] == today:
            for prog in work_programs:
                if prog in entry["Active window"].lower():
                    totals[prog] += entry["Duration"]
    return totals

def main():
    current_window = None
    start_time = None

    def update_gui():
        nonlocal current_window, start_time
        active_window = get_active_window().lower()
        if any(prog in active_window for prog in work_programs):
            if active_window != current_window:
                if current_window:
                    end_time = datetime.now()
                    duration = (end_time - start_time).total_seconds()
                    active_window_data.append({
                        "Active window": current_window,
                        "Duration": duration,
                        "Date": start_time.strftime("%Y-%m-%d")
                    })
                    update_json(active_window_data)
                    counter_label.config(text=f"Entries Recorded: {len(active_window_data)}")
                current_window = active_window
                start_time = datetime.now()
        else:
            if current_window:
                end_time = datetime.now()
                duration = (end_time - start_time).total_seconds()
                active_window_data.append({
                    "Active window": current_window,
                    "Duration": duration,
                    "Date": start_time.strftime("%Y-%m-%d")
                })
                update_json(active_window_data)
                counter_label.config(text=f"Entries Recorded: {len(active_window_data)}")
                current_window = None
                start_time = None

        if current_window:
            elapsed_time = (datetime.now() - start_time).total_seconds()
            timer_label.config(text=f"Timer: {int(elapsed_time)} seconds")
        else:
            timer_label.config(text="Timer: 0 seconds")

        active_window_label.config(text=f"Active Window: {active_window}")
        date_label.config(text=f"Date: {datetime.now().strftime('%Y-%m-%d')}")
        time_label.config(text=f"Time: {datetime.now().strftime('%H:%M:%S')}")

        totals = calculate_totals()
        total_time = sum(totals.values())
        summary_text = f"8 hours - Total: {(8*3600 - total_time)/3600:.2f} hours\n"
        for prog, total in totals.items():
            summary_text += f"Total {prog.capitalize()}: {total/3600:.2f} hours\n"
        summary_label.config(text=summary_text)

        root.after(1000, update_gui)

    def exit_program():
        totals = calculate_totals()
        total_work_hours = sum(totals.values()) / 3600
        summary_data = [{"total work hours": total_work_hours}]
        for prog, total in totals.items():
            summary_data.append({f"Total {prog.capitalize()}": total / 3600})
        
        today_filename = f"{datetime.now().strftime('%Y-%m-%d')}.json"
        update_json(summary_data, today_filename)
        
        root.destroy()

    root = tk.Tk()
    root.title("Active Window Tracker")

    active_window_label = ttk.Label(root, text="Active Window: ", font=("Helvetica", 16))
    active_window_label.pack(pady=10)

    counter_label = ttk.Label(root, text="Entries Recorded: 0", font=("Helvetica", 16))
    counter_label.pack(pady=10)

    date_label = ttk.Label(root, text=f"Date: {datetime.now().strftime('%Y-%m-%d')}", font=("Helvetica", 16))
    date_label.pack(pady=10)

    time_label = ttk.Label(root, text=f"Time: {datetime.now().strftime('%H:%M:%S')}", font=("Helvetica", 16))
    time_label.pack(pady=10)

    timer_label = ttk.Label(root, text="Timer: 0 seconds", font=("Helvetica", 16))
    timer_label.pack(pady=10)

    summary_label = ttk.Label(root, text="Summary: ", font=("Helvetica", 16))
    summary_label.pack(pady=10)

    exit_button = ttk.Button(root, text="Exit", command=exit_program)
    exit_button.pack(pady=20)

    root.after(1000, update_gui)
    root.mainloop()

if __name__ == "__main__":
    main()
