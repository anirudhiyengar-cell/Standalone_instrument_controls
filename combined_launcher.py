#!/usr/bin/env python3  # Shebang: allows the script to be run as an executable on Unix-like systems
"""  # Start of a module-level docstring for high-level description
Combined Keithley Instrument Automation Launcher  # Brief title of the program

This script provides a single entry point to launch either the Keithley DMM6500 automation GUI or the
Keithley multi-channel power supply automation GUI.  # Clarifies purpose of script

The actual GUIs are defined in their respective modules:  # Points to source modules
- keithley_dmm_automation_main.py  # Module containing the DMM GUI class and main() function
- keithley_power_supply_automation.py  # Module containing the power-supply GUI class and main() function

The launcher displays a minimal Tkinter window with two buttons that let the user choose which
automation suite to run.  # Explains launcher interface

Selecting a button spawns a **separate subprocess** running the chosen module so the GUIs remain independent
and avoid conflicts such as multiple Tk() roots in the same interpreter.  # Rationale for subprocess
"""  # End of module-level docstring

import sys  # Provides access to the current Python interpreter and system parameters
import subprocess  # Enables launching external Python scripts as new processes
import tkinter as tk  # Imports Tkinter for GUI elements (buttons, window)
from tkinter import ttk  # Imports themed Tk widgets for modern look
from pathlib import Path  # Simplifies filesystem path handling in a cross-platform way

# ---------------------------------------------------------------------------
# Resolve absolute paths to the two target automation scripts  # Section header comment
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent  # Directory where this launcher script resides
DMM_SCRIPT = SCRIPT_DIR / 'keithley_dmm_main.py'  # Path object for DMM automation script
PS_SCRIPT = SCRIPT_DIR / 'keithley_power_supply_automation.py'  # Path object for power-supply script
SCOPE_SCRIPT = SCRIPT_DIR / 'keysight_oscilloscope_main.py'  # Path object for scope automation script
Voltage_ramp = SCRIPT_DIR / 'IMPROVED_SAFE_voltage_ramping_v2.3.py'  # Path object for voltage ramp script
# Validate that both scripts exist so we can provide helpful errors early.  # Prevent silent failure
if not DMM_SCRIPT.is_file():  # Check presence of DMM script file
    sys.exit(f'ERROR: {DMM_SCRIPT} not found – cannot launch DMM GUI')  # Abort with message
if not PS_SCRIPT.is_file():  # Check presence of power-supply script file
    sys.exit(f'ERROR: {PS_SCRIPT} not found – cannot launch power-supply GUI')  # Abort with message
if not SCOPE_SCRIPT.is_file():  # Check presence of scope script file
    sys.exit(f'ERROR: {SCOPE_SCRIPT} not found – cannot launch scope GUI')  # Abort with message
if not Voltage_ramp.is_file():  # Check presence of voltage ramp script file
    sys.exit(f'ERROR: {Voltage_ramp} not found – cannot launch voltage ramp GUI')  # Abort with message

# ---------------------------------------------------------------------------
# Helper function to spawn a new Python subprocess for a target script         # Section header
# ---------------------------------------------------------------------------
def launch_script(script_path: Path) -> None:  # Function definition with type hint
    """Launch the given automation script in a separate process.  # Docstring
    Args:                                                                      # Parameter list description
        script_path: Fully-qualified Path to the Python file to execute.       # Explains 'script_path'
    """  # End of docstring
    # Build the command: [current_python_interpreter, target_script]          # Inline comment
    cmd = [sys.executable, str(script_path)]  # Uses same Python executable that ran this launcher
    # Run the command without blocking the launcher GUI                       # Explain subprocess call
    subprocess.Popen(cmd)  # Fire-and-forget; user can launch multiple GUIs if desired

# ---------------------------------------------------------------------------
# Build the simple Tkinter launcher window                                    # Section header
# ---------------------------------------------------------------------------
root = tk.Tk()  # Create the main (and only) Tk root window for the launcher
root.title('Keithley Instrument Automation Launcher')  # Window title displayed in title bar
root.geometry('800x180')
root.configure(bg="#99565f")  # Set reasonable default size (width x height in pixels)
root.resizable(False, False)  # Fix window size so layout stays simple

# Configure a single column & row to expand so widgets center nicely          # Grid layout config
root.columnconfigure(0, weight=1)  # Allow column 0 to expand horizontally
root.rowconfigure(0, weight=1)  # Allow row 0 to expand vertically

# Create a frame to hold the buttons so we can add uniform padding.           # Layout container
#frame = ttk.Frame(root, bg='#00008b', padding=20)
frame = tk.Frame(root, bg="#BBB7A0", padx=20, pady=20)  # 20-pixel padding on all sides
frame.grid(sticky='nsew')  # Attach frame to root and have it expand

# Style configuration for a consistent professional look                     # Style setup
style = ttk.Style()  # Access the style manager
style.theme_use('clam')  # Use the modern 'clam' theme for Tkinter widgets
#style.configure('Launch.TButton', font=('Arial', 11, 'bold'), padding=10)  # Custom button style
style.configure('Launch.TButton',
                font=('Bookman Old Style', 11, 'bold'),
                foreground='white')

# Colors per state
style.map('Launch.TButton',
          background=[('pressed', '#000000'),
                      ('active',  '#000000'),
                      ('!disabled', "#C03E3E")],     # normal state
          foreground=[('disabled', '#000000'),
                      ('!disabled', 'white')])
# ---------------------------------------------------------------------------
# Button callbacks: simply call launch_script with the proper target          # Section header
# ---------------------------------------------------------------------------
def launch_dmm():  # Callback for DMM button
    launch_script(DMM_SCRIPT)  # Spawn DMM automation script as subprocess

def launch_ps():  # Callback for power-supply button
    launch_script(PS_SCRIPT)  # Spawn power-supply automation script

def launch_scope():  # Callback for scope button
    launch_script(SCOPE_SCRIPT)  # Spawn scope automation script

def launch_voltage_ramp():  # Callback for voltage ramp button
    launch_script(Voltage_ramp)  # Spawn voltage ramp automation script

# ---------------------------------------------------------------------------
# Create and place the launch buttons                                         # Section header
# ---------------------------------------------------------------------------
dmm_btn = ttk.Button(frame, text='Launch DMM Automation', style='Launch.TButton', command=launch_dmm)  # DMM button
dmm_btn.grid(row=0, column=0, sticky='nsew', pady=(0, 10))  # Top button with padding below

ps_btn = ttk.Button(frame, text='Launch Power Supply Automation', style='Launch.TButton', command=launch_ps)  # PS button
ps_btn.grid(row=0, column=1, sticky='nsew', pady=(0, 10))  # Second button directly under the first

scope_btn = ttk.Button(frame, text='Launch Scope Automation', style='Launch.TButton', command=launch_scope)  # Scope button
scope_btn.grid(row=0, column=2, sticky='nsew', pady=(0, 10))  # Third button directly under the second

voltage_ramp_btn = ttk.Button(frame, text='Launch Voltage Ramp Automation', style='Launch.TButton', command=launch_voltage_ramp)  # Voltage ramp button
voltage_ramp_btn.grid(row=1, column=1, sticky='nsew', pady=(0, 10))  # Fourth button directly under the third

# Add a quit button for convenience                                           # Optional exit button
quit_btn = ttk.Button(frame, text='Exit', command=root.destroy)  # Close launcher window
quit_btn.grid(row=2, column=1, sticky='ew', pady=(10, 0))  # Padding above quit button

# Give equal weight to each row so buttons distribute spacing evenly          # Row weight config
for i in range(4):  # Iterate over rows 0-2
    frame.rowconfigure(i, weight=1)  # Allow each row to expand equally

# ---------------------------------------------------------------------------
# Start the Tkinter event loop                                                # Final step
# ---------------------------------------------------------------------------
if __name__ == '__main__':  # Standard idiom to allow import without auto-run
    root.mainloop()  # Enter the Tk event loop to display the launcher window