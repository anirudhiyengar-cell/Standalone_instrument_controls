#!/usr/bin/env python3
"""
═══════════════════════════════════════════════════════════════════════════════
    UNIFIED LABORATORY AUTOMATION SYSTEM - COMPLETE READY-TO-USE VERSION
═══════════════════════════════════════════════════════════════════════════════

Professional-grade unified automation system integrating:
    • Keithley 2230G-30-1 Multi-Channel Power Supply
    • Keithley DMM6500 Digital Multimeter  
    • Keysight DSOX6004A Digital Oscilloscope

Author: Professional Test Automation Engineer
Version: 3.0 - Complete Single File
Date: October 23, 2025

IMPORTANT: This file uses your existing instrument wrappers from:
    - instrument_control/keithley_power_supply.py (KeithleyPowerSupply)
    - instrument_control/keithley_dmm.py (KeithleyDMM6500)
    - instrument_control/keysight_oscilloscope.py (KeysightDSOX6004A)

No file renaming required - works with your exact existing setup!

═══════════════════════════════════════════════════════════════════════════════
"""

import sys
import os
import logging
import threading
import queue
import time
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, List, Tuple, Any
from enum import Enum

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext

try:
    import numpy as np
    import pandas as pd
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
except ImportError as e:
    print(f"ERROR: Required package missing: {e}")
    print("Install: pip install numpy pandas matplotlib")
    sys.exit(1)

# Import instrument wrappers - using YOUR existing file names!
try:
    from instrument_control.keithley_power_supply import KeithleyPowerSupply
    from instrument_control.keithley_dmm import KeithleyDMM6500
    from instrument_control.keysight_oscilloscope import KeysightDSOX6004A
except ImportError as e:
    print(f"ERROR: Cannot import instrument control wrappers: {e}")
    print("Expected directory structure:")
    print("  instrument_control/")
    print("    ├── keithley_power_supply.py")
    print("    ├── keithley_dmm.py")
    print("    └── keysight_oscilloscope.py")
    sys.exit(1)

# Optional import of advanced oscilloscope data acquisition utilities
try:
    from instrument_control.keysight_oscilloscope import KeysightDSOX6004A as OscilloscopeDataAcquisition
except Exception:
    OscilloscopeDataAcquisition = None

# Logging configuration
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(f'automation_{datetime.now().strftime("%Y%m%d")}.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger('UnifiedAutomation')

# Global constants
class WaveformType(Enum):
    SINE = "Sine"
    RAMP = "Ramp"
    SQUARE = "Square"
    TRIANGLE = "Triangle"

COLORS = {
    'bg_main': '#f0f0f0',
    'bg_frame': '#ffffff',
    'bg_emergency': '#ffebee',
    'bg_log': '#1e1e1e',
    'fg_log': '#00ff41',
    'btn_connect': '#4caf50',
    'btn_disconnect': '#f44336',
    'btn_emergency': '#d32f2f',
    'btn_primary': '#2196f3',
    'btn_success': '#4caf50',
    'btn_warning': '#ff9800',
    'btn_info': '#00bcd4',
    'status_connected': '#4caf50',
    'status_disconnected': '#f44336',
    'status_safe': '#4caf50',
    'status_active': '#ff9800',
}


class UnifiedDataHandler:
    """Centralized data management for all instruments"""
    
    def __init__(self):
        self.psu_data: List[Dict] = []
        self.dmm_data: List[Dict] = []
        self.scope_data: List[Dict] = []
        self.ramping_data: List[Dict] = []
        
        self.data_folder = Path.home() / "Desktop" / "UnifiedAutomation" / "Data"
        self.graph_folder = Path.home() / "Desktop" / "UnifiedAutomation" / "Graphs"
        self.screenshot_folder = Path.home() / "Desktop" / "UnifiedAutomation" / "Screenshots"
        
        for folder in [self.data_folder, self.graph_folder, self.screenshot_folder]:
            folder.mkdir(parents=True, exist_ok=True)
        
        logger.info(f"Data handler initialized")
    
    def add_psu_measurement(self, channel: int, voltage: float, current: float, power: float):
        self.psu_data.append({
            'timestamp': datetime.now(),
            'channel': channel,
            'voltage': voltage,
            'current': current,
            'power': power
        })
    
    def add_dmm_measurement(self, function: str, value: float, unit: str):
        self.dmm_data.append({
            'timestamp': datetime.now(),
            'function': function,
            'value': value,
            'unit': unit
        })
    
    def add_scope_data(self, channel: int, time_data: np.ndarray, voltage_data: np.ndarray):
        self.scope_data.append({
            'timestamp': datetime.now(),
            'channel': channel,
            'time': time_data,
            'voltage': voltage_data
        })
    
    def add_ramping_data(self, cycle: int, point: int, set_voltage: float, 
                        meas_voltage: float, meas_current: float):
        self.ramping_data.append({
            'timestamp': datetime.now(),
            'cycle': cycle,
            'point': point,
            'set_voltage': set_voltage,
            'measured_voltage': meas_voltage,
            'measured_current': meas_current
        })
    
    def export_psu_data(self, filename: Optional[str] = None) -> str:
        if not self.psu_data:
            raise ValueError("No PSU data to export")
        
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = self.data_folder / f"PSU_Data_{timestamp}.csv"
        
        df = pd.DataFrame(self.psu_data)
        df.to_csv(filename, index=False)
        logger.info(f"Exported {len(self.psu_data)} PSU measurements")
        return str(filename)
    
    def export_dmm_data(self, filename: Optional[str] = None) -> str:
        if not self.dmm_data:
            raise ValueError("No DMM data to export")
        
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = self.data_folder / f"DMM_Data_{timestamp}.csv"
        
        df = pd.DataFrame(self.dmm_data)
        df.to_csv(filename, index=False)
        logger.info(f"Exported {len(self.dmm_data)} DMM measurements")
        return str(filename)
    
    def export_ramping_data(self, filename: Optional[str] = None) -> str:
        if not self.ramping_data:
            raise ValueError("No ramping data to export")
        
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = self.data_folder / f"Ramping_Test_{timestamp}.csv"
        
        df = pd.DataFrame(self.ramping_data)
        df.to_csv(filename, index=False)
        logger.info(f"Exported {len(self.ramping_data)} ramping data points")
        return str(filename)
    
    def generate_ramping_graph(self) -> str:
        if not self.ramping_data:
            raise ValueError("No ramping data to plot")
        
        df = pd.DataFrame(self.ramping_data)
        
        fig, axes = plt.subplots(2, 2, figsize=(14, 10))
        
        # Plot 1: Set vs Measured Voltage
        axes[0, 0].scatter(df['set_voltage'], df['measured_voltage'], alpha=0.6)
        axes[0, 0].plot([df['set_voltage'].min(), df['set_voltage'].max()],
                       [df['set_voltage'].min(), df['set_voltage'].max()],
                       'r--', label='Ideal')
        axes[0, 0].set_xlabel('Set Voltage (V)')
        axes[0, 0].set_ylabel('Measured Voltage (V)')
        axes[0, 0].set_title('Voltage Accuracy')
        axes[0, 0].legend()
        axes[0, 0].grid(True, alpha=0.3)
        
        # Plot 2: Voltage over time
        axes[0, 1].plot(df.index, df['measured_voltage'], 'b-', alpha=0.7)
        axes[0, 1].set_xlabel('Data Point')
        axes[0, 1].set_ylabel('Measured Voltage (V)')
        axes[0, 1].set_title('Voltage Ramping Profile')
        axes[0, 1].grid(True, alpha=0.3)
        
        # Plot 3: Current over time
        axes[1, 0].plot(df.index, df['measured_current'], 'g-', alpha=0.7)
        axes[1, 0].set_xlabel('Data Point')
        axes[1, 0].set_ylabel('Measured Current (A)')
        axes[1, 0].set_title('Current During Ramping')
        axes[1, 0].grid(True, alpha=0.3)
        
        # Plot 4: Voltage error by cycle
        df['voltage_error'] = df['measured_voltage'] - df['set_voltage']
        for cycle in df['cycle'].unique():
            cycle_data = df[df['cycle'] == cycle]
            axes[1, 1].plot(cycle_data['point'], cycle_data['voltage_error'], 
                          marker='o', label=f'Cycle {cycle}')
        axes[1, 1].set_xlabel('Point in Cycle')
        axes[1, 1].set_ylabel('Voltage Error (V)')
        axes[1, 1].set_title('Voltage Error by Cycle')
        axes[1, 1].legend()
        axes[1, 1].grid(True, alpha=0.3)
        
        plt.tight_layout()
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = self.graph_folder / f"Ramping_Analysis_{timestamp}.png"
        plt.savefig(filename, dpi=300, bbox_inches='tight')
        plt.close()
        
        logger.info(f"Generated ramping graph: {filename}")
        return str(filename)
    
    def generate_psu_graph(self, save_path: Optional[str] = None) -> str:
        """Generate professional PSU measurement graph with user-selectable save location"""
        if not self.psu_data:
            raise ValueError("No PSU data to plot")

        df = pd.DataFrame(self.psu_data)
        start_time = df['timestamp'].min()
        df['time_seconds'] = (df['timestamp'] - start_time).dt.total_seconds()

        fig, axes = plt.subplots(3, 1, figsize=(14, 10), sharex=True)
        plt.style.use('seaborn-v0_8-darkgrid')

        colors = {'1': '#1f77b4', '2': '#ff7f0e', '3': '#2ca02c'}

        for channel in df['channel'].unique():
            ch_data = df[df['channel'] == channel]
            axes[0].plot(ch_data['time_seconds'], ch_data['voltage'], label=f'Channel {channel}',
                        color=colors.get(str(channel), '#d62728'), linewidth=2, marker='o', markersize=4, alpha=0.8)
        axes[0].set_ylabel('Voltage (V)', fontsize=12, fontweight='bold')
        axes[0].set_title('PSU Measurements Over Time', fontsize=14, fontweight='bold')
        axes[0].legend(loc='best')
        axes[0].grid(True, ls='--', alpha=0.3)

        for channel in df['channel'].unique():
            ch_data = df[df['channel'] == channel]
            axes[1].plot(ch_data['time_seconds'], ch_data['current'], label=f'Channel {channel}',
                        color=colors.get(str(channel), '#d62728'), linewidth=2, marker='s', markersize=4, alpha=0.8)
        axes[1].set_ylabel('Current (A)', fontsize=12, fontweight='bold')
        axes[1].legend(loc='best')
        axes[1].grid(True, ls='--', alpha=0.3)

        for channel in df['channel'].unique():
            ch_data = df[df['channel'] == channel]
            axes[2].plot(ch_data['time_seconds'], ch_data['power'], label=f'Channel {channel}',
                        color=colors.get(str(channel), '#d62728'), linewidth=2, marker='^', markersize=4, alpha=0.8)
        axes[2].set_xlabel('Time (seconds)', fontsize=12, fontweight='bold')
        axes[2].set_ylabel('Power (W)', fontsize=12, fontweight='bold')
        axes[2].legend(loc='best')
        axes[2].grid(True, ls='--', alpha=0.3)

        plt.tight_layout()

        if save_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            save_path = self.graph_folder / f"PSU_Graph_{timestamp}.png"
        else:
            save_path = Path(save_path)
            save_path.parent.mkdir(parents=True, exist_ok=True)

        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close(fig)
        logger.info(f"Generated PSU graph: {save_path}")
        return str(save_path)

    def generate_dmm_graph(self, save_path: Optional[str] = None) -> str:
        """Generate professional DMM measurement graph with user-selectable save location"""
        if not self.dmm_data:
            raise ValueError("No DMM data to plot")

        df = pd.DataFrame(self.dmm_data)
        start_time = df['timestamp'].min()
        df['time_seconds'] = (df['timestamp'] - start_time).dt.total_seconds()

        functions = df['function'].unique()
        n_functions = len(functions)

        fig, axes = plt.subplots(n_functions, 1, figsize=(14, 4*n_functions), sharex=True, squeeze=False)
        plt.style.use('seaborn-v0_8-darkgrid')
        axes = axes.flatten()
        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']

        for idx, func in enumerate(functions):
            func_data = df[df['function'] == func]
            unit = func_data['unit'].iloc[0] if len(func_data) > 0 else ''

            axes[idx].plot(func_data['time_seconds'], func_data['value'], label=f'{func} ({unit})',
                          color=colors[idx % len(colors)], linewidth=2, marker='o', markersize=5, alpha=0.8)
            axes[idx].set_ylabel(f'{func} ({unit})', fontsize=12, fontweight='bold')
            axes[idx].grid(True, ls='--', alpha=0.3)
            axes[idx].legend(loc='best')

            stats = func_data['value'].describe()
            stats_text = f"Mean: {stats['mean']:.6e}\nStd: {stats['std']:.6e}\nMin: {stats['min']:.6e}\nMax: {stats['max']:.6e}"
            axes[idx].text(0.98, 0.98, stats_text, transform=axes[idx].transAxes,
                          verticalalignment='top', horizontalalignment='right', fontsize=9,
                          bbox=dict(boxstyle='round', facecolor='white', alpha=0.8))

            if idx == 0:
                axes[idx].set_title('DMM Measurements Over Time', fontsize=14, fontweight='bold')

        axes[-1].set_xlabel('Time (seconds)', fontsize=12, fontweight='bold')
        plt.tight_layout()

        if save_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            save_path = self.graph_folder / f"DMM_Graph_{timestamp}.png"
        else:
            save_path = Path(save_path)
            save_path.parent.mkdir(parents=True, exist_ok=True)

        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close(fig)
        logger.info(f"Generated DMM graph: {save_path}")
        return str(save_path)

    def clear_all_data(self):
        self.psu_data.clear()
        self.dmm_data.clear()
        self.scope_data.clear()
        self.ramping_data.clear()
        logger.info("All data cleared")


class UnifiedAutomationGUI:
    """Main GUI class for unified automation system"""
    
    def __init__(self):
        logger.info("="*80)
        logger.info("UNIFIED AUTOMATION - SESSION STARTED")
        logger.info("="*80)
        
        self.root = tk.Tk()
        self.root.title("Unified Laboratory Automation System - Professional Edition")
        # Start maximized for better user experience
        try:
            self.root.state('zoomed')  # Windows
        except Exception:
            self.root.attributes('-fullscreen', False)
            self.root.geometry("1400x900")
        self.root.configure(bg=COLORS['bg_main'])
        
        # Instruments
        self.psu: Optional[KeithleyPowerSupply] = None
        self.dmm: Optional[KeithleyDMM6500] = None
        self.scope: Optional[KeysightDSOX6004A] = None
        
        self.instruments_connected = {
            'psu': False,
            'dmm': False,
            'scope': False
        }

        # Scope data helpers
        self.scope_data_acq = None
        self.scope_last_data = None
        
        self.data_handler = UnifiedDataHandler()
        self.status_queue = queue.Queue()
        
        self.emergency_stop_active = False
        self.ramping_active = False
        self.ramping_thread: Optional[threading.Thread] = None
        self.auto_measurement_active = False
        self.auto_measurement_thread: Optional[threading.Thread] = None
        self.dmm_continuous_active = False
        
        self._setup_styles()
        self._create_widgets()
        self._start_status_checker()
        
        logger.info("Unified Automation GUI initialized")
    
    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'), 
                       foreground='#1976d2', background=COLORS['bg_main'])
        style.configure('Subtitle.TLabel', font=('Arial', 12, 'bold'), 
                       foreground='#424242', background=COLORS['bg_frame'])

    def _create_menubar(self):
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label='Export PSU Data', command=lambda: self._export_data('psu'))
        file_menu.add_command(label='Export DMM Data', command=lambda: self._export_data('dmm'))
        file_menu.add_separator()
        file_menu.add_command(label='Exit', command=self._on_closing)
        menubar.add_cascade(label='File', menu=file_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label='Toggle Activity Log', command=self._toggle_activity_log)
        menubar.add_cascade(label='View', menu=view_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label='About', command=lambda: messagebox.showinfo('About', 'Unified Lab Automation - Professional'))
        menubar.add_cascade(label='Help', menu=help_menu)

        self.root.config(menu=menubar)

    def _toggle_activity_log(self):
        try:
            parent = self.activity_log.master
            if parent.winfo_ismapped():
                parent.pack_forget()
                if hasattr(self, 'status_var'):
                    self.status_var.set('Activity log hidden')
            else:
                parent.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
                if hasattr(self, 'status_var'):
                    self.status_var.set('Activity log visible')
        except Exception:
            pass
    
    def _create_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        title_label = ttk.Label(
            main_frame,
            text="UNIFIED LABORATORY AUTOMATION SYSTEM",
            style='Title.TLabel'
        )
        title_label.pack(pady=(0, 10))
        
        self._create_emergency_controls(main_frame)
        self._create_connection_panel(main_frame)
        self._create_global_file_prefs(main_frame)
        # Create application menubar
        self._create_menubar()
        
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self._create_ramping_tab()
        self._create_psu_tab()
        self._create_dmm_tab()
        self._create_scope_tab()
        self._create_data_export_tab()
        
        self._create_activity_log(main_frame)

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor='w',
                              bg='#f5f5f5', font=('Segoe UI', 9))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _create_emergency_controls(self, parent):
        emergency_frame = tk.Frame(parent, bg=COLORS['bg_emergency'], 
                                  relief=tk.RAISED, borderwidth=3)
        emergency_frame.pack(fill=tk.X, pady=(0, 10))
        
        title = tk.Label(
            emergency_frame,
            text="EMERGENCY CONTROLS",
            font=('Arial', 14, 'bold'),
            bg=COLORS['bg_emergency'],
            fg=COLORS['btn_emergency']
        )
        title.pack(pady=5)
        
        btn_frame = tk.Frame(emergency_frame, bg=COLORS['bg_emergency'])
        btn_frame.pack(pady=5)
        
        self.emergency_btn = tk.Button(
            btn_frame,
            text="EMERGENCY STOP\nImmediate PSU Shutdown",
            font=('Arial', 12, 'bold'),
            bg=COLORS['btn_emergency'],
            fg='white',
            width=30,
            height=3,
            command=self._emergency_stop,
            relief=tk.RAISED,
            borderwidth=3
        )
        self.emergency_btn.pack(side=tk.LEFT, padx=10)
        
        status_frame = tk.Frame(btn_frame, bg=COLORS['bg_emergency'])
        status_frame.pack(side=tk.LEFT, padx=20)
        
        tk.Label(
            status_frame,
            text="System Status:",
            font=('Arial', 11, 'bold'),
            bg=COLORS['bg_emergency']
        ).pack()
        
        self.system_status_label = tk.Label(
            status_frame,
            text="System Safe - Ready for operation",
            font=('Arial', 11, 'bold'),
            bg=COLORS['bg_emergency'],
            fg=COLORS['status_safe']
        )
        self.system_status_label.pack()
    
    def _create_connection_panel(self, parent):
        conn_frame = tk.LabelFrame(
            parent,
            text="Instrument Connections",
            font=('Arial', 11, 'bold'),
            bg=COLORS['bg_frame'],
            relief=tk.GROOVE,
            borderwidth=2
        )
        conn_frame.pack(fill=tk.X, pady=(0, 10))
        
        # PSU
        psu_frame = tk.Frame(conn_frame, bg=COLORS['bg_frame'])
        psu_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(
            psu_frame,
            text="Power Supply VISA:",
            font=('Arial', 10, 'bold'),
            bg=COLORS['bg_frame'],
            width=20,
            anchor='w'
        ).pack(side=tk.LEFT)
        
        self.psu_visa_entry = tk.Entry(psu_frame, font=('Courier', 10), width=50)
        self.psu_visa_entry.insert(0, "USB0::0x05E6::0x2230::805224014806770001::INSTR")
        self.psu_visa_entry.pack(side=tk.LEFT, padx=5)
        
        self.psu_connect_btn = tk.Button(
            psu_frame,
            text="Connect",
            font=('Arial', 10, 'bold'),
            bg=COLORS['btn_connect'],
            fg='white',
            width=12,
            command=lambda: self._connect_instrument('psu')
        )
        self.psu_connect_btn.pack(side=tk.LEFT, padx=2)
        
        self.psu_disconnect_btn = tk.Button(
            psu_frame,
            text="Disconnect",
            font=('Arial', 10, 'bold'),
            bg=COLORS['btn_disconnect'],
            fg='white',
            width=12,
            command=lambda: self._disconnect_instrument('psu'),
            state=tk.DISABLED
        )
        self.psu_disconnect_btn.pack(side=tk.LEFT, padx=2)
        
        self.psu_status_label = tk.Label(
            psu_frame,
            text="Disconnected",
            font=('Arial', 10, 'bold'),
            fg=COLORS['status_disconnected'],
            bg=COLORS['bg_frame']
        )
        self.psu_status_label.pack(side=tk.LEFT, padx=10)
        
        # DMM
        dmm_frame = tk.Frame(conn_frame, bg=COLORS['bg_frame'])
        dmm_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(
            dmm_frame,
            text="DMM VISA:",
            font=('Arial', 10, 'bold'),
            bg=COLORS['bg_frame'],
            width=20,
            anchor='w'
        ).pack(side=tk.LEFT)
        
        self.dmm_visa_entry = tk.Entry(dmm_frame, font=('Courier', 10), width=50)
        self.dmm_visa_entry.insert(0, "USB0::0x05E6::0x6500::04561287::INSTR")
        self.dmm_visa_entry.pack(side=tk.LEFT, padx=5)
        
        self.dmm_connect_btn = tk.Button(
            dmm_frame,
            text="Connect",
            font=('Arial', 10, 'bold'),
            bg=COLORS['btn_connect'],
            fg='white',
            width=12,
            command=lambda: self._connect_instrument('dmm')
        )
        self.dmm_connect_btn.pack(side=tk.LEFT, padx=2)
        
        self.dmm_disconnect_btn = tk.Button(
            dmm_frame,
            text="Disconnect",
            font=('Arial', 10, 'bold'),
            bg=COLORS['btn_disconnect'],
            fg='white',
            width=12,
            command=lambda: self._disconnect_instrument('dmm'),
            state=tk.DISABLED
        )
        self.dmm_disconnect_btn.pack(side=tk.LEFT, padx=2)
        
        self.dmm_status_label = tk.Label(
            dmm_frame,
            text="Disconnected",
            font=('Arial', 10, 'bold'),
            fg=COLORS['status_disconnected'],
            bg=COLORS['bg_frame']
        )
        self.dmm_status_label.pack(side=tk.LEFT, padx=10)
        
        # Scope
        scope_frame = tk.Frame(conn_frame, bg=COLORS['bg_frame'])
        scope_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(
            scope_frame,
            text="Oscilloscope VISA:",
            font=('Arial', 10, 'bold'),
            bg=COLORS['bg_frame'],
            width=20,
            anchor='w'
        ).pack(side=tk.LEFT)
        
        self.scope_visa_entry = tk.Entry(scope_frame, font=('Courier', 10), width=50)
        self.scope_visa_entry.insert(0, "USB0::0x0957::0x1780::MY65220169::INSTR")
        self.scope_visa_entry.pack(side=tk.LEFT, padx=5)
        
        self.scope_connect_btn = tk.Button(
            scope_frame,
            text="Connect",
            font=('Arial', 10, 'bold'),
            bg=COLORS['btn_connect'],
            fg='white',
            width=12,
            command=lambda: self._connect_instrument('scope')
        )
        self.scope_connect_btn.pack(side=tk.LEFT, padx=2)
        
        self.scope_disconnect_btn = tk.Button(
            scope_frame,
            text="Disconnect",
            font=('Arial', 10, 'bold'),
            bg=COLORS['btn_disconnect'],
            fg='white',
            width=12,
            command=lambda: self._disconnect_instrument('scope'),
            state=tk.DISABLED
        )
        self.scope_disconnect_btn.pack(side=tk.LEFT, padx=2)
        
        self.scope_status_label = tk.Label(
            scope_frame,
            text="Disconnected",
            font=('Arial', 10, 'bold'),
            fg=COLORS['status_disconnected'],
            bg=COLORS['bg_frame']
        )
        self.scope_status_label.pack(side=tk.LEFT, padx=10)
        
        # Bulk operations
        bulk_frame = tk.Frame(conn_frame, bg=COLORS['bg_frame'])
        bulk_frame.pack(pady=10)
        
        tk.Button(
            bulk_frame,
            text="Connect All Instruments",
            font=('Arial', 10, 'bold'),
            bg=COLORS['btn_primary'],
            fg='white',
            width=25,
            command=self._connect_all_instruments
        ).pack(side=tk.LEFT, padx=5)

    def _create_global_file_prefs(self, parent):
        prefs = tk.LabelFrame(
            parent,
            text="Global File Preferences",
            font=('Arial', 11, 'bold'),
            bg=COLORS['bg_frame'],
            relief=tk.GROOVE,
            borderwidth=2
        )
        prefs.pack(fill=tk.X, pady=(0, 10))

        def add_row(label, get_path, set_type):
            row = tk.Frame(prefs, bg=COLORS['bg_frame'])
            row.pack(fill=tk.X, padx=10, pady=4)
            tk.Label(row, text=label, font=('Arial', 10, 'bold'), bg=COLORS['bg_frame'], width=15, anchor='w').pack(side=tk.LEFT)
            entry = tk.Entry(row, font=('Arial', 9))
            entry.insert(0, str(get_path()))
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            tk.Button(row, text="...", width=3, command=lambda: self._browse_global_folder(set_type, entry)).pack(side=tk.LEFT)
            return entry

        self.global_data_entry = add_row("Data:", lambda: self.data_handler.data_folder, 'data')
        self.global_graph_entry = add_row("Graphs:", lambda: self.data_handler.graph_folder, 'graphs')
        self.global_ss_entry = add_row("Screenshots:", lambda: self.data_handler.screenshot_folder, 'screenshots')

    def _browse_global_folder(self, folder_type: str, entry_widget: tk.Entry):
        try:
            current = entry_widget.get().strip() or str(Path.home())
            folder = filedialog.askdirectory(initialdir=current, title=f"Select {folder_type.title()} Folder")
            if folder:
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, folder)
                if folder_type == 'data':
                    self.data_handler.data_folder = Path(folder)
                    # Sync scope tab field if present
                    if hasattr(self, 'scope_data_path'):
                        self.scope_data_path.delete(0, tk.END)
                        self.scope_data_path.insert(0, folder)
                elif folder_type == 'graphs':
                    self.data_handler.graph_folder = Path(folder)
                    if hasattr(self, 'scope_graph_path'):
                        self.scope_graph_path.delete(0, tk.END)
                        self.scope_graph_path.insert(0, folder)
                elif folder_type == 'screenshots':
                    self.data_handler.screenshot_folder = Path(folder)
                    if hasattr(self, 'scope_screenshot_path'):
                        self.scope_screenshot_path.delete(0, tk.END)
                        self.scope_screenshot_path.insert(0, folder)
                Path(folder).mkdir(parents=True, exist_ok=True)
                self.log_message(f"Global {folder_type} folder updated: {folder}", 'SUCCESS')
        except Exception as e:
            self.log_message(f"Global folder selection error: {e}", 'ERROR')

    def _create_scrollable_frame(self, parent, title=None) -> tuple[tk.Frame, tk.Frame]:
        """Create a standardized scrollable frame with proper layout configuration
        Returns: (container_frame, content_frame)
        """
        # Create main container
        if title:
            container = tk.LabelFrame(parent, text=title, font=('Arial', 11, 'bold'), 
                                    bg=COLORS['bg_frame'])
        else:
            container = tk.Frame(parent, bg=COLORS['bg_frame'])
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        container.grid_columnconfigure(0, weight=1)
        
        # Create canvas with scrollbar
        canvas = tk.Canvas(container, bg=COLORS['bg_frame'])
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        content = tk.Frame(canvas, bg=COLORS['bg_frame'])
        
        # Configure scrolling
        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=content, anchor="nw", width=canvas.winfo_width())
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas.find_withtag('all')[0], width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Layout with grid
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Configure grid weights
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        # Enable mouse wheel scrolling
        self._enable_mousewheel(canvas, content)
        
        return container, content

    def _create_labeled_frame(self, parent, title, padx=10, pady=5):
        """Create a standardized labeled frame with proper layout configuration"""
        frame = tk.LabelFrame(parent, text=title, font=('Arial', 11, 'bold'), 
                            bg=COLORS['bg_frame'])
        frame.pack(fill=tk.BOTH, expand=True, padx=padx, pady=pady)
        frame.grid_columnconfigure(0, weight=1)
        return frame

    def _create_button_grid(self, parent, buttons, columns=4):
        """Create a standardized grid of buttons
        buttons: list of (text, command, color) tuples
        """
        frame = tk.Frame(parent, bg=COLORS['bg_frame'])
        frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Configure grid columns
        for i in range(columns):
            frame.grid_columnconfigure(i, weight=1)
        
        # Create buttons
        for i, (text, command, color) in enumerate(buttons):
            row, col = divmod(i, columns)
            tk.Button(frame, text=text, command=command,
                     font=('Arial', 10, 'bold'), bg=color, fg='white',
                     width=15).grid(row=row, column=col, padx=2, pady=2, sticky='ew')
        
        return frame

    def _enable_mousewheel(self, canvas: tk.Canvas, widget: tk.Widget):
        # Windows and macOS use <MouseWheel>, Linux uses Button-4/5
        def bind_all():
            widget.bind_all('<MouseWheel>', lambda e: self._on_mousewheel(canvas, e))
            widget.bind_all('<Shift-MouseWheel>', lambda e: self._on_mousewheel(canvas, e))
            widget.bind_all('<Button-4>', lambda e: self._on_mousewheel_linux(canvas, -1))
            widget.bind_all('<Button-5>', lambda e: self._on_mousewheel_linux(canvas, 1))
        def unbind_all():
            widget.unbind_all('<MouseWheel>')
            widget.unbind_all('<Shift-MouseWheel>')
            widget.unbind_all('<Button-4>')
            widget.unbind_all('<Button-5>')
        
        widget.bind('<Enter>', lambda e: bind_all())
        widget.bind('<Leave>', lambda e: unbind_all())

    def _on_mousewheel(self, canvas: tk.Canvas, event):
        # On Windows, event.delta is multiples of 120; on macOS it's small values
        delta = -1 if event.delta > 0 else 1
        canvas.yview_scroll(delta, 'units')

    def _on_mousewheel_linux(self, canvas: tk.Canvas, direction: int):
        canvas.yview_scroll(direction, 'units')
    
    def _create_ramping_tab(self):
        ramping_frame = tk.Frame(self.notebook, bg=COLORS['bg_frame'])
        self.notebook.add(ramping_frame, text="Voltage Ramping")
        
        # Create scrollable container that fills the width
        container = tk.Frame(ramping_frame)
        container.pack(fill=tk.BOTH, expand=True)
        container.grid_columnconfigure(0, weight=1)  # Make column expandable
        
        canvas = tk.Canvas(container, bg=COLORS['bg_frame'])
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=COLORS['bg_frame'])
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=canvas.winfo_width())
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas.find_withtag('all')[0], width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.grid(row=0, column=0, sticky="nsew")  # Use grid for better control
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Configure container grid
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # Enable mouse wheel scrolling
        self._enable_mousewheel(canvas, scrollable_frame)
        
        title = tk.Label(
            scrollable_frame,
            text="Safe Voltage Ramping Configuration",
            font=('Arial', 14, 'bold'),
            bg=COLORS['bg_frame'],
            fg='#1976d2'
        )
        title.pack(pady=10)
        
        # Create main content area that fills available width
        content_frame = tk.Frame(scrollable_frame, bg=COLORS['bg_frame'])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10)
        content_frame.grid_columnconfigure(0, weight=1)  # Make content expandable
        
        config_frame = tk.LabelFrame(
            content_frame,
            text="Ramping Parameters",
            font=('Arial', 11, 'bold'),
            bg=COLORS['bg_frame']
        )
        config_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        config_frame.grid_columnconfigure(0, weight=1)  # Make frame content expandable
        
        # Create grid layout for parameters
        params_frame = tk.Frame(config_frame, bg=COLORS['bg_frame'])
        params_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        params_frame.grid_columnconfigure(1, weight=1)  # Make second column expandable
        params_frame.grid_columnconfigure(3, weight=1)  # Make fourth column expandable
        
        # Left column parameters
        tk.Label(params_frame, text="Target Voltage:", font=('Arial', 10, 'bold'), bg=COLORS['bg_frame']).grid(row=0, column=0, sticky='w', padx=5, pady=3)
        voltage_entry = tk.Entry(params_frame, font=('Arial', 10), width=10)
        voltage_entry.insert(0, "3.00")
        voltage_entry.grid(row=0, column=1, sticky='w', padx=5)
        tk.Label(params_frame, text="V (MAX: 4.0V)", font=('Arial', 9, 'italic'), fg='#666666', bg=COLORS['bg_frame']).grid(row=0, column=2, sticky='w', padx=5)
        
        tk.Label(params_frame, text="Points per Cycle:", font=('Arial', 10, 'bold'), bg=COLORS['bg_frame']).grid(row=1, column=0, sticky='w', padx=5, pady=3)
        points_entry = tk.Entry(params_frame, font=('Arial', 10), width=10)
        points_entry.insert(0, "50")
        points_entry.grid(row=1, column=1, sticky='w', padx=5)
        tk.Label(params_frame, text="(20-100)", font=('Arial', 9, 'italic'), fg='#666666', bg=COLORS['bg_frame']).grid(row=1, column=2, sticky='w', padx=5)
        
        # Right column parameters
        tk.Label(params_frame, text="Cycles:", font=('Arial', 10, 'bold'), bg=COLORS['bg_frame']).grid(row=0, column=3, sticky='w', padx=5, pady=3)
        cycles_entry = tk.Entry(params_frame, font=('Arial', 10), width=10)
        cycles_entry.insert(0, "3")
        cycles_entry.grid(row=0, column=4, sticky='w', padx=5)
        tk.Label(params_frame, text="(MAX: 10)", font=('Arial', 9, 'italic'), fg='#666666', bg=COLORS['bg_frame']).grid(row=0, column=5, sticky='w', padx=5)
        
        tk.Label(params_frame, text="Duration:", font=('Arial', 10, 'bold'), bg=COLORS['bg_frame']).grid(row=1, column=3, sticky='w', padx=5, pady=3)
        duration_entry = tk.Entry(params_frame, font=('Arial', 10), width=10)
        duration_entry.insert(0, "8.0")
        duration_entry.grid(row=1, column=4, sticky='w', padx=5)
        tk.Label(params_frame, text="seconds (MIN: 3s)", font=('Arial', 9, 'italic'), fg='#666666', bg=COLORS['bg_frame']).grid(row=1, column=5, sticky='w', padx=5)
        
        # Store references to entries
        self.ramp_target_voltage = voltage_entry
        self.ramp_points_per_cycle = points_entry
        self.ramp_cycles = cycles_entry
        self.ramp_cycle_duration = duration_entry
        
        # Add settling time parameter
        tk.Label(params_frame, text="Settling Time:", font=('Arial', 10, 'bold'), bg=COLORS['bg_frame']).grid(row=2, column=0, sticky='w', padx=5, pady=3)
        settle_entry = tk.Entry(params_frame, font=('Arial', 10), width=10)
        settle_entry.insert(0, "0.5")
        settle_entry.grid(row=2, column=1, sticky='w', padx=5)
        tk.Label(params_frame, text="seconds (0.1 - 2.0)", font=('Arial', 9, 'italic'), fg='#666666', bg=COLORS['bg_frame']).grid(row=2, column=2, sticky='w', padx=5)
        self.ramp_psu_settle = settle_entry
        
        waveform_frame = tk.Frame(config_frame, bg=COLORS['bg_frame'])
        waveform_frame.pack(pady=10)
        
        tk.Label(
            waveform_frame,
            text="Waveform:",
            font=('Arial', 10, 'bold'),
            bg=COLORS['bg_frame']
        ).pack(side=tk.LEFT, padx=5)
        
        self.ramp_waveform = ttk.Combobox(
            waveform_frame,
            values=[w.value for w in WaveformType],
            state='readonly',
            width=15,
            font=('Arial', 10)
        )
        self.ramp_waveform.set("Sine")
        self.ramp_waveform.pack(side=tk.LEFT, padx=5)
        
        tk.Label(
            waveform_frame,
            text="Active PSU Channel:",
            font=('Arial', 10, 'bold'),
            bg=COLORS['bg_frame']
        ).pack(side=tk.LEFT, padx=(20, 5))
        
        self.ramp_psu_channel = ttk.Combobox(
            waveform_frame,
            values=['1', '2', '3'],
            state='readonly',
            width=5,
            font=('Arial', 10)
        )
        self.ramp_psu_channel.set('1')
        self.ramp_psu_channel.pack(side=tk.LEFT, padx=5)
        
        control_frame = tk.Frame(scrollable_frame, bg=COLORS['bg_frame'])
        control_frame.pack(pady=20)
        
        self.start_ramping_btn = tk.Button(
            control_frame,
            text="Start Ramping Test",
            font=('Segoe UI', 12, 'bold'),
            bg=COLORS['btn_success'],
            fg='white',
            width=25,
            height=2,
            command=self._start_ramping_test
        )
        self.start_ramping_btn.pack(side=tk.LEFT, padx=10)
        self.stop_ramping_btn = tk.Button(
            control_frame,
            text="Stop Ramping Test",
            font=('Segoe UI', 12, 'bold'),
            bg=COLORS['btn_emergency'],
            fg='white',
            width=25,
            height=2,
            command=self._stop_ramping_test,
            state=tk.DISABLED
        )
        self.stop_ramping_btn.pack(side=tk.LEFT, padx=10)
        
        progress_frame = tk.LabelFrame(
            scrollable_frame,
            text="Test Progress",
            font=('Arial', 11, 'bold'),
            bg=COLORS['bg_frame']
        )
        progress_frame.pack(fill=tk.X, padx=20, pady=10)
        
        self.ramp_progress_label = tk.Label(
            progress_frame,
            text="Ready to start ramping test",
            font=('Arial', 10),
            bg=COLORS['bg_frame']
        )
        self.ramp_progress_label.pack(pady=10)
        
        self.ramp_progress_bar = ttk.Progressbar(
            progress_frame,
            length=400,
            mode='determinate'
        )
        self.ramp_progress_bar.pack(pady=10)
    
    def _add_param_row(self, parent, label_text, default_value, hint, attr_name):
        row = tk.Frame(parent, bg=COLORS['bg_frame'])
        row.pack(fill=tk.X, pady=5)
        
        tk.Label(
            row,
            text=label_text,
            font=('Arial', 10, 'bold'),
            bg=COLORS['bg_frame'],
            width=20,
            anchor='w'
        ).pack(side=tk.LEFT)
        
        entry = tk.Entry(row, font=('Arial', 10), width=10)
        entry.insert(0, default_value)
        entry.pack(side=tk.LEFT, padx=5)
        setattr(self, attr_name, entry)
        
        tk.Label(
            row,
            text=hint,
            font=('Arial', 9, 'italic'),
            bg=COLORS['bg_frame'],
            fg='#666666'
        ).pack(side=tk.LEFT, padx=5)
    
    def _create_psu_tab(self):
        """Create the power supply control tab using standardized components"""
        # Create main tab container
        psu_frame = tk.Frame(self.notebook, bg=COLORS['bg_frame'])
        self.notebook.add(psu_frame, text="Power Supply")
        
        # Create scrollable container
        _, content = self._create_scrollable_frame(psu_frame)
        
        # Add title
        title = tk.Label(content, text="Keithley Multi-Channel Power Supply Control",
                        font=('Arial', 14, 'bold'), bg=COLORS['bg_frame'], fg='#1976d2')
        title.pack(pady=10)
        
        # Store channel widget references
        self.psu_channel_widgets = {}
        
        # Create channel controls
        for channel in range(1, 4):
            self._create_psu_channel_control(content, channel)
        
        # Create global controls section
        global_frame = self._create_labeled_frame(content, "Global Controls")
        
        # Create button grid for global controls
        buttons = [
            ("Measure All Channels", self._measure_all_psu_channels, COLORS['btn_info']),
            ("Disable All Outputs", self._disable_all_psu_outputs, COLORS['btn_warning'])
        ]
        self._create_button_grid(global_frame, buttons, columns=2)

        # Create file preferences section
        prefs = self._create_labeled_frame(content, "File Preferences")
        
        # Create data path selector
        row = tk.Frame(prefs, bg=COLORS['bg_frame'])
        row.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(row, text="Data:", font=('Arial', 10, 'bold'), bg=COLORS['bg_frame']).pack(side=tk.LEFT)
        entry = tk.Entry(row, font=('Arial', 9), width=60)
        entry.insert(0, str(self.data_handler.data_folder))
        entry.pack(side=tk.LEFT, padx=5)
        tk.Button(row, text="...", width=3, command=lambda: self._browse_global_folder('data', entry)).pack(side=tk.LEFT)
        
        # Add export button
        tk.Button(
            prefs,
            text="Export PSU Measurements",
            font=('Arial', 11, 'bold'),
            bg=COLORS['btn_success'],
            fg='white',
            command=lambda: self._export_data('psu')
        ).pack(pady=10)
    
    def _create_psu_channel_control(self, parent, channel: int):
        """Create a standardized channel control panel for the power supply"""
        # Create channel frame using helper method
        channel_frame = self._create_labeled_frame(parent, f"Channel {channel}")
        
        # Create settings panel
        settings_frame = self._create_labeled_frame(channel_frame, "Settings")
        param_frame = tk.Frame(settings_frame, bg=COLORS['bg_frame'])
        param_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Voltage controls
        tk.Label(
            param_frame,
            text="Voltage (V):",
            font=('Arial', 10, 'bold'),
            bg=COLORS['bg_frame']
        ).pack(side=tk.LEFT, padx=5)
        
        voltage_spinbox = tk.Spinbox(
            param_frame,
            from_=0.0,
            to=30.0,
            increment=0.1,
            format="%.2f",
            font=('Arial', 10),
            width=8
        )
        voltage_spinbox.delete(0, tk.END)
        voltage_spinbox.insert(0, "0.00")
        voltage_spinbox.pack(side=tk.LEFT, padx=5)
        
        # Current limit controls
        tk.Label(
            param_frame,
            text="Current Limit (A):",
            font=('Arial', 10, 'bold'),
            bg=COLORS['bg_frame']
        ).pack(side=tk.LEFT, padx=(20, 5))
        
        current_spinbox = tk.Spinbox(
            param_frame,
            from_=0.001,
            to=3.0,
            increment=0.01,
            format="%.3f",
            font=('Arial', 10),
            width=8
        )
        current_spinbox.delete(0, tk.END)
        current_spinbox.insert(0, "0.100")
        current_spinbox.pack(side=tk.LEFT, padx=5)
        
        # OVP controls
        tk.Label(
            param_frame,
            text="OVP (V):",
            font=('Arial', 10, 'bold'),
            bg=COLORS['bg_frame']
        ).pack(side=tk.LEFT, padx=(20, 5))
        
        ovp_spinbox = tk.Spinbox(
            param_frame,
            from_=1.0,
            to=35.0,
            increment=0.5,
            format="%.1f",
            font=('Arial', 10),
            width=8
        )
        ovp_spinbox.delete(0, tk.END)
        ovp_spinbox.insert(0, "5.0")
        ovp_spinbox.pack(side=tk.LEFT, padx=5)
        
        # Create control buttons using grid layout
        buttons = [
            ("Configure", lambda: self._configure_psu_channel(channel), COLORS['btn_info']),
            ("Enable Output", lambda: self._enable_psu_channel(channel), COLORS['btn_success']),
            ("Disable Output", lambda: self._disable_psu_channel(channel), COLORS['btn_warning']),
            ("Measure", lambda: self._measure_psu_channel(channel), COLORS['btn_info'])
        ]
        btn_frame = self._create_button_grid(channel_frame, buttons, columns=4)
        
        # Create measurements display panel
        meas_frame = self._create_labeled_frame(channel_frame, "Measurements")
        meas_inner = tk.Frame(meas_frame, bg=COLORS['bg_frame'])
        meas_inner.pack(fill=tk.X, padx=10, pady=5)
        
        meas_label = tk.Label(
            meas_inner,
            text="0.000 V    0.000 A    0.000 W",
            font=('Courier', 11, 'bold'),
            bg=COLORS['bg_frame'],
            fg='#0000ff'
        )
        meas_label.pack(side=tk.LEFT, padx=10)
        
        status_label = tk.Label(
            meas_inner,
            text="OFF",
            font=('Arial', 11, 'bold'),
            bg=COLORS['bg_frame'],
            fg=COLORS['status_disconnected']
        )
        status_label.pack(side=tk.RIGHT, padx=10)
        
        # Store widget references
        self.psu_channel_widgets[channel] = {
            'voltage': voltage_spinbox,
            'current': current_spinbox,
            'ovp': ovp_spinbox,
            'configure_btn': btn_frame.winfo_children()[0],
            'enable_btn': btn_frame.winfo_children()[1],
            'disable_btn': btn_frame.winfo_children()[2],
            'measure_btn': btn_frame.winfo_children()[3],
            'meas_label': meas_label,
            'status_label': status_label
        }
    
    def _create_dmm_tab(self):
        """Create the DMM control tab using standardized components"""
        # Create main tab container
        dmm_frame = tk.Frame(self.notebook, bg=COLORS['bg_frame'])
        self.notebook.add(dmm_frame, text="DMM")
        
        # Create scrollable container
        _, content = self._create_scrollable_frame(dmm_frame)
        
        # Add title
        title = tk.Label(content, text="Keithley DMM6500 Professional Automation",
                        font=('Arial', 14, 'bold'), bg=COLORS['bg_frame'], fg='#1976d2')
        title.pack(pady=10)
        
        # Create measurement configuration section
        config_frame = self._create_labeled_frame(content, "Measurement Configuration")
        
        config_row1 = tk.Frame(config_frame, bg=COLORS['bg_frame'])
        config_row1.pack(fill=tk.X, pady=5)
        
        tk.Label(
            config_row1,
            text="Function:",
            font=('Arial', 10, 'bold'),
            bg=COLORS['bg_frame'],
            width=15,
            anchor='w'
        ).pack(side=tk.LEFT, padx=5)
        
        self.dmm_function = ttk.Combobox(
            config_row1,
            values=[
                "DC Voltage",
                "AC Voltage",
                "DC Current",
                "AC Current",
                "2-Wire Resistance",
                "4-Wire Resistance",
                "Capacitance",
                "Frequency"
            ],
            state='readonly',
            width=20,
            font=('Arial', 10)
        )
        self.dmm_function.set("DC Voltage")
        self.dmm_function.pack(side=tk.LEFT, padx=5)
        
        # Create operations section
        ops_frame = self._create_labeled_frame(content, "Operations")
        
        # Create button grid
        buttons = [
            ("Single Measurement", self._dmm_single_measurement, COLORS['btn_primary']),
            ("Start Continuous", self._dmm_start_continuous, COLORS['btn_success']),
            ("Stop Continuous", self._dmm_stop_continuous, COLORS['btn_emergency'])
        ]
        btn_frame = self._create_button_grid(ops_frame, buttons, columns=3)
        
        # Store button references
        self.dmm_single_btn = btn_frame.winfo_children()[0]
        self.dmm_continuous_start_btn = btn_frame.winfo_children()[1]
        self.dmm_continuous_stop_btn = btn_frame.winfo_children()[2]
        self.dmm_continuous_stop_btn.configure(state=tk.DISABLED)
        
        ops_row2 = tk.Frame(ops_frame, bg=COLORS['bg_frame'])
        ops_row2.pack(pady=5)
        
        tk.Label(
            ops_row2,
            text="Sampling Interval (s):",
            font=('Arial', 10, 'bold'),
            bg=COLORS['bg_frame']
        ).pack(side=tk.LEFT, padx=5)
        
        self.dmm_interval = tk.Entry(ops_row2, font=('Arial', 10), width=10)
        self.dmm_interval.insert(0, "1.0")
        self.dmm_interval.pack(side=tk.LEFT, padx=5)
        
        # Create measurement display section
        display_frame = self._create_labeled_frame(content, "Current Measurement")
        
        display_inner = tk.Frame(display_frame, bg=COLORS['bg_frame'])
        display_inner.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        self.dmm_measurement_label = tk.Label(
            display_inner,
            text="-- Ready for measurement --",
            font=('Courier', 16, 'bold'),
            bg=COLORS['bg_frame'],
            fg='#0000ff',
            pady=10
        )
        self.dmm_measurement_label.pack()
        
        # Add file preferences section
        prefs = self._create_labeled_frame(content, "File Preferences")
        
        # Create data path selector
        row = tk.Frame(prefs, bg=COLORS['bg_frame'])
        row.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(row, text="Data:", font=('Arial', 10, 'bold'), bg=COLORS['bg_frame']).pack(side=tk.LEFT)
        entry = tk.Entry(row, font=('Arial', 9), width=60)
        entry.insert(0, str(self.data_handler.data_folder))
        entry.pack(side=tk.LEFT, padx=5)
        tk.Button(row, text="...", width=3, command=lambda e=entry: self._browse_global_folder('data', e)).pack(side=tk.LEFT)
        
        # Add export button for saving measurements
        tk.Button(
            display_inner,
            text="Export Measurements",
            font=('Arial', 11, 'bold'),
            bg=COLORS['btn_success'],
            fg='white',
            command=lambda: self._export_data('dmm')
        ).pack(pady=(0, 10))
    
    def _create_scope_tab(self):
        """Create the oscilloscope control tab using standardized components"""
        # Create main tab container
        scope_frame = tk.Frame(self.notebook, bg=COLORS['bg_frame'])
        self.notebook.add(scope_frame, text="Oscilloscope")
        
        # Create scrollable container
        _, content = self._create_scrollable_frame(scope_frame)
        
        # Add title
        title = tk.Label(content, text="Keysight Oscilloscope Controls",
                        font=('Arial', 14, 'bold'), bg=COLORS['bg_frame'], fg='#1976d2')
        title.pack(pady=10)
        
        # Create file preferences section
        prefs = self._create_labeled_frame(content, "File Preferences")
        
        def add_row(parent, label, entry_attr, initial, browse_type):
            row = tk.Frame(parent, bg=COLORS['bg_frame'])
            row.pack(fill=tk.X, padx=10, pady=5)
            tk.Label(row, text=label, font=('Arial', 10, 'bold'), bg=COLORS['bg_frame'], width=12, anchor='w').pack(side=tk.LEFT)
            entry = tk.Entry(row, font=('Arial', 9))
            entry.insert(0, initial)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            setattr(self, entry_attr, entry)
            tk.Button(row, text="...", width=3, command=lambda: self._browse_save_folder(browse_type)).pack(side=tk.LEFT)

        add_row(prefs, "Data:", 'scope_data_path', str(self.data_handler.data_folder), 'data')
        add_row(prefs, "Graphs:", 'scope_graph_path', str(self.data_handler.graph_folder), 'graphs')
        add_row(prefs, "Screenshots:", 'scope_screenshot_path', str(self.data_handler.screenshot_folder), 'screenshots')

        plot_row = tk.Frame(prefs, bg=COLORS['bg_frame'])
        plot_row.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(plot_row, text="Plot Title:", font=('Arial', 10, 'bold'), bg=COLORS['bg_frame'], width=12, anchor='w').pack(side=tk.LEFT)
        self.scope_plot_title = tk.Entry(plot_row, font=('Arial', 10))
        self.scope_plot_title.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # Create channel selection section
        ch_frame = self._create_labeled_frame(content, "Channels")
        ch_inner = tk.Frame(ch_frame, bg=COLORS['bg_frame'])
        ch_inner.pack(fill=tk.X, padx=10, pady=5)
        
        def on_channel_toggle(channel):
            """Toggle oscilloscope channel display using the scope wrapper's available API.
            This will attempt to use the high-level wrapper methods first, then fall back to
            directly using the underlying SCPI wrapper if present.
            """
            if not self.instruments_connected['scope']:
                # No-op when scope not connected
                return
            try:
                enabled = self.scope_channel_vars[channel].get() == 1

                # Prefer high-level API if available
                if hasattr(self.scope, 'configure_channel'):
                    # configure_channel will set display on as part of configuration
                    # We call with a no-op configuration (keep existing scale) if possible
                    # Here we attempt to read current vertical scale; if not possible, use 1.0
                    try:
                        current_scale = 1.0
                        # Try to query current scale using the wrapper's SCPI layer if available
                        if hasattr(self.scope, '_scpi_wrapper') and hasattr(self.scope._scpi_wrapper, 'query'):
                            resp = self.scope._scpi_wrapper.query(f":CHANnel{channel}:SCALe?").strip()
                            current_scale = float(resp) if resp else 1.0
                    except Exception:
                        current_scale = 1.0

                    # Use configure_channel to set display ON/OFF and keep other params
                    # Some wrappers always turn display on; when disabling, we send direct SCPI below
                    ok = self.scope.configure_channel(channel, current_scale, vertical_offset=0.0)
                    if not ok and not enabled:
                        # fall back to SCPI write for disabling if configure_channel didn't work
                        raise RuntimeError('configure_channel failed')

                # If high-level methods are not available or we explicitly need to toggle display,
                # use underlying SCPI wrapper directly (safe check for attribute presence)
                if hasattr(self.scope, '_scpi_wrapper') and hasattr(self.scope._scpi_wrapper, 'write'):
                    cmd = f":CHANnel{channel}:DISPlay {'ON' if enabled else 'OFF'}"
                    self.scope._scpi_wrapper.write(cmd)
                    # Verify using query if available
                    if hasattr(self.scope._scpi_wrapper, 'query'):
                        resp = self.scope._scpi_wrapper.query(f":CHANnel{channel}:DISPlay?").strip()
                        # Some scopes return '1'/'0' or 'ON'/'OFF', normalize
                        if resp.upper() in ('ON', 'OFF'):
                            actual_enabled = resp.upper() == 'ON'
                        else:
                            try:
                                actual_enabled = int(resp) == 1
                            except Exception:
                                actual_enabled = enabled

                        if actual_enabled == enabled:
                            self.log_message(f"Oscilloscope Channel {channel} {'enabled' if enabled else 'disabled'}", 'SUCCESS')
                        else:
                            raise RuntimeError(f"Channel state verification failed (resp={resp})")
                else:
                    # If neither high-level nor SCPI write available, raise informative error
                    raise AttributeError('Scope wrapper does not expose SCPI interface or configure_channel')

            except Exception as e:
                self.log_message(f"Error toggling channel {channel}: {e}", 'ERROR')
                # Reset checkbox to previous state
                self.scope_channel_vars[channel].set(not self.scope_channel_vars[channel].get())
        
        self.scope_channel_vars = {}
        for ch in [1, 2, 3, 4]:
            var = tk.IntVar(value=1 if ch == 1 else 0)
            self.scope_channel_vars[ch] = var
            cb = tk.Checkbutton(ch_inner, text=f"Channel {ch}", variable=var, bg=COLORS['bg_frame'],
                              font=('Arial', 10), command=lambda c=ch: on_channel_toggle(c))
            cb.pack(side=tk.LEFT, padx=10)

        # Create function generator section
        wgen_frame = self._create_labeled_frame(content, "Function Generator")
        wgen_grid = tk.Frame(wgen_frame, bg=COLORS['bg_frame'])
        wgen_grid.pack(fill=tk.X, padx=10, pady=10)
        
        for i in range(6):
            wgen_grid.columnconfigure(i, weight=1)

        self.wgen_gen = tk.StringVar(value="1")
        self.wgen_waveform = tk.StringVar(value="SIN")
        self.wgen_freq = tk.StringVar(value="1000")
        self.wgen_amp = tk.StringVar(value="1.0")
        self.wgen_offset = tk.StringVar(value="0.0")
        self.wgen_enabled = tk.BooleanVar(value=True)

        # Generator controls
        row1 = tk.Frame(wgen_grid, bg=COLORS['bg_frame'])
        row1.pack(fill=tk.X, pady=5)
        
        ttk.Label(row1, text="Generator:", background=COLORS['bg_frame'], font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)
        ttk.Combobox(row1, values=["1", "2"], textvariable=self.wgen_gen, width=6, state="readonly").pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row1, text="Waveform:", background=COLORS['bg_frame'], font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)
        ttk.Combobox(row1, values=["SIN", "SQU", "RAMP", "PULS", "DC", "NOIS"], 
                    textvariable=self.wgen_waveform, state="readonly", width=10).pack(side=tk.LEFT, padx=5)

        # Parameters row
        row2 = tk.Frame(wgen_grid, bg=COLORS['bg_frame'])
        row2.pack(fill=tk.X, pady=5)
        
        ttk.Label(row2, text="Frequency (Hz):", background=COLORS['bg_frame'], font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)
        ttk.Entry(row2, textvariable=self.wgen_freq, width=12).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row2, text="Amplitude (Vpp):", background=COLORS['bg_frame'], font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)
        ttk.Entry(row2, textvariable=self.wgen_amp, width=12).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row2, text="Offset (V):", background=COLORS['bg_frame'], font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)
        ttk.Entry(row2, textvariable=self.wgen_offset, width=12).pack(side=tk.LEFT, padx=5)

        # Function generator buttons
        buttons = [
            ("Apply", self._wgen_apply, COLORS['btn_primary']),
            ("Enable", lambda: self._wgen_enable(True), COLORS['btn_success']),
            ("Disable", lambda: self._wgen_enable(False), COLORS['btn_warning']),
            ("Read Config", self._wgen_read, COLORS['btn_info'])
        ]
        self._create_button_grid(wgen_frame, buttons, columns=4)

        # Create operations section with button grid
        ops_frame = self._create_labeled_frame(content, "Operations")
        
        # Row 1 buttons
        buttons_row1 = [
            ("Take Screenshot", self._scope_screenshot, COLORS['btn_info']),
            ("Acquire Data", self._scope_acquire_data, COLORS['btn_primary']),
            ("Export CSV", self._scope_export_csv, COLORS['btn_success'])
        ]
        self._create_button_grid(ops_frame, buttons_row1, columns=3)
        
        # Row 2 buttons
        buttons_row2 = [
            ("Open Folders", lambda: self._open_folders([
                ("Data", self.scope_data_path.get()),
                ("Graphs", self.scope_graph_path.get()),
                ("Screenshots", self.scope_screenshot_path.get())
            ]), COLORS['btn_info']),
            ("Generate Plot", self._scope_generate_plot, COLORS['btn_success']),
            ("Full Automation", self._scope_full_auto, COLORS['btn_primary'])
        ]
        self._create_button_grid(ops_frame, buttons_row2, columns=3)

    def _wgen_apply(self):
        if not self.instruments_connected['scope']:
            messagebox.showerror("Error", "Oscilloscope not connected")
            return
        def worker():
            try:
                gen = int(self.wgen_gen.get())
                wf = self.wgen_waveform.get().upper()
                freq = float(self.wgen_freq.get())
                amp = float(self.wgen_amp.get())
                off = float(self.wgen_offset.get())
                ok = self.scope.configure_function_generator(gen, wf, freq, amp, off, True)
                if ok:
                    self.status_queue.put(('wgen_config_ok', f'WGEN{gen}: {wf} {freq}Hz {amp}Vpp {off}V'))
                else:
                    self.status_queue.put(('wgen_error', 'Failed to apply configuration'))
            except Exception as e:
                self.status_queue.put(('wgen_error', str(e)))
        threading.Thread(target=worker, daemon=True).start()

    def _wgen_enable(self, enable: bool):
        if not self.instruments_connected['scope']:
            messagebox.showerror("Error", "Oscilloscope not connected")
            return
        def worker():
            try:
                gen = int(self.wgen_gen.get())
                ok = self.scope.enable_function_generator(gen, enable)
                if ok:
                    self.status_queue.put(('wgen_enable_ok', f'WGEN{gen} {"ENABLED" if enable else "DISABLED"}'))
                else:
                    self.status_queue.put(('wgen_error', 'Failed to change output state'))
            except Exception as e:
                self.status_queue.put(('wgen_error', str(e)))
        threading.Thread(target=worker, daemon=True).start()

    def _wgen_read(self):
        if not self.instruments_connected['scope']:
            messagebox.showerror("Error", "Oscilloscope not connected")
            return
        def worker():
            try:
                gen = int(self.wgen_gen.get())
                cfg = self.scope.get_function_generator_config(gen)
                if cfg:
                    self.status_queue.put(('wgen_config_read', cfg))
                else:
                    self.status_queue.put(('wgen_error', 'Failed to read configuration'))
            except Exception as e:
                self.status_queue.put(('wgen_error', str(e)))
        threading.Thread(target=worker, daemon=True).start()
    
    def _create_data_export_tab(self):
        export_frame = tk.Frame(self.notebook, bg=COLORS['bg_frame'])
        self.notebook.add(export_frame, text="Data Export")
        
        title = tk.Label(
            export_frame,
            text="Data Export & Analysis",
            font=('Arial', 14, 'bold'),
            bg=COLORS['bg_frame'],
            fg='#1976d2'
        )
        title.pack(pady=20)
        
        btn_frame = tk.Frame(export_frame, bg=COLORS['bg_frame'])
        btn_frame.pack(pady=20)
        
        tk.Button(
            btn_frame,
            text="Export PSU Data (CSV)",
            font=('Arial', 11),
            bg=COLORS['btn_success'],
            fg='white',
            width=30,
            command=lambda: self._export_data('psu')
        ).pack(pady=5)
        
        tk.Button(
            btn_frame,
            text="Export DMM Data (CSV)",
            font=('Arial', 11),
            bg=COLORS['btn_success'],
            fg='white',
            width=30,
            command=lambda: self._export_data('dmm')
        ).pack(pady=5)
        
        tk.Button(
            btn_frame,
            text="Export Ramping Data (CSV)",
            font=('Arial', 11),
            bg=COLORS['btn_success'],
            fg='white',
            width=30,
            command=lambda: self._export_data('ramping')
        ).pack(pady=5)
        
        tk.Button(
            btn_frame,
            text="Generate Ramping Analysis Graph",
            font=('Arial', 11),
            bg=COLORS['btn_info'],
            fg='white',
            width=30,
            command=self._generate_ramping_graph
        ).pack(pady=5)
    
    def _create_activity_log(self, parent):
        log_frame = tk.LabelFrame(
            parent,
            text="Activity Log",
            font=('Arial', 11, 'bold'),
            bg=COLORS['bg_frame']
        )
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        self.activity_log = scrolledtext.ScrolledText(
            log_frame,
            height=8,
            font=('Consolas', 9),
            bg=COLORS['bg_log'],
            fg=COLORS['fg_log'],
            insertbackground='white',
            state=tk.DISABLED
        )
        self.activity_log.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.activity_log.tag_config('INFO', foreground='#00ff41')
        self.activity_log.tag_config('SUCCESS', foreground='#00ff00', font=('Consolas', 9, 'bold'))
        self.activity_log.tag_config('ERROR', foreground='#ff0000', font=('Consolas', 9, 'bold'))
        self.activity_log.tag_config('WARNING', foreground='#ff9800', font=('Consolas', 9, 'bold'))
    
    def log_message(self, message: str, level: str = 'INFO'):
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        self.activity_log.configure(state=tk.NORMAL)
        self.activity_log.insert(tk.END, formatted_message, level)
        self.activity_log.see(tk.END)
        self.activity_log.configure(state=tk.DISABLED)

        # Update status bar if present
        if hasattr(self, 'status_var'):
            # Show level and short message
            short = message if len(message) < 80 else message[:77] + '...'
            self.status_var.set(f"{level}: {short}")

        logger.info(message)
    
    # Connection methods
    def _connect_instrument(self, instrument_type: str):
        def connect_thread():
            try:
                if instrument_type == 'psu':
                    visa_address = self.psu_visa_entry.get()
                    self.log_message(f"Connecting to Power Supply: {visa_address}")
                    self.psu = KeithleyPowerSupply(visa_address)
                    self.psu.connect()
                    info = self.psu.get_instrument_info()
                    self.instruments_connected['psu'] = True
                    self.status_queue.put(('psu_connected', str(info)))
                    
                elif instrument_type == 'dmm':
                    visa_address = self.dmm_visa_entry.get()
                    self.log_message(f"Connecting to DMM: {visa_address}")
                    self.dmm = KeithleyDMM6500(visa_address)
                    self.dmm.connect()
                    info = self.dmm.get_instrument_info()
                    self.instruments_connected['dmm'] = True
                    self.status_queue.put(('dmm_connected', str(info)))
                    
                elif instrument_type == 'scope':
                    visa_address = self.scope_visa_entry.get()
                    self.log_message(f"Connecting to Oscilloscope: {visa_address}")
                    self.scope = KeysightDSOX6004A(visa_address)
                    self.scope.connect()
                    info = self.scope.get_instrument_info()
                    self.instruments_connected['scope'] = True
                    # Initialize data acquisition utilities if available
                    if OscilloscopeDataAcquisition is not None:
                        try:
                            self.scope_data_acq = OscilloscopeDataAcquisition(self.scope)
                        except Exception:
                            self.scope_data_acq = None
                    # Ensure folders exist
                    for folder in [self.data_handler.data_folder, self.data_handler.graph_folder, self.data_handler.screenshot_folder]:
                        folder.mkdir(parents=True, exist_ok=True)
                    self.status_queue.put(('scope_connected', str(info)))
                    
            except Exception as e:
                self.status_queue.put((f'{instrument_type}_error', str(e)))
        
        thread = threading.Thread(target=connect_thread, daemon=True)
        thread.start()
    
    def _disconnect_instrument(self, instrument_type: str):
        try:
            if instrument_type == 'psu' and self.psu:
                self.psu.disconnect()
                self.psu = None
                self.instruments_connected['psu'] = False
                self.log_message("Power Supply disconnected", 'INFO')
                self._update_connection_status('psu', False)
                
            elif instrument_type == 'dmm' and self.dmm:
                self.dmm.disconnect()
                self.dmm = None
                self.instruments_connected['dmm'] = False
                self.log_message("DMM disconnected", 'INFO')
                self._update_connection_status('dmm', False)
                
            elif instrument_type == 'scope' and self.scope:
                self.scope.disconnect()
                self.scope = None
                self.instruments_connected['scope'] = False
                self.log_message("Oscilloscope disconnected", 'INFO')
                self._update_connection_status('scope', False)
                
        except Exception as e:
            self.log_message(f"Error disconnecting {instrument_type}: {e}", 'ERROR')
    
    def _connect_all_instruments(self):
        self.log_message("Connecting to all instruments...", 'INFO')
        for inst_type in ['psu', 'dmm', 'scope']:
            if not self.instruments_connected[inst_type]:
                self._connect_instrument(inst_type)
                time.sleep(0.5)
    
    def _disconnect_all_instruments(self):
        self.log_message("Disconnecting all instruments...", 'INFO')
        for inst_type in ['psu', 'dmm', 'scope']:
            if self.instruments_connected[inst_type]:
                self._disconnect_instrument(inst_type)
    
    def _update_connection_status(self, instrument_type: str, connected: bool):
        status_text = "Connected" if connected else "Disconnected"
        status_color = COLORS['status_connected'] if connected else COLORS['status_disconnected']
        
        if instrument_type == 'psu':
            self.psu_status_label.config(text=status_text, fg=status_color)
            self.psu_connect_btn.config(state=tk.DISABLED if connected else tk.NORMAL)
            self.psu_disconnect_btn.config(state=tk.NORMAL if connected else tk.DISABLED)
        elif instrument_type == 'dmm':
            self.dmm_status_label.config(text=status_text, fg=status_color)
            self.dmm_connect_btn.config(state=tk.DISABLED if connected else tk.NORMAL)
            self.dmm_disconnect_btn.config(state=tk.NORMAL if connected else tk.DISABLED)
        elif instrument_type == 'scope':
            self.scope_status_label.config(text=status_text, fg=status_color)
            self.scope_connect_btn.config(state=tk.DISABLED if connected else tk.NORMAL)
            self.scope_disconnect_btn.config(state=tk.NORMAL if connected else tk.DISABLED)
    
    # Emergency stop
    def _emergency_stop(self):
        self.log_message("=" * 60, 'ERROR')
        self.log_message("EMERGENCY STOP ACTIVATED", 'ERROR')
        self.log_message("=" * 60, 'ERROR')
        
        self.emergency_stop_active = True
        self.system_status_label.config(
            text="EMERGENCY STOP ACTIVE",
            fg=COLORS['btn_emergency']
        )
        
        if self.ramping_active:
            self._stop_ramping_test()
        
        if self.psu and self.instruments_connected['psu']:
            try:
                for channel in [1, 2, 3]:
                    self.psu.disable_channel_output(channel)
                    self.log_message(f"PSU Channel {channel} DISABLED", 'ERROR')
                self.log_message("All PSU outputs disabled successfully", 'SUCCESS')
            except Exception as e:
                self.log_message(f"Error during emergency stop: {e}", 'ERROR')
        else:
            self.log_message("PSU not connected - cannot disable outputs", 'WARNING')
        
        messagebox.showwarning(
            "Emergency Stop",
            "EMERGENCY STOP ACTIVATED\n\nAll power supply outputs have been disabled.\n\n"
            "Please verify system safety before resuming operations."
        )
        
        self.emergency_stop_active = False
        self.system_status_label.config(
            text="● System Safe - Ready for operation",
            fg=COLORS['status_safe']
        )
    
    # PSU methods
    def _configure_psu_channel(self, channel: int):
        if not self.instruments_connected['psu']:
            messagebox.showerror("Error", "Power Supply not connected")
            return
        
        try:
            widgets = self.psu_channel_widgets[channel]
            voltage = float(widgets['voltage'].get())
            current = float(widgets['current'].get())
            ovp = float(widgets['ovp'].get())
            
            self.psu.configure_channel(channel, voltage, current, ovp, enable_output=False)
            
            self.log_message(
                f"CH{channel} configured: {voltage}V, {current}A limit, {ovp}V OVP",
                'SUCCESS'
            )
            
        except Exception as e:
            self.log_message(f"Error configuring CH{channel}: {e}", 'ERROR')
            messagebox.showerror("Configuration Error", str(e))
    
    def _enable_psu_channel(self, channel: int):
        if not self.instruments_connected['psu']:
            messagebox.showerror("Error", "Power Supply not connected")
            return
        
        try:
            self.psu.enable_channel_output(channel)
            self.psu_channel_widgets[channel]['status_label'].config(
                text="ON",
                fg=COLORS['status_connected']
            )
            self.log_message(f"CH{channel} output ENABLED", 'SUCCESS')
        except Exception as e:
            self.log_message(f"Error enabling CH{channel}: {e}", 'ERROR')
    
    def _disable_psu_channel(self, channel: int):
        if not self.instruments_connected['psu']:
            return
        
        try:
            self.psu.disable_channel_output(channel)
            self.psu_channel_widgets[channel]['status_label'].config(
                text="OFF",
                fg=COLORS['status_disconnected']
            )
            self.log_message(f"CH{channel} output DISABLED", 'WARNING')
        except Exception as e:
            self.log_message(f"Error disabling CH{channel}: {e}", 'ERROR')
    
    def _measure_psu_channel(self, channel: int):
        if not self.instruments_connected['psu']:
            messagebox.showerror("Error", "Power Supply not connected")
            return
        
        def measure_thread():
            try:
                result = self.psu.measure_channel_output(channel)
                if result:
                    v, i = result
                    p = v * i
                    self.data_handler.add_psu_measurement(channel, v, i, p)
                    self.status_queue.put(('psu_measurement', (channel, v, i, p)))
            except Exception as e:
                self.status_queue.put(('psu_meas_error', (channel, str(e))))
        
        thread = threading.Thread(target=measure_thread, daemon=True)
        thread.start()
    
    def _measure_all_psu_channels(self):
        if not self.instruments_connected['psu']:
            messagebox.showerror("Error", "Power Supply not connected")
            return
        
        self.log_message("Measuring all PSU channels...", 'INFO')
        for channel in [1, 2, 3]:
            self._measure_psu_channel(channel)
            time.sleep(0.1)
    
    def _disable_all_psu_outputs(self):
        if not self.instruments_connected['psu']:
            messagebox.showerror("Error", "Power Supply not connected")
            return
        
        result = messagebox.askyesno(
            "Confirm",
            "Disable all power supply outputs?"
        )
        
        if result:
            for channel in [1, 2, 3]:
                self._disable_psu_channel(channel)
            self.log_message("All PSU outputs disabled", 'WARNING')
    
    # DMM methods
    def _dmm_single_measurement(self):
        if not self.instruments_connected['dmm']:
            messagebox.showerror("Error", "DMM not connected")
            return
        
        def measure_thread():
            try:
                function = self.dmm_function.get()
                
                if function == "DC Voltage":
                    value = self.dmm.measure_dc_voltage()
                    unit = "V"
                elif function == "AC Voltage":
                    value = self.dmm.measure_ac_voltage()
                    unit = "V"
                elif function == "DC Current":
                    value = self.dmm.measure_dc_current()
                    unit = "A"
                elif function == "AC Current":
                    value = self.dmm.measure_ac_current()
                    unit = "A"
                elif function == "2-Wire Resistance":
                    value = self.dmm.measure_resistance_2w()
                    unit = "Ω"
                elif function == "4-Wire Resistance":
                    value = self.dmm.measure_resistance_4w()
                    unit = "Ω"
                elif function == "Capacitance":
                    value = self.dmm.measure_capacitance()
                    unit = "F"
                elif function == "Frequency":
                    value = self.dmm.measure_frequency()
                    unit = "Hz"
                
                if value is not None:
                    self.data_handler.add_dmm_measurement(function, value, unit)
                    self.status_queue.put(('dmm_measurement', (function, value, unit)))
                
            except Exception as e:
                self.status_queue.put(('dmm_error', str(e)))
        
        thread = threading.Thread(target=measure_thread, daemon=True)
        thread.start()
    
    def _dmm_start_continuous(self):
        if not self.instruments_connected['dmm']:
            messagebox.showerror("Error", "DMM not connected")
            return
        
        try:
            interval = float(self.dmm_interval.get())
            if interval < 0.1:
                raise ValueError("Interval must be >= 0.1 seconds")
        except ValueError as e:
            messagebox.showerror("Invalid Interval", str(e))
            return
        
        self.dmm_continuous_active = True
        self.dmm_single_btn.config(state=tk.DISABLED)
        self.dmm_continuous_start_btn.config(state=tk.DISABLED)
        self.dmm_continuous_stop_btn.config(state=tk.NORMAL)
        
        def continuous_measure():
            while self.dmm_continuous_active:
                self._dmm_single_measurement()
                time.sleep(interval)
        
        thread = threading.Thread(target=continuous_measure, daemon=True)
        thread.start()
        self.log_message(f"DMM continuous measurement started ({interval}s)", 'SUCCESS')
    
    def _dmm_stop_continuous(self):
        self.dmm_continuous_active = False
        self.dmm_single_btn.config(state=tk.NORMAL)
        self.dmm_continuous_start_btn.config(state=tk.NORMAL)
        self.dmm_continuous_stop_btn.config(state=tk.DISABLED)
        self.log_message("DMM continuous measurement stopped", 'INFO')
    
    # Scope methods
    def _scope_screenshot(self):
        if not self.instruments_connected['scope']:
            messagebox.showerror("Error", "Oscilloscope not connected")
            return
        
        def screenshot_thread():
            try:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                # Use current UI selection for screenshot path
                screenshot_dir = Path(self.scope_screenshot_path.get() or str(self.data_handler.screenshot_folder))
                screenshot_dir.mkdir(parents=True, exist_ok=True)
                filename = screenshot_dir / f"Scope_Screenshot_{timestamp}.png"
                
                result = self.scope.capture_screenshot(str(filename))
                if result:
                    self.status_queue.put(('scope_screenshot_done', str(filename)))
                
            except Exception as e:
                self.status_queue.put(('scope_error', str(e)))
        
        thread = threading.Thread(target=screenshot_thread, daemon=True)
        thread.start()
        self.log_message("Capturing oscilloscope screenshot...", 'INFO')

    def _browse_save_folder(self, folder_type: str):
        try:
            initial = {
                'data': self.scope_data_path.get(),
                'graphs': self.scope_graph_path.get(),
                'screenshots': self.scope_screenshot_path.get(),
            }.get(folder_type, str(self.data_handler.data_folder))
            folder = filedialog.askdirectory(initialdir=initial or str(Path.home()), title=f"Select {folder_type.title()} Folder")
            if folder:
                if folder_type == 'data':
                    self.scope_data_path.delete(0, tk.END)
                    self.scope_data_path.insert(0, folder)
                    self.data_handler.data_folder = Path(folder)
                elif folder_type == 'graphs':
                    self.scope_graph_path.delete(0, tk.END)
                    self.scope_graph_path.insert(0, folder)
                    self.data_handler.graph_folder = Path(folder)
                elif folder_type == 'screenshots':
                    self.scope_screenshot_path.delete(0, tk.END)
                    self.scope_screenshot_path.insert(0, folder)
                    self.data_handler.screenshot_folder = Path(folder)
                Path(folder).mkdir(parents=True, exist_ok=True)
                self.log_message(f"Updated {folder_type} folder: {folder}", 'SUCCESS')
        except Exception as e:
            self.log_message(f"Folder selection error: {e}", 'ERROR')

    def _get_selected_scope_channels(self) -> List[int]:
        return [ch for ch, var in self.scope_channel_vars.items() if var.get() == 1]

    def _scope_acquire_data(self):
        if not self.instruments_connected['scope']:
            messagebox.showerror("Error", "Oscilloscope not connected")
            return
        if self.scope_data_acq is None:
            messagebox.showerror("Unavailable", "Oscilloscope data acquisition module not available")
            return
        
        selected = self._get_selected_scope_channels()
        if not selected:
            messagebox.showwarning("No Channels", "Select at least one channel")
            return
        
        def acquire_thread():
            try:
                all_data = {}
                for ch in selected:
                    self.log_message(f"Acquiring Ch{ch}...", 'INFO')
                    data = self.scope_data_acq.acquire_waveform_data(ch)
                    if data:
                        all_data[ch] = data
                if all_data:
                    self.scope_last_data = all_data if len(all_data) > 1 else list(all_data.values())[0]
                    self.status_queue.put(('scope_data_acquired', self.scope_last_data))
                else:
                    self.status_queue.put(('scope_error', 'Data acquisition failed for all selected channels'))
            except Exception as e:
                self.status_queue.put(('scope_error', str(e)))
        
        threading.Thread(target=acquire_thread, daemon=True).start()
        self.log_message("Starting data acquisition...", 'INFO')

    def _scope_export_csv(self):
        if self.scope_last_data is None:
            messagebox.showwarning("No Data", "Acquire data first")
            return
        if self.scope_data_acq is None:
            messagebox.showerror("Unavailable", "Oscilloscope data acquisition module not available")
            return
        
        def export_thread():
            try:
                out = []
                save_dir = self.scope_data_path.get() or str(self.data_handler.data_folder)
                Path(save_dir).mkdir(parents=True, exist_ok=True)
                if isinstance(self.scope_last_data, dict) and 'channel' not in self.scope_last_data:
                    for ch, data in self.scope_last_data.items():
                        fn = self.scope_data_acq.export_to_csv(data, custom_path=save_dir)
                        if fn:
                            out.append(fn)
                else:
                    fn = self.scope_data_acq.export_to_csv(self.scope_last_data, custom_path=save_dir)
                    if fn:
                        out.append(fn)
                if out:
                    self.status_queue.put(('scope_csv_exported', out))
                else:
                    self.status_queue.put(('scope_error', 'CSV export failed'))
            except Exception as e:
                self.status_queue.put(('scope_error', str(e)))
        
        threading.Thread(target=export_thread, daemon=True).start()
        self.log_message("Exporting CSV...", 'INFO')

    def _scope_generate_plot(self):
        if self.scope_last_data is None:
            messagebox.showwarning("No Data", "Acquire data first")
            return
        if self.scope_data_acq is None:
            messagebox.showerror("Unavailable", "Oscilloscope data acquisition module not available")
            return
        
        def plot_thread():
            try:
                out = []
                save_dir = self.scope_graph_path.get() or str(self.data_handler.graph_folder)
                Path(save_dir).mkdir(parents=True, exist_ok=True)
                base_title = self.scope_plot_title.get().strip() or None
                if isinstance(self.scope_last_data, dict) and 'channel' not in self.scope_last_data:
                    for ch, data in self.scope_last_data.items():
                        title = f"{base_title} - Channel {ch}" if base_title else None
                        fn = self.scope_data_acq.generate_waveform_plot(data, custom_path=save_dir, plot_title=title)
                        if fn:
                            out.append(fn)
                else:
                    fn = self.scope_data_acq.generate_waveform_plot(self.scope_last_data, custom_path=save_dir, plot_title=base_title)
                    if fn:
                        out.append(fn)
                if out:
                    self.status_queue.put(('scope_plot_generated', out))
                else:
                    self.status_queue.put(('scope_error', 'Plot generation failed'))
            except Exception as e:
                self.status_queue.put(('scope_error', str(e)))
        
        threading.Thread(target=plot_thread, daemon=True).start()
        self.log_message("Generating plot(s)...", 'INFO')

    def _scope_full_auto(self):
        if not self.instruments_connected['scope']:
            messagebox.showerror("Error", "Oscilloscope not connected")
            return
        if self.scope_data_acq is None:
            messagebox.showerror("Unavailable", "Oscilloscope data acquisition module not available")
            return
        
        selected = self._get_selected_scope_channels()
        if not selected:
            messagebox.showwarning("No Channels", "Select at least one channel")
            return
        
        def full_thread():
            try:
                # Screenshot
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                screenshot_dir = Path(self.scope_screenshot_path.get() or str(self.data_handler.screenshot_folder))
                screenshot_dir.mkdir(parents=True, exist_ok=True)
                screenshot_path = screenshot_dir / f"Scope_Screenshot_{timestamp}.png"
                self.scope.capture_screenshot(str(screenshot_path))
                
                # Acquire
                all_data = {}
                for ch in selected:
                    data = self.scope_data_acq.acquire_waveform_data(ch)
                    if data:
                        all_data[ch] = data
                if not all_data:
                    raise RuntimeError("Data acquisition failed for all channels")
                
                # CSV
                csv_dir = Path(self.scope_data_path.get() or str(self.data_handler.data_folder))
                csv_dir.mkdir(parents=True, exist_ok=True)
                csv_files = []
                for ch, data in all_data.items():
                    fn = self.scope_data_acq.export_to_csv(data, custom_path=str(csv_dir))
                    if fn:
                        csv_files.append(fn)
                
                # Plots
                plot_dir = Path(self.scope_graph_path.get() or str(self.data_handler.graph_folder))
                plot_dir.mkdir(parents=True, exist_ok=True)
                plot_files = []
                base_title = self.scope_plot_title.get().strip() or None
                for ch, data in all_data.items():
                    title = f"{base_title} - Channel {ch}" if base_title else None
                    fn = self.scope_data_acq.generate_waveform_plot(data, custom_path=str(plot_dir), plot_title=title)
                    if fn:
                        plot_files.append(fn)
                
                result = {
                    'screenshot': str(screenshot_path),
                    'csv': csv_files,
                    'plot': plot_files,
                    'data': all_data,
                    'channels': selected
                }
                self.scope_last_data = all_data
                self.status_queue.put(('scope_full_auto_done', result))
            except Exception as e:
                self.status_queue.put(('scope_error', str(e)))
        
        threading.Thread(target=full_thread, daemon=True).start()
        self.log_message("Starting full automation...", 'INFO')

    def _open_folders(self, folders: List[Tuple[str, str]]):
        try:
            import subprocess, platform
            for name, path in folders:
                p = Path(path)
                p.mkdir(parents=True, exist_ok=True)
                if platform.system() == 'Windows':
                    subprocess.run(['explorer', str(p)], check=True)
                elif platform.system() == 'Darwin':
                    subprocess.run(['open', str(p)], check=True)
                else:
                    subprocess.run(['xdg-open', str(p)], check=True)
                self.log_message(f"Opened {name} folder", 'INFO')
        except Exception as e:
            self.log_message(f"Open folder error: {e}", 'ERROR')
    
    # Ramping methods
    def _start_ramping_test(self):
        if not self.instruments_connected['psu']:
            messagebox.showerror("Error", "Power Supply not connected")
            return
        
        try:
            # Validate all input parameters
            target_voltage = float(self.ramp_target_voltage.get())
            points_per_cycle = int(self.ramp_points_per_cycle.get())
            cycles = int(self.ramp_cycles.get())
            cycle_duration = float(self.ramp_cycle_duration.get())
            settle_time = float(self.ramp_psu_settle.get())
            waveform = self.ramp_waveform.get()
            psu_channel = int(self.ramp_psu_channel.get())
            
            # Validate parameter ranges
            if target_voltage > 4.0:
                raise ValueError("Target voltage exceeds 4.0V safety limit")
            if target_voltage < 0:
                raise ValueError("Target voltage cannot be negative")
            if cycles > 10:
                raise ValueError("Cycles exceed maximum of 10")
            if cycles < 1:
                raise ValueError("Must have at least 1 cycle")
            if points_per_cycle < 20 or points_per_cycle > 100:
                raise ValueError("Points per cycle must be between 20 and 100")
            if cycle_duration < 3:
                raise ValueError("Cycle duration must be at least 3 seconds")
            if not 0.1 <= settle_time <= 2.0:
                raise ValueError("Settling time must be between 0.1 and 2.0 seconds")
            
        except ValueError as e:
            messagebox.showerror("Invalid Parameters", str(e))
            return
        
        self.log_message("Starting Voltage Ramping Test", 'SUCCESS')
        self.log_message(f"Parameters: {target_voltage}V, {cycles} cycles, {points_per_cycle} points", 'INFO')
        
        self.ramping_active = True
        self.start_ramping_btn.config(state=tk.DISABLED)
        self.stop_ramping_btn.config(state=tk.NORMAL)
        self.ramp_progress_bar['value'] = 0
        
        self.ramping_thread = threading.Thread(
            target=self._ramping_worker,
            args=(target_voltage, points_per_cycle, cycles, cycle_duration, 
                  settle_time, waveform, psu_channel),
            daemon=True
        )
        self.ramping_thread.start()
    
    def _stop_ramping_test(self):
        self.log_message("Stopping Voltage Ramping Test", 'WARNING')
        self.ramping_active = False
        self.start_ramping_btn.config(state=tk.NORMAL)
        self.stop_ramping_btn.config(state=tk.DISABLED)
    
    def _ramping_worker(self, target_voltage, points_per_cycle, cycles, 
                       cycle_duration, psu_settle, waveform, psu_channel):
        try:
            total_points = cycles * points_per_cycle
            point_interval = cycle_duration / points_per_cycle
            
            current_point = 0
            
            for cycle in range(1, cycles + 1):
                if not self.ramping_active:
                    break
                
                self.status_queue.put(('ramp_progress', f"Cycle {cycle}/{cycles}"))
                
                for point in range(points_per_cycle):
                    if not self.ramping_active:
                        break
                    
                    progress = point / points_per_cycle
                    
                    if waveform == "Sine":
                        voltage = target_voltage * np.sin(2 * np.pi * progress)
                    elif waveform == "Ramp":
                        voltage = target_voltage * progress
                    elif waveform == "Triangle":
                        if progress < 0.5:
                            voltage = target_voltage * 2 * progress
                        else:
                            voltage = target_voltage * 2 * (1 - progress)
                    elif waveform == "Square":
                        voltage = target_voltage if progress < 0.5 else 0
                    
                    # Set voltage and measure
                    self.psu.configure_channel(psu_channel, abs(voltage), 0.5, 5.0, enable_output=True)
                    time.sleep(psu_settle)
                    
                    result = self.psu.measure_channel_output(psu_channel)
                    if result:
                        v_meas, i_meas = result
                        
                        self.data_handler.add_ramping_data(
                            cycle, point, voltage, v_meas, i_meas
                        )
                    
                    current_point += 1
                    progress_percent = (current_point / total_points) * 100
                    self.status_queue.put(('ramp_progress_bar', progress_percent))
                    
                    time.sleep(point_interval)
            
            self.psu.configure_channel(psu_channel, 0.0, 0.5, 5.0, enable_output=False)
            self.status_queue.put(('ramp_complete', None))
            
        except Exception as e:
            self.status_queue.put(('ramp_error', str(e)))
    
    # Data export methods
    def _export_data(self, data_type: str):
        try:
            if data_type == 'psu':
                filename = self.data_handler.export_psu_data()
                self.log_message(f"PSU data exported: {filename}", 'SUCCESS')
            elif data_type == 'dmm':
                filename = self.data_handler.export_dmm_data()
                self.log_message(f"DMM data exported: {filename}", 'SUCCESS')
            elif data_type == 'ramping':
                filename = self.data_handler.export_ramping_data()
                self.log_message(f"Ramping data exported: {filename}", 'SUCCESS')
                
            messagebox.showinfo("Export Complete", f"Data exported to:\n{filename}")
            
        except ValueError as e:
            messagebox.showwarning("No Data", str(e))
        except Exception as e:
            messagebox.showerror("Export Error", str(e))
            self.log_message(f"Export error: {e}", 'ERROR')
    
    def _generate_ramping_graph(self):
        try:
            filename = self.data_handler.generate_ramping_graph()
            self.log_message(f"Ramping graph generated: {filename}", 'SUCCESS')
            messagebox.showinfo("Graph Generated", f"Graph saved to:\n{filename}")
        except ValueError as e:
            messagebox.showwarning("No Data", str(e))
        except Exception as e:
            messagebox.showerror("Graph Error", str(e))
            self.log_message(f"Graph generation error: {e}", 'ERROR')
    
    # Status monitoring
    def _start_status_checker(self):
        self._check_status_updates()
    
    def _check_status_updates(self):
        try:
            while True:
                status_type, data = self.status_queue.get_nowait()
                
                if status_type == 'psu_connected':
                    self.log_message(f"PSU Connected: {data}", 'SUCCESS')
                    self._update_connection_status('psu', True)
                elif status_type == 'dmm_connected':
                    self.log_message(f"DMM Connected: {data}", 'SUCCESS')
                    self._update_connection_status('dmm', True)
                elif status_type == 'scope_connected':
                    self.log_message(f"Scope Connected: {data}", 'SUCCESS')
                    self._update_connection_status('scope', True)
                elif status_type == 'psu_error':
                    self.log_message(f"PSU Connection Error: {data}", 'ERROR')
                    messagebox.showerror("Connection Error", f"PSU: {data}")
                elif status_type == 'dmm_error':
                    self.log_message(f"DMM Connection Error: {data}", 'ERROR')
                    messagebox.showerror("Connection Error", f"DMM: {data}")
                elif status_type == 'scope_error':
                    self.log_message(f"Scope Connection Error: {data}", 'ERROR')
                    messagebox.showerror("Connection Error", f"Scope: {data}")
                elif status_type == 'psu_measurement':
                    channel, v, i, p = data
                    widgets = self.psu_channel_widgets[channel]
                    widgets['meas_label'].config(
                        text=f"{v:.3f} V    {i:.3f} A    {p:.3f} W"
                    )
                elif status_type == 'dmm_measurement':
                    function, value, unit = data
                    self.dmm_measurement_label.config(
                        text=f"{value:.6e} {unit}"
                    )
                    self.log_message(f"DMM: {value:.6e} {unit} ({function})", 'SUCCESS')
                elif status_type == 'scope_screenshot_done':
                    self.log_message(f"Screenshot saved: {data}", 'SUCCESS')
                elif status_type == 'scope_data_acquired':
                    # data can be dict (multi) or single dict
                    if isinstance(data, dict) and 'channel' not in data:
                        total_points = sum(v.get('points_count', 0) for v in data.values())
                        self.log_message(f"Data acquired: {len(data)} channel(s), {total_points} total points", 'SUCCESS')
                    else:
                        self.log_message(f"Data acquired: Ch{data.get('channel', '?')} - {data.get('points_count', 0)} points", 'SUCCESS')
                elif status_type == 'scope_csv_exported':
                    files = data if isinstance(data, list) else [data]
                    for f in files:
                        self.log_message(f"CSV exported: {f}", 'SUCCESS')
                elif status_type == 'scope_plot_generated':
                    files = data if isinstance(data, list) else [data]
                    for f in files:
                        self.log_message(f"Plot generated: {f}", 'SUCCESS')
                elif status_type == 'scope_full_auto_done':
                    res = data
                    self.log_message("Full automation complete", 'SUCCESS')
                    self.log_message(f"Screenshot: {Path(res['screenshot']).name}", 'SUCCESS')
                    for f in res.get('csv', []):
                        self.log_message(f"CSV: {Path(f).name}", 'SUCCESS')
                    for f in res.get('plot', []):
                        self.log_message(f"Plot: {Path(f).name}", 'SUCCESS')
                elif status_type == 'wgen_config_ok':
                    self.log_message(str(data), 'SUCCESS')
                elif status_type == 'wgen_enable_ok':
                    self.log_message(str(data), 'SUCCESS')
                elif status_type == 'wgen_config_read':
                    cfg = data
                    msg = (
                        f"WGEN{cfg.get('generator')}: {cfg.get('function')} "
                        f"{cfg.get('frequency')}Hz {cfg.get('amplitude')}Vpp {cfg.get('offset')}V "
                        f"Output={'ON' if cfg.get('output_enabled') else 'OFF'}"
                    )
                    self.log_message(msg, 'INFO')
                elif status_type == 'wgen_error':
                    self.log_message(f"WGEN error: {data}", 'ERROR')
                elif status_type == 'ramp_progress':
                    self.ramp_progress_label.config(text=data)
                elif status_type == 'ramp_progress_bar':
                    self.ramp_progress_bar['value'] = data
                elif status_type == 'ramp_complete':
                    self.log_message("Voltage Ramping Test Complete", 'SUCCESS')
                    self.ramping_active = False
                    self.start_ramping_btn.config(state=tk.NORMAL)
                    self.stop_ramping_btn.config(state=tk.DISABLED)
                    messagebox.showinfo("Test Complete", 
                                      "Voltage ramping test completed successfully")
                elif status_type == 'ramp_error':
                    self.log_message(f"Ramping Test Error: {data}", 'ERROR')
                    self.ramping_active = False
                    self.start_ramping_btn.config(state=tk.NORMAL)
                    self.stop_ramping_btn.config(state=tk.DISABLED)
                    messagebox.showerror("Test Error", f"Ramping test failed:\n{data}")
                    
        except queue.Empty:
            pass
        
        self.root.after(100, self._check_status_updates)
    
    # Cleanup
    def _on_closing(self):
        if self.ramping_active:
            result = messagebox.askyesno(
                "Test in Progress",
                "A voltage ramping test is currently running.\n\n"
                "Do you want to stop the test and exit?"
            )
            if not result:
                return
            self._stop_ramping_test()
        
        result = messagebox.askyesno(
            "Confirm Exit",
            "Are you sure you want to exit?\n\n"
            "All instrument connections will be closed."
        )
        
        if result:
            self.log_message("Cleaning up...", 'INFO')
            self._disconnect_all_instruments()
            self.log_message("Cleanup complete", 'SUCCESS')
            self.root.destroy()
    
    def run(self):
        self.log_message("Unified Automation System Started", 'SUCCESS')
        self.log_message("Ready to connect instruments", 'INFO')
        self.root.mainloop()


def main():
    try:
        app = UnifiedAutomationGUI()
        app.run()
    except KeyboardInterrupt:
        logger.info("Application interrupted by user")
    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
        messagebox.showerror("Fatal Error", f"Application failed to start:\n{e}")
    finally:
        logger.info("Application shutdown complete")


if __name__ == "__main__":
    main()
