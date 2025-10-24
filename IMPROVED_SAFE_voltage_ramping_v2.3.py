#!/usr/bin/env python3

"""
SAFE Voltage Ramping with IMPROVED UI v2.3
==========================================

Conservative rebase of v2.2 with small, safe improvements:
- Adds waveform selection (Sine, Square, Triangle, Ramp Up, Ramp Down)
- Removes the previously-broken fast-mode option (timing always respects profile)
- Improves oscilloscope screenshot saving with directory creation and clear errors

Preserves v2.2 safety logic and UI structure to minimize regressions.

Author: Rebase of v2.2
Version: 2.3
"""

import sys
import os
import logging
import time
import threading
import queue
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Tuple
import math

# Configure matplotlib for thread safety BEFORE importing pyplot
import matplotlib
matplotlib.use('Agg')  # Non-GUI backend for threading
import matplotlib.pyplot as plt
plt.ioff()  # Turn off interactive mode

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import numpy as np
import pandas as pd

# Import instrument control modules. If unavailable, allow UI to run for development.
try:
    from instrument_control.keithley_power_supply import KeithleyPowerSupply
    from instrument_control.keithley_dmm import KeithleyDMM6500, MeasurementFunction
    from instrument_control.keysight_oscilloscope import KeysightDSOX6004A
except Exception as e:
    print(f"Warning: instrument modules unavailable: {e}")

# Optional helper from the oscilloscope automation app (reuse its data-acquisition/export/plot helpers)
try:
    from oscilloscope_automation_main import OscilloscopeDataAcquisition
except Exception:
    OscilloscopeDataAcquisition = None


class WaveformGenerator:
    """Simple waveform generator supporting a small set of user-selectable waveforms.

    Produces a list of (time_from_start_seconds, voltage) tuples.
    """

    TYPES = ["Sine", "Square", "Triangle", "Ramp Up", "Ramp Down"]

    def __init__(self, waveform_type: str = "Sine", target_voltage: float = 3.0,
                 cycles: int = 3, points_per_cycle: int = 50, cycle_duration: float = 8.0):
        self.waveform_type = waveform_type if waveform_type in self.TYPES else "Sine"
        self.target_voltage = min(float(target_voltage), 5.0)
        self.cycles = max(1, int(cycles))
        self.points_per_cycle = max(1, int(points_per_cycle))
        self.cycle_duration = float(cycle_duration)
        self.logger = logging.getLogger('WaveformGenerator')

    def generate(self) -> List[Tuple[float, float]]:
        profile: List[Tuple[float, float]] = []
        for cycle in range(self.cycles):
            for point in range(self.points_per_cycle):
                pos = point / max(1, self.points_per_cycle - 1) if self.points_per_cycle > 1 else 0.0
                t = cycle * self.cycle_duration + pos * self.cycle_duration

                if self.waveform_type == "Sine":
                    # Half-sine from 0..pi (positive only)
                    v = math.sin(pos * math.pi) * self.target_voltage
                elif self.waveform_type == "Square":
                    v = self.target_voltage if pos < 0.5 else 0.0
                elif self.waveform_type == "Triangle":
                    if pos < 0.5:
                        v = (pos * 2.0) * self.target_voltage
                    else:
                        v = (2.0 - pos * 2.0) * self.target_voltage
                elif self.waveform_type == "Ramp Up":
                    v = pos * self.target_voltage
                elif self.waveform_type == "Ramp Down":
                    v = (1.0 - pos) * self.target_voltage
                else:
                    v = 0.0

                # Enforce safety limits
                v = max(0.0, min(v, 5.0))
                profile.append((round(t, 6), round(v, 6)))

        self.logger.info(f"Generated {len(profile)} points for waveform '{self.waveform_type}' (target {self.target_voltage}V)")
        return profile


class ImprovedDataManager:
    """Data manager with user-configurable save locations and simple exports."""

    def __init__(self):
        self.logger = logging.getLogger('ImprovedDataManager')
        self.base_dir = Path.cwd()
        self.data_dir = self.base_dir / "voltage_ramp_data"
        self.graphs_dir = self.base_dir / "voltage_ramp_graphs"
        self.screenshots_dir = self.base_dir / "oscilloscope_screenshots"

        for d in (self.data_dir, self.graphs_dir, self.screenshots_dir):
            d.mkdir(parents=True, exist_ok=True)

        self.voltage_data: List[dict] = []
        self.timestamps: List[datetime] = []
        self.waveform_type = 'Sine'

    def add_data_point(self, timestamp: datetime, set_voltage: float, measured_voltage: float, cycle_number: int, point_in_cycle: int):
        dp = {
            'timestamp': timestamp,
            'set_voltage': set_voltage,
            'measured_voltage': measured_voltage,
            'cycle_number': cycle_number,
            'point_in_cycle': point_in_cycle,
            'waveform_type': self.waveform_type,
            'time_from_start': (timestamp - self.timestamps[0]).total_seconds() if self.timestamps else 0.0
        }
        self.voltage_data.append(dp)
        self.timestamps.append(timestamp)

    def export_to_csv(self, custom_filename: Optional[str] = None) -> str:
        if not self.voltage_data:
            raise ValueError('No data to export')

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        fn = custom_filename or f"voltage_ramp_{self.waveform_type.lower()}_{ts}.csv"
        fp = self.data_dir / fn

        df = pd.DataFrame(self.voltage_data)
        with open(fp, 'w') as f:
            f.write(f"# Voltage Ramping Data - {self.waveform_type}\n")
            f.write(f"# Generated: {datetime.now().isoformat()}\n\n")
        df.to_csv(fp, mode='a', index=False)
        self.logger.info(f"Exported CSV: {fp}")
        return str(fp)

    def generate_thread_safe_graph(self, custom_title: Optional[str] = None, save_location: Optional[str] = None) -> str:
        if not self.voltage_data:
            raise ValueError('No data to plot')

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        fn = f"{self.waveform_type.lower()}_voltage_vs_time_{ts}.png"
        if save_location and os.path.exists(save_location):
            fp = Path(save_location) / fn
        else:
            fp = self.graphs_dir / fn

        fig, ax = plt.subplots(figsize=(12, 7))
        times = [d['time_from_start'] for d in self.voltage_data]
        set_v = [d['set_voltage'] for d in self.voltage_data]
        meas_v = [d['measured_voltage'] for d in self.voltage_data]

        ax.plot(times, set_v, label='Set Voltage', color='tab:blue')
        ax.plot(times, meas_v, label='Measured Voltage', color='tab:red')
        ax.set_xlabel('Time (s)')
        ax.set_ylabel('Voltage (V)')
        ax.set_title(custom_title or f"{self.waveform_type} Wave Voltage Ramping")
        ax.grid(True, ls='--', alpha=0.4)
        ax.legend()
        plt.tight_layout()
        plt.savefig(fp, dpi=300)
        plt.close(fig)
        self.logger.info(f"Saved graph: {fp}")
        return str(fp)

    def generate_psu_graph(self, custom_title: Optional[str] = None, save_location: Optional[str] = None) -> str:
        """Generate graph showing PSU set voltages over time."""
        if not self.voltage_data:
            raise ValueError('No data to plot')

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        fn = f"{self.waveform_type.lower()}_psu_set_voltage_{ts}.png"
        if save_location and os.path.exists(save_location):
            fp = Path(save_location) / fn
        else:
            fp = self.graphs_dir / fn

        fig, ax = plt.subplots(figsize=(12, 6))
        times = [d['time_from_start'] for d in self.voltage_data]
        set_v = [d['set_voltage'] for d in self.voltage_data]

        ax.plot(times, set_v, label='PSU Set Voltage', color='tab:blue', linewidth=2)
        ax.set_xlabel('Time (s)')
        ax.set_ylabel('Voltage (V)')
        ax.set_title(custom_title or f'PSU Set Voltage vs Time ({self.waveform_type})')
        ax.grid(True, ls='--', alpha=0.4)
        ax.legend()
        plt.tight_layout()
        plt.savefig(fp, dpi=300)
        plt.close(fig)
        self.logger.info(f"Saved PSU graph: {fp}")
        return str(fp)

    def generate_dmm_graph(self, custom_title: Optional[str] = None, save_location: Optional[str] = None) -> str:
        """Generate graph showing DMM measured voltages over time."""
        if not self.voltage_data:
            raise ValueError('No data to plot')

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        fn = f"{self.waveform_type.lower()}_dmm_measured_voltage_{ts}.png"
        if save_location and os.path.exists(save_location):
            fp = Path(save_location) / fn
        else:
            fp = self.graphs_dir / fn

        fig, ax = plt.subplots(figsize=(12, 6))
        times = [d['time_from_start'] for d in self.voltage_data]
        meas_v = [d['measured_voltage'] for d in self.voltage_data]

        ax.plot(times, meas_v, label='DMM Measured Voltage', color='tab:red', linewidth=2)
        ax.set_xlabel('Time (s)')
        ax.set_ylabel('Voltage (V)')
        ax.set_title(custom_title or f'DMM Measured Voltage vs Time ({self.waveform_type})')
        ax.grid(True, ls='--', alpha=0.4)
        ax.legend()
        plt.tight_layout()
        plt.savefig(fp, dpi=300)
        plt.close(fig)
        self.logger.info(f"Saved DMM graph: {fp}")
        return str(fp)

    def generate_comparator_graph(self, custom_title: Optional[str] = None, save_location: Optional[str] = None) -> str:
        """Generate comparator graph: overlay of set vs measured and an error subplot."""
        if not self.voltage_data:
            raise ValueError('No data to plot')

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        fn = f"{self.waveform_type.lower()}_comparator_{ts}.png"
        if save_location and os.path.exists(save_location):
            fp = Path(save_location) / fn
        else:
            fp = self.graphs_dir / fn

        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 8), height_ratios=[3, 1])
        times = [d['time_from_start'] for d in self.voltage_data]
        set_v = [d['set_voltage'] for d in self.voltage_data]
        meas_v = [d['measured_voltage'] for d in self.voltage_data]
        errors = [abs(s - m) for s, m in zip(set_v, meas_v)]

        ax1.plot(times, set_v, label='PSU Set Voltage', color='tab:blue', linewidth=2)
        ax1.plot(times, meas_v, label='DMM Measured Voltage', color='tab:red', linewidth=1.5, alpha=0.9)
        ax1.set_ylabel('Voltage (V)')
        ax1.set_title(custom_title or f'Comparator: Set vs Measured ({self.waveform_type})')
        ax1.grid(True, ls='--', alpha=0.3)
        ax1.legend()

        ax2.plot(times, errors, color='orange', linewidth=1.5)
        ax2.set_xlabel('Time (s)')
        ax2.set_ylabel('Absolute Error (V)')
        ax2.grid(True, ls='--', alpha=0.3)

        plt.tight_layout()
        plt.savefig(fp, dpi=300)
        plt.close(fig)
        self.logger.info(f"Saved comparator graph: {fp}")
        return str(fp)

    def generate_combined_graph(self, custom_title: Optional[str] = None, save_location: Optional[str] = None):
        """Generate a single figure containing PSU, DMM and comparator views plus statistics.
        Returns (filepath, stats_dict).
        """
        if not self.voltage_data:
            self.logger.warning('No voltage data for combined graph')
            return (str(self.graphs_dir / 'no_data.png'), {})

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        fn = f"{self.waveform_type.lower()}_combined_{ts}.png"
        if save_location and os.path.exists(save_location):
            fp = Path(save_location) / fn
        else:
            fp = self.graphs_dir / fn

        times = [d['time_from_start'] for d in self.voltage_data]
        set_v = np.array([d['set_voltage'] for d in self.voltage_data], dtype=float)
        meas_v = np.array([d['measured_voltage'] for d in self.voltage_data], dtype=float)
        errors = np.abs(set_v - meas_v)

        stats = {
            'points': len(times),
            'set_mean': float(np.mean(set_v)),
            'set_std': float(np.std(set_v)),
            'set_min': float(np.min(set_v)),
            'set_max': float(np.max(set_v)),
            'meas_mean': float(np.mean(meas_v)),
            'meas_std': float(np.std(meas_v)),
            'meas_min': float(np.min(meas_v)),
            'meas_max': float(np.max(meas_v)),
            'error_mean': float(np.mean(errors)),
            'error_std': float(np.std(errors)),
            'error_max': float(np.max(errors)),
            'error_rms': float(np.sqrt(np.mean(errors**2)))
        }

        plt.style.use('seaborn-v0_8-darkgrid')
        fig = plt.figure(figsize=(14, 11))
        # Reserve right column for info (top) and table (bottom)
        gs = fig.add_gridspec(3, 3, height_ratios=[1, 1, 1.2], width_ratios=[4, 0.12, 1.2], hspace=0.42, wspace=0.28)
        ax0 = fig.add_subplot(gs[0, 0])                 # PSU set
        ax1 = fig.add_subplot(gs[1, 0])                 # DMM measured
        ax2 = fig.add_subplot(gs[2, 0])                 # Overlay comparator
        ax_hist = fig.add_subplot(gs[2, 1])             # small histogram column
        ax_info = fig.add_subplot(gs[0, 2])             # metadata/info box
        ax_table = fig.add_subplot(gs[1:, 2])           # stats table column (rows 1..)
        ax_info.axis('off')
        ax_table.axis('off')

        # Suptitle with metadata
        suptitle = custom_title or f'Combined Voltage Report ({self.waveform_type})'
        fig.suptitle(f"{suptitle} — {self.waveform_type} — {ts}", fontsize=14, fontweight='bold')

        # Common formatting
        base_font = {'fontname': 'DejaVu Sans', 'fontsize': 10}
        ax0.plot(times, set_v, label='PSU Set', color='#1f77b4', lw=2)
        ax0.scatter(times, set_v, s=10, color='#1f77b4', alpha=0.65)
        ax0.set_ylabel('Voltage (V)', **base_font)
        ax0.set_title('PSU: Set Voltage vs Time', **base_font)
        ax0.grid(True, ls='--', alpha=0.28)
        ax0.legend(loc='upper right', frameon=True, fontsize=9)
        ax0.tick_params(which='both', width=1)
        ax0.minorticks_on()

        ax1.plot(times, meas_v, label='DMM Measured', color='#d62728', lw=2)
        ax1.scatter(times, meas_v, s=10, color='#d62728', alpha=0.65)
        ax1.set_ylabel('Voltage (V)', **base_font)
        ax1.set_title('DMM: Measured Voltage vs Time', **base_font)
        ax1.grid(True, ls='--', alpha=0.28)
        ax1.legend(loc='upper right', frameon=True, fontsize=9)
        ax1.tick_params(which='both', width=1)
        ax1.minorticks_on()

        # Overlay with error shading
        ax2.plot(times, set_v, label='PSU Set', color='#1f77b4', lw=2)
        ax2.plot(times, meas_v, label='DMM Measured', color='#d62728', lw=1.5, alpha=0.95)
        ax2.fill_between(times, np.minimum(set_v, meas_v), np.maximum(set_v, meas_v), color='#ff7f0e', alpha=0.18, label='Absolute Error')
        ax2.set_xlabel('Time (s)', **base_font)
        ax2.set_ylabel('Voltage (V)', **base_font)
        ax2.set_title('Comparator: Set vs Measured (overlay)', **base_font)
        ax2.grid(True, ls='--', alpha=0.28)
        ax2.legend(loc='upper right', framealpha=0.9, fontsize=9)
        ax2.tick_params(which='both', width=1)
        ax2.minorticks_on()

        # Annotate the largest error point and also mark 95th percentile
        max_idx = int(np.argmax(errors)) if len(errors) else 0
        try:
            t_max = times[max_idx]
            err_max = errors[max_idx]
            ax2.annotate(f'Max err {err_max:.4f} V', xy=(t_max, meas_v[max_idx]), xytext=(t_max, max(set_v[max_idx], meas_v[max_idx]) + 0.12),
                         arrowprops=dict(arrowstyle='->', color='black'), fontsize=9, bbox=dict(boxstyle='round', fc='white', alpha=0.85))
            p95 = np.percentile(errors, 95)
            ax2.axhline(p95, color='#9467bd', ls='--', lw=1.25, alpha=0.8)
            ax2.text(0.98, 0.02, f'95% err ≤ {p95:.4f} V', transform=ax2.transAxes, ha='right', va='bottom', fontsize=9, bbox=dict(boxstyle='round', fc='white', alpha=0.6))
        except Exception:
            pass

        # Histogram of errors (vertical) in small column
        ax_hist.hist(errors, bins=30, orientation='horizontal', color='#ff7f0e', alpha=0.9)
        ax_hist.set_xlabel('Count', **base_font)
        ax_hist.set_ylabel('Absolute Error (V)', **base_font)
        ax_hist.grid(False)
        ax_hist.tick_params(axis='y', labelsize=9)

        # Regression and scatter inset on ax2 (set vs measured)
        inset = ax2.inset_axes([0.02, 0.52, 0.30, 0.40])
        inset.scatter(set_v, meas_v, s=18, c='#6a3d9a', alpha=0.55, edgecolors='none')
        inset.set_xlabel('Set V (V)', fontsize=9)
        inset.set_ylabel('Meas V (V)', fontsize=9)
        inset.grid(True, ls=':', alpha=0.25)
        # 1:1 line and robust linear fit
        mn = float(min(np.min(set_v), np.min(meas_v)))
        mx = float(max(np.max(set_v), np.max(meas_v)))
        inset.plot([mn, mx], [mn, mx], color='gray', lw=1, ls='--')
        try:
            p = np.polyfit(set_v, meas_v, 1)
            slope, intercept = p[0], p[1]
            fit_x = np.array([mn, mx])
            fit_y = slope * fit_x + intercept
            inset.plot(fit_x, fit_y, color='#2ca02c', lw=1.25)
            r = float(np.corrcoef(set_v, meas_v)[0, 1])
            inset.text(0.05, 0.06, f'y={slope:.3f}x+{intercept:.3f}\nR={r:.4f}', transform=inset.transAxes, fontsize=9, va='bottom', bbox=dict(boxstyle='round', fc='white', alpha=0.7))
        except Exception:
            pass

        # # Rolling mean inset for time series stability (window = 5% of points, min 3)
        # try:
        #     w = max(3, int(len(set_v) * 0.05))
        #     roll_set = pd.Series(set_v).rolling(window=w, center=True, min_periods=1).mean()
        #     roll_meas = pd.Series(meas_v).rolling(window=w, center=True, min_periods=1).mean()
        #     inset_ts = ax1.inset_axes([0.60, 0.05, 0.35, 0.35])
        #     inset_ts.plot(times, roll_set, color='#1f77b4', lw=1.5, label='Set (rolling)')
        #     inset_ts.plot(times, roll_meas, color='#d62728', lw=1.25, label='Meas (rolling)')
        #     inset_ts.set_title('Rolling mean', fontsize=9)
        #     inset_ts.grid(True, ls=':', alpha=0.25)
        #     inset_ts.legend(fontsize=8)
        # except Exception:
        #     pass

        # Draw cycle boundaries and shaded bands if available in the data
        cycle_starts = []
        try:
            cycle_marks = sorted({d['cycle_number']: None for d in self.voltage_data})
            cycle_starts = [next((dd['time_from_start'] for dd in self.voltage_data if dd['cycle_number'] == c), None) for c in cycle_marks]
            cycle_starts = [cs for cs in cycle_starts if cs is not None]
            if cycle_starts:
                end_time = float(times[-1])
                spans = []
                for idx, cs in enumerate(cycle_starts):
                    start = float(cs)
                    if idx + 1 < len(cycle_starts):
                        end = float(cycle_starts[idx + 1])
                    else:
                        end = end_time
                    spans.append((start, end))
                y_top0 = ax0.get_ylim()[1]
                for i, (s, e) in enumerate(spans):
                    color = '#e9f2ff' if (i % 2 == 0) else '#ffffff'
                    for a in (ax0, ax1, ax2):
                        try:
                            a.axvspan(s, e, facecolor=color, alpha=0.35, linewidth=0)
                        except Exception:
                            pass
                    mid = (s + e) / 2.0
                    label = f'Cycle {i+1}\n{(e - s):.3f}s'
                    try:
                        ax0.text(mid, y_top0 * 0.98, label, ha='center', va='top', fontsize=8, bbox=dict(boxstyle='round', fc='white', alpha=0.6))
                    except Exception:
                        pass
        except Exception:
            cycle_starts = []

        # Prepare styled stats table for ax_table (improved formatting)
        table_data = [
            ['Points', stats['points']],
            ['Set mean (V)', f"{stats['set_mean']:.6f}"],
            ['Set std (V)', f"{stats['set_std']:.6f}"],
            ['Set min (V)', f"{stats['set_min']:.6f}"],
            ['Set max (V)', f"{stats['set_max']:.6f}"],
            ['Meas mean (V)', f"{stats['meas_mean']:.6f}"],
            ['Meas std (V)', f"{stats['meas_std']:.6f}"],
            ['Meas min (V)', f"{stats['meas_min']:.6f}"],
            ['Meas max (V)', f"{stats['meas_max']:.6f}"],
            ['Error mean (V)', f"{stats['error_mean']:.6f}"],
            ['Error std (V)', f"{stats['error_std']:.6f}"],
            ['Error max (V)', f"{stats['error_max']:.6f}"],
            ['Error RMS (V)', f"{stats['error_rms']:.6f}"]
        ]

        # Build a table with nicer cell colors
        table = ax_table.table(cellText=table_data, colWidths=[0.8, 0.7], cellLoc='left', loc='center')
        table.auto_set_font_size(False)
        table.set_fontsize(9)
        table.scale(1, 1.2)
        for (row, col), cell in table.get_celld().items():
            if col == 0:
                cell.set_facecolor('#f0f4f8')
            else:
                cell.set_facecolor('white')

        # Metadata / info box content placed in ax_info to avoid overlap with table
        try:
            total_points = stats.get('points', len(times))
            cycles_count = len(set(d.get('cycle_number', 0) for d in self.voltage_data))
            try:
                from collections import Counter
                counts = Counter(d.get('point_in_cycle', 0) for d in self.voltage_data)
                points_per_cycle_est = max(counts.values()) if counts else 0
            except Exception:
                points_per_cycle_est = 0

            cycle_dur = None
            if len(cycle_starts) >= 2:
                cycle_dur = float(cycle_starts[1] - cycle_starts[0])
            elif cycles_count > 0:
                cycle_dur = float(times[-1]) / max(1, cycles_count)
            sample_interval = None
            try:
                diffs = np.diff(times)
                sample_interval = float(np.median(diffs)) if len(diffs) else 0.0
            except Exception:
                sample_interval = 0.0

            info_lines = [
                f"Waveform: {self.waveform_type}",
                f"Points: {total_points}",
                f"Cycles: {cycles_count}",
                f"Pts/cycle (est): {points_per_cycle_est}",
                f"Cycle dur (s): {cycle_dur:.3f}" if cycle_dur is not None else "Cycle dur (s): N/A",
                f"Sample Δt (s): {sample_interval:.4f}",
                f"Generated: {ts}"
            ]
            info_text = "\n".join(info_lines)
            ax_info.text(0.01, 0.98, info_text, va='top', ha='left', fontsize=9, family='DejaVu Sans', bbox=dict(boxstyle='round', fc='white', alpha=0.9))
        except Exception:
            pass

        # Footer with filename and path
        footer = f"Saved: {fp}"
        fig.text(0.01, 0.01, footer, fontsize=8, color='gray')

        plt.tight_layout(rect=[0, 0.02, 1, 0.96])
        plt.savefig(fp, dpi=300)
        plt.close(fig)
        self.logger.info(f"Saved combined graph: {fp}")
        return (str(fp), stats)

    def clear_data(self):
        self.voltage_data.clear()
        self.timestamps.clear()


class ImprovedSafeVoltageRampingGUI:
    """GUI re-based on v2.2 with localized waveform and screenshot fixes."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title('SAFE Voltage Ramping - v2.3')
        self.root.geometry('1000x650')
        self.root.minsize(800, 500)
        # Configure grid expansion
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

    

        # Create main frame (store as attribute so other methods can parent widgets to it)
        self.main_frame = ttk.Frame(self.root, padding="5")
        self.main_frame.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)

        # Instruments
        self.power_supply: Optional[KeithleyPowerSupply] = None
        self.dmm: Optional[KeithleyDMM6500] = None
        self.oscilloscope: Optional[KeysightDSOX6004A] = None

        # Data manager & waveform
        self.data_manager = ImprovedDataManager()
        self.waveform_type_var = tk.StringVar(value='Sine')
        # Allow selecting which PSU channel is active and per-channel waveform assignment
        self.psu_active_channel_var = tk.IntVar(value=1)
    # Support up to 4 channels for assignment (UI only; driver controls which channels actually exist)
    # Leave per-channel selection empty by default so the global waveform selection is used unless overridden.
        self.psu_channel_waveform_vars = {ch: tk.StringVar(value='') for ch in range(1, 5)}

        # GUI variables
        self.target_voltage_var = tk.DoubleVar(value=3.0)
        self.cycles_var = tk.IntVar(value=3)
        self.points_per_cycle_var = tk.IntVar(value=50)
        self.cycle_duration_var = tk.DoubleVar(value=8.0)
        self.current_limit_var = tk.DoubleVar(value=1.0)
        self.nplc_var = tk.DoubleVar(value=1.0)
        self.psu_settle_var = tk.DoubleVar(value=0.05)

        self.psu_visa_var = tk.StringVar(value='USB0::0x05E6::0x2230::805224014806770001::INSTR')
        self.dmm_visa_var = tk.StringVar(value='USB0::0x05E6::0x6500::04561287::INSTR')
        self.scope_visa_var = tk.StringVar(value='USB0::0x0957::0x1780::MY65220169::INSTR')

        self.data_path_var = tk.StringVar(value=str(self.data_manager.data_dir))
        self.graphs_path_var = tk.StringVar(value=str(self.data_manager.graphs_dir))
        self.screenshots_path_var = tk.StringVar(value=str(self.data_manager.screenshots_dir))

        self.progress_var = tk.DoubleVar(value=0.0)
        self.eta_display_var = tk.StringVar(value='ETA: --')
        self.point_interval_var = tk.StringVar(value='Point interval: -- s')
        self.safety_status_var = tk.StringVar(value='System Safe - Ready for operation')
        self.current_operation_var = tk.StringVar(value='Ready - Connect instruments and configure settings')

        # Logging
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        self.logger = logging.getLogger('ImprovedSafeVoltageRamping')

        # Build UI
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure_professional_styles()
        self.create_scrollable_main_frame()
        self.create_title_section()
        # toolbar uses header created by create_title_section
        try:
            self.create_toolbar()
        except Exception:
            pass
        self.create_emergency_section()
        self.create_connection_section()
        self.create_configuration_section()
        # Oscilloscope controls (integrated from oscilloscope helper)
        try:
            self.create_oscilloscope_section()
        except Exception:
            pass
        self.create_save_location_section()
        self.create_operations_section()
        self.create_status_section()
        # persistent status bar at bottom of main window
        try:
            self.create_status_bar()
        except Exception:
            pass

        # Threading
        self.status_queue = queue.Queue()
        self.operation_thread = None

        # Periodic status check
        self.root.after(100, self.check_status_updates)
        self.root.protocol('WM_DELETE_WINDOW', self.safe_shutdown)

    # --- UI helper methods ---
    def configure_professional_styles(self):
        # Lighter modern color palette and tactile button styles
        accent = '#2563eb'  # medium blue accent
        accent_secondary = '#0ea5a4'  # teal accent for subtle highlights
        panel_bg = '#f7fafc'  # very light gray panel background
        main_bg = '#ffffff'   # white window background
        subtle = '#6b7280'    # neutral text color

        # Base styles
        self.style.configure('.', background=main_bg, foreground='#0f172a', font=('Segoe UI', 10))
        self.style.configure('TFrame', background=panel_bg)
        self.style.configure('TLabel', background=panel_bg, foreground='#0f172a')
        self.style.configure('TLabelframe', background=panel_bg, foreground='#0f172a')
        self.style.configure('TLabelframe.Label', font=('Segoe UI', 11, 'bold'))

        # Buttons
        self.style.configure('Emergency.TButton', font=('Segoe UI', 12, 'bold'), background='#ef4444', foreground='white')
        self.style.map('Emergency.TButton', background=[('active', '#dc2626')])
        self.style.configure('Connect.TButton', font=('Segoe UI', 9, 'bold'), background=accent, foreground='white')
        self.style.map('Connect.TButton', background=[('active', '#1e40af')])
        self.style.configure('Disconnect.TButton', font=('Segoe UI', 9, 'bold'), background='#ef4444', foreground='white')
        self.style.map('Disconnect.TButton', background=[('active', '#dc2626')])
        self.style.configure('Primary.TButton', font=('Segoe UI', 10, 'bold'), background=accent, foreground='white')
        self.style.map('Primary.TButton', background=[('active', '#1d4ed8')])
        self.style.configure('Success.TButton', font=('Segoe UI', 9, 'bold'), background=accent_secondary, foreground='white')
        self.style.map('Success.TButton', background=[('active', '#059669')])

        # Compact inputs
        self.style.configure('TCombobox', padding=4)
        self.style.configure('TEntry', padding=4)
        self.style.configure('TSpinbox', padding=3)

    def create_scrollable_main_frame(self):
        # Parent the canvas to the main_frame to avoid mixing pack/grid on root
        self.canvas = tk.Canvas(self.main_frame, bg='#f0f2f5')
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient='vertical', command=self.canvas.yview)
        # The scrollable_frame is placed inside the canvas. Save the window id so we can
        # resize the embedded frame to match the canvas width — this avoids an empty
        # right-side gap when the window is resized.
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        self.window_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor='nw')

        # When the canvas changes size, force the internal frame to the same width so
        # widgets using grid(column=0, sticky='ew') will expand correctly.
        def _on_canvas_config(event):
            try:
                self.canvas.itemconfig(self.window_id, width=event.width)
            except Exception:
                pass

        self.canvas.bind('<Configure>', _on_canvas_config)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        # Use grid inside the main_frame so we only use grid on the root window
        self.canvas.grid(row=0, column=0, sticky='nsew')
        self.scrollbar.grid(row=0, column=1, sticky='ns')
        try:
            self.main_frame.columnconfigure(0, weight=1)
            self.main_frame.rowconfigure(0, weight=1)
        except Exception:
            pass
        self.canvas.bind('<MouseWheel>', self._on_mousewheel)
        self.root.bind('<MouseWheel>', self._on_mousewheel)

        # Ensure the scrollable_frame's column 0 expands so content fills the available width
        try:
            self.scrollable_frame.columnconfigure(0, weight=1)
        except Exception:
            pass

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

    def create_title_section(self):
        # Sleek header with accent strip
        self.header = ttk.Frame(self.scrollable_frame, padding=(12, 8))
        self.header.grid(row=0, column=0, sticky='ew', pady=(0, 10))
        self.header.columnconfigure(1, weight=1)

        # small icon block (colored square) with lighter background
        icon_bg = '#f1f5f9'
        icon_fill = '#0ea5a4'
        icon = tk.Canvas(self.header, width=56, height=56, highlightthickness=0, bg=icon_bg)
        icon.create_rectangle(6, 6, 50, 50, fill=icon_fill, outline='')
        icon.grid(row=0, column=0, rowspan=2, padx=(0, 12))

        title = ttk.Label(self.header, text='SAFE VOLTAGE RAMPING', font=('Segoe UI', 18, 'bold'), foreground='#0f172a')
        title.grid(row=0, column=1, sticky='w')
        subtitle = ttk.Label(self.header, text='Multi-waveform • PSU Safety Priority • Thread-safe graphs', font=('Segoe UI', 10), foreground='#475569')
        subtitle.grid(row=1, column=1, sticky='w')

        ver = ttk.Label(self.header, text='v2.3', font=('Segoe UI', 9, 'bold'), foreground='#6b7280')
        ver.grid(row=0, column=2, sticky='e')

        # thin accent separator inside header to avoid shifting grid rows below
        sep = tk.Frame(self.header, height=4, bg='#e6f6f4')
        sep.grid(row=2, column=0, columnspan=4, sticky='ew', pady=(8, 0))

    def create_toolbar(self):
        """Compact toolbar placed in the header for quick actions."""
        try:
            tb = ttk.Frame(self.header)
            tb.grid(row=0, column=3, rowspan=2, sticky='e', padx=(8, 0))
            ttk.Button(tb, text='Connect All', command=self.connect_all_instruments, style='Connect.TButton').pack(side=tk.LEFT, padx=4)
            ttk.Button(tb, text='Start', command=self.start_safe_ramping, style='Primary.TButton').pack(side=tk.LEFT, padx=4)
            ttk.Button(tb, text='Stop', command=self.stop_ramping, style='Disconnect.TButton').pack(side=tk.LEFT, padx=4)
            ttk.Button(tb, text='Screenshot', command=self.capture_screenshot).pack(side=tk.LEFT, padx=4)
            ttk.Button(tb, text='Export', command=self.export_data, style='Success.TButton').pack(side=tk.LEFT, padx=4)
            ttk.Button(tb, text='Graph', command=self.generate_safe_graph, style='Success.TButton').pack(side=tk.LEFT, padx=4)
        except Exception:
            pass

    def create_status_bar(self):
        """Persistent status bar docked to the bottom of the main window."""
        try:
            status = ttk.Frame(self.root, padding=(6, 4))
            # Place the status bar in row 1 so it doesn't mix geometry managers on the root
            status.grid(row=1, column=0, sticky='ew')
            # Ensure root row 1 doesn't expand
            try:
                self.root.rowconfigure(1, weight=0)
            except Exception:
                pass
            left = ttk.Label(status, textvariable=self.safety_status_var, font=('Segoe UI', 9, 'bold'), foreground='#0f172a')
            left.grid(row=0, column=0, sticky='w')
            mid = ttk.Label(status, textvariable=self.current_operation_var, font=('Segoe UI', 9), foreground='#374151')
            mid.grid(row=0, column=1, sticky='w', padx=(10, 0))
            right = ttk.Label(status, textvariable=self.eta_display_var, font=('Segoe UI', 9), foreground='#6b7280')
            right.grid(row=0, column=2, sticky='e')
        except Exception:
            pass

    def create_emergency_section(self):
        frame = ttk.LabelFrame(self.scrollable_frame, text='Emergency Controls', padding=15)
        frame.grid(row=1, column=0, sticky='ew', pady=10)
        frame.columnconfigure(0, weight=1)
        self.emergency_stop_btn = ttk.Button(frame, text='EMERGENCY STOP\nImmediate PSU Shutdown', style='Emergency.TButton', command=self.emergency_psu_stop)
        self.emergency_stop_btn.grid(row=0, column=0, pady=10, ipadx=20, ipady=10)
        ttk.Label(frame, textvariable=self.safety_status_var, font=('Segoe UI', 12, 'bold'), foreground='#059669').grid(row=1, column=0, pady=(10, 0))

    def create_connection_section(self):
        frame = ttk.LabelFrame(self.scrollable_frame, text='Instrument Connections', padding=15)
        frame.grid(row=2, column=0, sticky='ew', pady=10)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text='Power Supply VISA:', font=('Segoe UI', 11, 'bold')).grid(row=0, column=0, sticky='w', padx=(0, 10))
        ttk.Entry(frame, textvariable=self.psu_visa_var, font=('Consolas', 10), width=55).grid(row=0, column=1, sticky='ew', padx=(0, 10))
        self.psu_connect_btn = ttk.Button(frame, text='Connect', command=self.connect_psu, style='Connect.TButton')
        self.psu_connect_btn.grid(row=0, column=2, padx=(0, 5))
        self.psu_disconnect_btn = ttk.Button(frame, text='Disconnect', command=self.disconnect_psu, style='Disconnect.TButton', state='disabled')
        self.psu_disconnect_btn.grid(row=0, column=3, padx=(0, 10))
        self.psu_status_var = tk.StringVar(value='Disconnected')
        ttk.Label(frame, textvariable=self.psu_status_var, font=('Segoe UI', 10, 'bold')).grid(row=0, column=4)

        ttk.Label(frame, text='DMM VISA:', font=('Segoe UI', 11, 'bold')).grid(row=1, column=0, sticky='w', padx=(0, 10), pady=(15, 0))
        ttk.Entry(frame, textvariable=self.dmm_visa_var, font=('Consolas', 10), width=55).grid(row=1, column=1, sticky='ew', padx=(0, 10), pady=(15, 0))
        self.dmm_connect_btn = ttk.Button(frame, text='Connect', command=self.connect_dmm, style='Connect.TButton')
        self.dmm_connect_btn.grid(row=1, column=2, padx=(0, 5), pady=(15, 0))
        self.dmm_disconnect_btn = ttk.Button(frame, text='Disconnect', command=self.disconnect_dmm, style='Disconnect.TButton', state='disabled')
        self.dmm_disconnect_btn.grid(row=1, column=3, padx=(0, 10), pady=(15, 0))
        self.dmm_status_var = tk.StringVar(value='Disconnected')
        ttk.Label(frame, textvariable=self.dmm_status_var, font=('Segoe UI', 10, 'bold')).grid(row=1, column=4, pady=(15, 0))

        ttk.Label(frame, text='Oscilloscope VISA:', font=('Segoe UI', 11, 'bold')).grid(row=2, column=0, sticky='w', padx=(0, 10), pady=(15, 0))
        ttk.Entry(frame, textvariable=self.scope_visa_var, font=('Consolas', 10), width=55).grid(row=2, column=1, sticky='ew', padx=(0, 10), pady=(15, 0))
        self.scope_connect_btn = ttk.Button(frame, text='Connect', command=self.connect_scope, style='Connect.TButton')
        self.scope_connect_btn.grid(row=2, column=2, padx=(0, 5), pady=(15, 0))
        self.scope_disconnect_btn = ttk.Button(frame, text='Disconnect', command=self.disconnect_scope, style='Disconnect.TButton', state='disabled')
        self.scope_disconnect_btn.grid(row=2, column=3, padx=(0, 10), pady=(15, 0))
        self.scope_status_var = tk.StringVar(value='Disconnected')
        ttk.Label(frame, textvariable=self.scope_status_var, font=('Segoe UI', 10, 'bold')).grid(row=2, column=4, pady=(15, 0))

        control_frame = ttk.Frame(frame)
        control_frame.grid(row=3, column=0, columnspan=5, pady=(20, 0))
        ttk.Button(control_frame, text='Connect All Instruments', command=self.connect_all_instruments, style='Primary.TButton').pack(side=tk.LEFT, padx=(0, 15))
        ttk.Button(control_frame, text='Disconnect All Instruments', command=self.disconnect_all_instruments, style='Disconnect.TButton').pack(side=tk.LEFT, padx=(0, 15))
        ttk.Button(control_frame, text='Test All Connections', command=self.test_all_connections, style='Success.TButton').pack(side=tk.LEFT)

    def create_configuration_section(self):
        config_frame = ttk.LabelFrame(self.scrollable_frame, text='Safe Voltage Ramping Configuration', padding='15')
        config_frame.grid(row=3, column=0, sticky='ew', pady=10)
        for i in range(6):
            config_frame.columnconfigure(i, weight=1)

        ttk.Label(config_frame, text='Target Voltage:', font=('Segoe UI', 11, 'bold')).grid(row=0, column=0, sticky='w', padx=(0, 10))
        voltage_spin = tk.Spinbox(config_frame, from_=0.1, to=4.0, increment=0.1, textvariable=self.target_voltage_var, width=10, format='%.2f', font=('Segoe UI', 10))
        voltage_spin.grid(row=0, column=1, sticky='w', padx=(0, 5))
        ttk.Label(config_frame, text='V (MAX: 4.0V)', font=('Segoe UI', 9), foreground='red').grid(row=0, column=2, sticky='w', padx=(0, 20))

        ttk.Label(config_frame, text='Cycles:', font=('Segoe UI', 11, 'bold')).grid(row=0, column=3, sticky='w', padx=(0, 10))
        cycles_spin = tk.Spinbox(config_frame, from_=1, to=10, increment=1, textvariable=self.cycles_var, width=8, font=('Segoe UI', 10))
        cycles_spin.grid(row=0, column=4, sticky='w', padx=(0, 5))
        ttk.Label(config_frame, text='(MAX: 10)', font=('Segoe UI', 9), foreground='blue').grid(row=0, column=5, sticky='w')

        ttk.Label(config_frame, text='Points per Cycle:', font=('Segoe UI', 11, 'bold')).grid(row=1, column=0, sticky='w', padx=(0, 10), pady=(15, 0))
        points_spin = tk.Spinbox(config_frame, from_=20, to=100, increment=10, textvariable=self.points_per_cycle_var, width=8, font=('Segoe UI', 10), command=self.update_point_interval_label)
        points_spin.grid(row=1, column=1, sticky='w', padx=(0, 5), pady=(15, 0))
        ttk.Label(config_frame, text='(20-100)', font=('Segoe UI', 9), foreground='blue').grid(row=1, column=2, sticky='w', padx=(0, 20), pady=(15, 0))

        ttk.Label(config_frame, text='Cycle Duration:', font=('Segoe UI', 11, 'bold')).grid(row=1, column=3, sticky='w', padx=(0, 10), pady=(15, 0))
        duration_spin = tk.Spinbox(config_frame, from_=3.0, to=30.0, increment=0.5, textvariable=self.cycle_duration_var, width=10, format='%.1f', font=('Segoe UI', 10), command=self.update_point_interval_label)
        duration_spin.grid(row=1, column=4, sticky='w', padx=(0, 5), pady=(15, 0))
        ttk.Label(config_frame, text='seconds (MIN: 3s)', font=('Segoe UI', 9), foreground='blue').grid(row=1, column=5, sticky='w', pady=(15, 0))

        ttk.Label(config_frame, text='PSU Settle (s):', font=('Segoe UI', 11, 'bold')).grid(row=2, column=0, sticky='w', padx=(0, 10), pady=(15, 0))
        ttk.Spinbox(config_frame, from_=0.00, to=0.50, increment=0.01, textvariable=self.psu_settle_var, width=8, format='%.2f', font=('Segoe UI', 10)).grid(row=2, column=1, sticky='w', pady=(15, 0))
        ttk.Label(config_frame, text='(Speed vs Accuracy)', font=('Segoe UI', 9), foreground='gray').grid(row=2, column=5, sticky='w', pady=(15, 0))

        # Waveform selection (new in v2.3)
        ttk.Label(config_frame, text='Waveform:', font=('Segoe UI', 11, 'bold')).grid(row=3, column=0, sticky='w', padx=(0, 10), pady=(15, 0))
        waveform_combo = ttk.Combobox(config_frame, textvariable=self.waveform_type_var, values=WaveformGenerator.TYPES, state='readonly', width=15)
        waveform_combo.grid(row=3, column=1, sticky='w', pady=(15, 0))

        # PSU active channel selector and per-channel waveform assignments
        ttk.Label(config_frame, text='Active PSU Channel:', font=('Segoe UI', 11, 'bold')).grid(row=3, column=2, sticky='w', padx=(10, 10), pady=(15, 0))
        channel_combo = ttk.Combobox(config_frame, values=[1, 2, 3, 4], textvariable=self.psu_active_channel_var, state='readonly', width=6)
        channel_combo.grid(row=3, column=3, sticky='w', pady=(15, 0))

        ttk.Label(config_frame, text='Per-Channel Waveforms (optional):', font=('Segoe UI', 10)).grid(row=4, column=0, sticky='w', padx=(0, 10), pady=(10, 0))
        

        # Calculated interval display
        ttk.Label(config_frame, textvariable=self.point_interval_var, font=('Segoe UI', 10, 'bold'), foreground='#1f2937').grid(row=4, column=0, columnspan=6, sticky='w', pady=(10, 0))

        # Initialize interval display
        self.update_point_interval_label()
        try:
            self.points_per_cycle_var.trace_add('write', lambda *_: self.update_point_interval_label())
            self.cycle_duration_var.trace_add('write', lambda *_: self.update_point_interval_label())
        except Exception:
            pass

    def update_point_interval_label(self):
        try:
            points = max(1, int(self.points_per_cycle_var.get()))
            duration = float(self.cycle_duration_var.get())
            interval = duration / points
            self.point_interval_var.set(f"Point interval: {interval:.3f} s")
        except Exception:
            self.point_interval_var.set("Point interval: -- s")

    def create_oscilloscope_section(self):
        """Add a compact oscilloscope control panel reusing patterns from the osc app."""
        try:
            frame = ttk.LabelFrame(self.scrollable_frame, text='Oscilloscope Controls', padding=12)
            frame.grid(row=7, column=0, sticky='ew', pady=10)
            for i in range(6):
                frame.columnconfigure(i, weight=1)

            # Connection row
            ttk.Label(frame, text='Scope VISA:', font=('Segoe UI', 11, 'bold')).grid(row=0, column=0, sticky='w')
            ttk.Entry(frame, textvariable=self.scope_visa_var, font=('Consolas', 10), width=48).grid(row=0, column=1, sticky='ew', columnspan=3)
            ttk.Button(frame, text='Connect', command=self.connect_scope, style='Connect.TButton').grid(row=0, column=4)
            ttk.Button(frame, text='Disconnect', command=self.disconnect_scope, style='Disconnect.TButton').grid(row=0, column=5)

            # Channel config compact row (checkboxes)
            ttk.Label(frame, text='Channels:', font=('Segoe UI', 10, 'bold')).grid(row=1, column=0, sticky='w', pady=(8,0))
            self.scope_channel_vars = {ch: tk.BooleanVar(value=(ch == 1)) for ch in [1,2,3,4]}
            col = 1
            for ch in [1,2,3,4]:
                ttk.Checkbutton(frame, text=f'Ch{ch}', variable=self.scope_channel_vars[ch]).grid(row=1, column=col, padx=4, pady=(8,0))
                col += 1

            # Operations row
            ttk.Button(frame, text='Screenshot', command=self.scope_capture_screenshot).grid(row=2, column=0, sticky='ew', pady=(10,0))
            ttk.Button(frame, text='Acquire', command=self.scope_acquire_data).grid(row=2, column=1, sticky='ew', pady=(10,0))
            ttk.Button(frame, text='Export CSV', command=self.scope_export_csv).grid(row=2, column=2, sticky='ew', pady=(10,0))
            ttk.Button(frame, text='Generate Plot', command=self.scope_generate_plot).grid(row=2, column=3, sticky='ew', pady=(10,0))
            ttk.Button(frame, text='Full Auto', command=self.scope_full_automation).grid(row=2, column=4, sticky='ew', pady=(10,0))
            ttk.Button(frame, text='Open Folder', command=self.scope_open_output_folder).grid(row=2, column=5, sticky='ew', pady=(10,0))

            # File prefs row
            ttk.Label(frame, text='Data Folder:', font=('Segoe UI', 10, 'bold')).grid(row=3, column=0, sticky='w', pady=(8,0))
            self.scope_data_path_var = tk.StringVar(value=str(self.data_manager.data_dir))
            ttk.Entry(frame, textvariable=self.scope_data_path_var).grid(row=3, column=1, columnspan=2, sticky='ew', pady=(8,0))
            ttk.Label(frame, text='Graphs Folder:', font=('Segoe UI', 10, 'bold')).grid(row=3, column=3, sticky='w', pady=(8,0))
            self.scope_graphs_path_var = tk.StringVar(value=str(self.data_manager.graphs_dir))
            ttk.Entry(frame, textvariable=self.scope_graphs_path_var).grid(row=3, column=4, columnspan=2, sticky='ew', pady=(8,0))

        except Exception:
            pass

    # --- Oscilloscope wrapper methods (threaded, use status_queue) ---
    def scope_capture_screenshot(self):
        def t():
            try:
                if not (self.oscilloscope and getattr(self.oscilloscope, 'is_connected', False)):
                    self.status_queue.put(('error', 'Oscilloscope not connected for screenshot'))
                    return
                res = None
                try:
                    res = self.oscilloscope.capture_screenshot()
                except Exception as e:
                    self.log_message(f'Osc screenshot error: {e}', 'ERROR')
                if res:
                    self.status_queue.put(('screenshot_captured', str(res)))
                else:
                    self.status_queue.put(('error', 'Screenshot failed'))
            except Exception as e:
                self.status_queue.put(('error', f'Screenshot thread error: {e}'))
        threading.Thread(target=t, daemon=True).start()

    def scope_acquire_data(self):
        def t():
            try:
                if not getattr(self, 'osc_data', None):
                    self.status_queue.put(('error', 'Oscilloscope helper not available'))
                    return
                channels = [ch for ch, var in self.scope_channel_vars.items() if var.get()]
                if not channels:
                    self.status_queue.put(('error', 'No oscilloscope channels selected'))
                    return
                all_data = {}
                for ch in channels:
                    data = self.osc_data.acquire_waveform_data(ch)
                    if data:
                        all_data[ch] = data
                        self.log_message(f'Osc Ch{ch} acquired: {data.get("points_count",0)} pts', 'SUCCESS')
                    else:
                        self.log_message(f'Osc Ch{ch} acquisition failed', 'ERROR')
                if all_data:
                    self.last_osc_data = all_data
                    self.status_queue.put(('data_acquired', all_data))
            except Exception as e:
                self.status_queue.put(('error', f'Acquire thread error: {e}'))
        threading.Thread(target=t, daemon=True).start()

    def scope_export_csv(self):
        def t():
            try:
                if not getattr(self, 'osc_data', None):
                    self.status_queue.put(('error', 'Osc data helper missing'))
                    return
                # Use last acquired data if present
                if not hasattr(self, 'last_osc_data') or not self.last_osc_data:
                    self.status_queue.put(('error', 'No oscilloscope data to export'))
                    return
                exported = []
                for ch, data in self.last_osc_data.items():
                    fp = self.osc_data.export_to_csv(data, custom_path=self.scope_data_path_var.get())
                    if fp:
                        exported.append(fp)
                        self.log_message(f'Osc CSV exported: {Path(fp).name}', 'SUCCESS')
                if exported:
                    self.status_queue.put(('csv_exported', exported))
            except Exception as e:
                self.status_queue.put(('error', f'CSV export error: {e}'))
        threading.Thread(target=t, daemon=True).start()

    def scope_generate_plot(self):
        def t():
            try:
                if not getattr(self, 'osc_data', None):
                    self.status_queue.put(('error', 'Osc data helper missing'))
                    return
                if not hasattr(self, 'last_osc_data') or not self.last_osc_data:
                    self.status_queue.put(('error', 'No oscilloscope data to plot'))
                    return
                files = []
                for ch, data in self.last_osc_data.items():
                    fp = self.osc_data.generate_waveform_plot(data, custom_path=self.scope_graphs_path_var.get(), plot_title=self.graph_title_var.get() or None)
                    if fp:
                        files.append(fp)
                        self.log_message(f'Osc plot generated: {Path(fp).name}', 'SUCCESS')
                if files:
                    self.status_queue.put(('plot_generated', files))
            except Exception as e:
                self.status_queue.put(('error', f'Plot error: {e}'))
        threading.Thread(target=t, daemon=True).start()

    def scope_full_automation(self):
        def t():
            try:
                if not getattr(self, 'osc_data', None):
                    self.status_queue.put(('error', 'Osc data helper missing'))
                    return
                channels = [ch for ch, var in self.scope_channel_vars.items() if var.get()]
                if not channels:
                    self.status_queue.put(('error', 'No osc channels selected'))
                    return
                # Run the high-level automation if the helper exposes a similar method
                # Fallback: perform screenshot, acquire, export, plot sequentially
                results = {'screenshot': None, 'csv': [], 'plot': [], 'data': {}}
                try:
                    shot = self.oscilloscope.capture_screenshot() if self.oscilloscope else None
                    results['screenshot'] = shot
                except Exception:
                    pass
                for ch in channels:
                    d = self.osc_data.acquire_waveform_data(ch)
                    if d:
                        results['data'][ch] = d
                        csvf = self.osc_data.export_to_csv(d, custom_path=self.scope_data_path_var.get())
                        if csvf:
                            results['csv'].append(csvf)
                        plotf = self.osc_data.generate_waveform_plot(d, custom_path=self.scope_graphs_path_var.get())
                        if plotf:
                            results['plot'].append(plotf)
                self.last_osc_data = results['data']
                self.status_queue.put(('full_automation_complete', results))
            except Exception as e:
                self.status_queue.put(('error', f'Full automation error: {e}'))
        threading.Thread(target=t, daemon=True).start()

    def scope_open_output_folder(self):
        try:
            path = Path(self.scope_data_path_var.get())
            path.mkdir(parents=True, exist_ok=True)
            os.startfile(str(path))
        except Exception as e:
            self.log_message(f'Open folder error: {e}', 'ERROR')

    def create_save_location_section(self):
        save_frame = ttk.LabelFrame(self.scrollable_frame, text='Custom Save Locations', padding='15')
        save_frame.grid(row=4, column=0, sticky='ew', pady=10)
        save_frame.columnconfigure(1, weight=1)

        ttk.Label(save_frame, text='Data Folder:', font=('Segoe UI', 11, 'bold')).grid(row=0, column=0, sticky='w', padx=(0, 10))
        ttk.Entry(save_frame, textvariable=self.data_path_var, font=('Segoe UI', 10)).grid(row=0, column=1, sticky='ew', padx=(0, 10))
        ttk.Button(save_frame, text='Browse', command=lambda: self.browse_save_location('data')).grid(row=0, column=2)

        ttk.Label(save_frame, text='Graphs Folder:', font=('Segoe UI', 11, 'bold')).grid(row=1, column=0, sticky='w', padx=(0, 10), pady=(15, 0))
        ttk.Entry(save_frame, textvariable=self.graphs_path_var, font=('Segoe UI', 10)).grid(row=1, column=1, sticky='ew', padx=(0, 10), pady=(15, 0))
        ttk.Button(save_frame, text='Browse', command=lambda: self.browse_save_location('graphs')).grid(row=1, column=2, pady=(15, 0))

        ttk.Label(save_frame, text='Screenshots Folder:', font=('Segoe UI', 11, 'bold')).grid(row=2, column=0, sticky='w', padx=(0, 10), pady=(15, 0))
        ttk.Entry(save_frame, textvariable=self.screenshots_path_var, font=('Segoe UI', 10)).grid(row=2, column=1, sticky='ew', padx=(0, 10), pady=(15, 0))
        ttk.Button(save_frame, text='Browse', command=lambda: self.browse_save_location('screenshots')).grid(row=2, column=2, pady=(15, 0))

        ttk.Label(save_frame, text='Graph Title:', font=('Segoe UI', 11, 'bold')).grid(row=3, column=0, sticky='w', padx=(0, 10), pady=(15, 0))
        self.graph_title_var = tk.StringVar(value='SAFE Wave Voltage Ramping')
        ttk.Entry(save_frame, textvariable=self.graph_title_var, font=('Segoe UI', 10)).grid(row=3, column=1, sticky='ew', padx=(0, 10), pady=(15, 0))
        ttk.Button(save_frame, text='Default', command=self.reset_graph_title).grid(row=3, column=2, pady=(15, 0))

    def create_operations_section(self):
        ops_frame = ttk.LabelFrame(self.scrollable_frame, text='Operations Control', padding='15')
        ops_frame.grid(row=5, column=0, sticky='ew', pady=10)
        for i in range(6):
            ops_frame.columnconfigure(i, weight=1)

        self.start_ramping_btn = ttk.Button(ops_frame, text='Start Safe Ramping', command=self.start_safe_ramping, style='Primary.TButton')
        self.start_ramping_btn.grid(row=0, column=0, sticky='ew', padx=3, ipady=8)
        self.stop_ramping_btn = ttk.Button(ops_frame, text='Stop Ramping', command=self.stop_ramping, state='disabled')
        self.stop_ramping_btn.grid(row=0, column=1, sticky='ew', padx=3, ipady=8)
        self.generate_graph_btn = ttk.Button(ops_frame, text='Generate Graph', command=self.generate_safe_graph, style='Success.TButton')
        self.generate_graph_btn.grid(row=0, column=2, sticky='ew', padx=3, ipady=8)
        self.export_csv_btn = ttk.Button(ops_frame, text='Export CSV', command=self.export_data, style='Success.TButton')
        self.export_csv_btn.grid(row=0, column=3, sticky='ew', padx=3, ipady=8)
        self.screenshot_btn = ttk.Button(ops_frame, text='Screenshot', command=self.capture_screenshot)
        self.screenshot_btn.grid(row=0, column=4, sticky='ew', padx=3, ipady=8)
        self.clear_data_btn = ttk.Button(ops_frame, text='Clear Data', command=self.clear_all_data)
        self.clear_data_btn.grid(row=0, column=5, sticky='ew', padx=3, ipady=8)

        progress_frame = ttk.Frame(ops_frame)
        progress_frame.grid(row=1, column=0, columnspan=6, sticky='ew', pady=(15, 10))
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.columnconfigure(1, weight=0)

        ttk.Label(progress_frame, text='Progress:', font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, sticky='w')
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100, mode='determinate')
        self.progress_bar.grid(row=1, column=0, sticky='ew', pady=(5, 0))
        ttk.Label(progress_frame, textvariable=self.eta_display_var, font=('Segoe UI', 10, 'bold')).grid(row=1, column=1, sticky='e', padx=(10, 0))

        # Real-time voltage displays
        voltage_frame = ttk.Frame(ops_frame)
        voltage_frame.grid(row=2, column=0, columnspan=6, pady=(10, 0))
        voltage_frame.columnconfigure(0, weight=1)
        voltage_frame.columnconfigure(1, weight=1)

        set_frame = ttk.Frame(voltage_frame)
        set_frame.grid(row=0, column=0, sticky='ew', padx=(0, 20))
        ttk.Label(set_frame, text='Set Voltage:', font=('Segoe UI', 12, 'bold')).pack(side=tk.LEFT)
        self.set_voltage_display = tk.StringVar(value='0.000 V')
        ttk.Label(set_frame, textvariable=self.set_voltage_display, font=('Consolas', 14, 'bold'), foreground='#2563eb').pack(side=tk.LEFT, padx=(10, 0))

        measured_frame = ttk.Frame(voltage_frame)
        measured_frame.grid(row=0, column=1, sticky='ew')
        ttk.Label(measured_frame, text='Measured Voltage:', font=('Segoe UI', 12, 'bold')).pack(side=tk.LEFT)
        self.measured_voltage_display = tk.StringVar(value='0.000 V')
        ttk.Label(measured_frame, textvariable=self.measured_voltage_display, font=('Consolas', 14, 'bold'), foreground='#dc2626').pack(side=tk.LEFT, padx=(10, 0))

    def create_status_section(self):
        status_frame = ttk.LabelFrame(self.scrollable_frame, text='System Status & Activity Log', padding='15')
        status_frame.grid(row=6, column=0, sticky='ew', pady=10)
        status_frame.columnconfigure(0, weight=1)

        self.current_operation_var = tk.StringVar(value='Ready - Connect instruments and configure settings')
        status_label = ttk.Label(status_frame, textvariable=self.current_operation_var, font=('Segoe UI', 12, 'bold'), foreground='#1e40af')
        status_label.grid(row=0, column=0, sticky='ew', pady=(0, 15))

        log_control_frame = ttk.Frame(status_frame)
        log_control_frame.grid(row=1, column=0, sticky='ew', pady=(0, 10))
        log_control_frame.columnconfigure(1, weight=1)

        ttk.Label(log_control_frame, text='Activity Log:', font=('Segoe UI', 11, 'bold')).grid(row=0, column=0, sticky='w')
        ttk.Button(log_control_frame, text='Clear Log', command=self.clear_log).grid(row=0, column=2, padx=(0, 10))
        ttk.Button(log_control_frame, text='Save Log', command=self.save_log).grid(row=0, column=3)

        # Use a light background and dark monospace text for readability
        self.log_text = scrolledtext.ScrolledText(status_frame, height=15, font=('Consolas', 10), bg='#ffffff', fg='#0f172a', insertbackground='#0f172a', selectbackground='#e2e8f0', selectforeground='#0f172a', wrap=tk.WORD)
        self.log_text.grid(row=2, column=0, sticky='ew')

    def browse_save_location(self, location_type: str):
        try:
            var_mapping = {
                'data': (self.data_path_var, 'data_dir'),
                'graphs': (self.graphs_path_var, 'graphs_dir'),
                'screenshots': (self.screenshots_path_var, 'screenshots_dir')
            }
            if location_type not in var_mapping:
                return
            var_obj, attr_name = var_mapping[location_type]
            current_path = var_obj.get()
            new_path = filedialog.askdirectory(initialdir=current_path if os.path.exists(current_path) else str(Path.cwd()), title=f"Select {location_type.title()} Save Location")
            if new_path:
                var_obj.set(new_path)
                setattr(self.data_manager, attr_name, Path(new_path))
                Path(new_path).mkdir(parents=True, exist_ok=True)
                self.log_message(f"{location_type.title()} save location updated: {new_path}", "SUCCESS")
        except Exception as e:
            self.log_message(f"Error selecting {location_type} folder: {e}", "ERROR")

    def reset_graph_title(self):
        self.graph_title_var.set('SAFE Wave Voltage Ramping')
        self.log_message('Graph title reset to default', 'SUCCESS')

    # Connection Methods
    def connect_psu(self):
        def connect_thread():
            try:
                self.log_message('Connecting to power supply...', 'INFO')
                visa_address = self.psu_visa_var.get().strip()
                self.power_supply = KeithleyPowerSupply(visa_address)
                if self.power_supply.connect():
                    self.power_supply.configure_channel(1, 0.0, 0.1, 0.5, False)
                    self.status_queue.put(('psu_connected', 'PSU connected and initialized to safe state'))
                else:
                    raise Exception('PSU connection failed')
            except Exception as e:
                self.status_queue.put(('error', f'PSU connection error: {str(e)}'))
        threading.Thread(target=connect_thread, daemon=True).start()

    def disconnect_psu(self):
        def disconnect_thread():
            try:
                if self.power_supply:
                    # Try to safely disable all common channels (1-4) before disconnecting.
                    for ch in range(1, 5):
                        try:
                            # Set voltage to 0 and disable output
                            self.power_supply.configure_channel(ch, 0.0, 0.01, 0.5, False)
                        except Exception:
                            # Some drivers may not support multiple channels or may raise - ignore
                            pass
                    time.sleep(0.5)
                    try:
                        self.power_supply.disconnect()
                    except Exception:
                        pass
                    self.power_supply = None
                    self.status_queue.put(('psu_disconnected', 'PSU safely disconnected'))
            except Exception as e:
                self.status_queue.put(('error', f'PSU disconnect error: {str(e)}'))
        threading.Thread(target=disconnect_thread, daemon=True).start()

    def connect_dmm(self):
        def connect_thread():
            try:
                self.log_message('Connecting to DMM...', 'INFO')
                visa_address = self.dmm_visa_var.get().strip()
                self.dmm = KeithleyDMM6500(visa_address)
                if self.dmm.connect():
                    self.status_queue.put(('dmm_connected', 'DMM connected successfully'))
                else:
                    raise Exception('DMM connection failed')
            except Exception as e:
                self.status_queue.put(('error', f'DMM connection error: {str(e)}'))
        threading.Thread(target=connect_thread, daemon=True).start()

    def disconnect_dmm(self):
        def disconnect_thread():
            try:
                if self.dmm:
                    self.dmm.disconnect()
                    self.dmm = None
                    self.status_queue.put(('dmm_disconnected', 'DMM disconnected'))
            except Exception as e:
                self.status_queue.put(('error', f'DMM disconnect error: {str(e)}'))
        threading.Thread(target=disconnect_thread, daemon=True).start()

    def connect_scope(self):
        def connect_thread():
            try:
                self.log_message('Connecting to oscilloscope...', 'INFO')
                visa_address = self.scope_visa_var.get().strip()
                self.oscilloscope = KeysightDSOX6004A(visa_address)
                if self.oscilloscope.connect():
                    # Create the high-level oscilloscope helper if available
                    try:
                        if OscilloscopeDataAcquisition:
                            self.osc_data = OscilloscopeDataAcquisition(self.oscilloscope)
                        else:
                            self.osc_data = None
                    except Exception:
                        self.osc_data = None
                    self.status_queue.put(('scope_connected', 'Oscilloscope connected successfully'))
                else:
                    raise Exception('Oscilloscope connection failed')
            except Exception as e:
                self.status_queue.put(('error', f'Oscilloscope connection error: {str(e)}'))
        threading.Thread(target=connect_thread, daemon=True).start()

    def disconnect_scope(self):
        def disconnect_thread():
            try:
                if self.oscilloscope:
                    self.oscilloscope.disconnect()
                    self.oscilloscope = None
                    self.status_queue.put(('scope_disconnected', 'Oscilloscope disconnected'))
            except Exception as e:
                self.status_queue.put(('error', f'Oscilloscope disconnect error: {str(e)}'))
        threading.Thread(target=disconnect_thread, daemon=True).start()

    def connect_all_instruments(self):
        def connect_all_thread():
            try:
                self.log_message('Connecting all instruments...', 'INFO')
                if not (self.power_supply and getattr(self.power_supply, 'is_connected', False)):
                    self.connect_psu()
                    time.sleep(3)
                if not (self.dmm and getattr(self.dmm, 'is_connected', False)):
                    self.connect_dmm()
                    time.sleep(2)
                if not (self.oscilloscope and getattr(self.oscilloscope, 'is_connected', False)):
                    self.connect_scope()
                    time.sleep(2)
                self.status_queue.put(('all_connected', 'All instruments connection process completed'))
            except Exception as e:
                self.status_queue.put(('error', f'Error connecting all instruments: {str(e)}'))
        threading.Thread(target=connect_all_thread, daemon=True).start()

    def disconnect_all_instruments(self):
        try:
            self.log_message('Disconnecting all instruments...', 'INFO')
            if getattr(self, 'ramping_active', False):
                self.emergency_psu_stop()
            if self.power_supply:
                self.disconnect_psu()
            if self.dmm:
                self.disconnect_dmm()
            if self.oscilloscope:
                self.disconnect_scope()
            self.log_message('All instruments disconnection initiated', 'SUCCESS')
        except Exception as e:
            self.log_message(f'Error disconnecting instruments: {e}', 'ERROR')

    def test_all_connections(self):
        def test_thread():
            try:
                self.log_message('Testing all connections...', 'INFO')
                results = []
                results.append('PSU: CONNECTED' if (self.power_supply and getattr(self.power_supply, 'is_connected', False)) else 'PSU: DISCONNECTED')
                results.append('DMM: CONNECTED' if (self.dmm and getattr(self.dmm, 'is_connected', False)) else 'DMM: DISCONNECTED')
                results.append('Oscilloscope: CONNECTED' if (self.oscilloscope and getattr(self.oscilloscope, 'is_connected', False)) else 'Oscilloscope: DISCONNECTED')
                self.status_queue.put(('connection_test', ' | '.join(results)))
            except Exception as e:
                self.status_queue.put(('error', f'Connection test failed: {str(e)}'))
        threading.Thread(target=test_thread, daemon=True).start()

    # Safety Methods
    def emergency_psu_stop(self):
        try:
            self.emergency_stop = True
            self.ramping_active = False
            self.log_message('EMERGENCY STOP activated', 'ERROR')
            self.safety_status_var.set('EMERGENCY STOP - System halted')
            if self.power_supply and getattr(self.power_supply, 'is_connected', False):
                # Immediately attempt to disable all channels
                for ch in range(1, 5):
                    try:
                        self.power_supply.configure_channel(ch, 0.0, 0.01, 0.5, False)
                    except Exception:
                        pass
                self.set_voltage_display.set('0.000 V')
                self.measured_voltage_display.set('0.000 V')
                self.log_message('PSU immediately disabled - Outputs set to 0 V', 'ERROR')
            else:
                self.log_message('EMERGENCY STOP - No PSU connected to disable', 'WARNING')
            self.current_operation_var.set('EMERGENCY STOPPED - System halted for safety')
        except Exception as e:
            self.log_message(f'Emergency stop error: {e}', 'ERROR')

    # Main Operation Methods
    def start_safe_ramping(self):
        if getattr(self, 'ramping_active', False):
            return

        if not (self.power_supply and getattr(self.power_supply, 'is_connected', False)):
            messagebox.showerror('Safety Error', 'Power Supply not connected\n\nConnect PSU before starting ramping.')
            return

        if not (self.dmm and getattr(self.dmm, 'is_connected', False)):
            messagebox.showerror('Safety Error', 'DMM not connected\n\nConnect DMM before starting ramping.')
            return

        target_v = self.target_voltage_var.get()
        if target_v > 3.5:
            response = messagebox.askyesno('High Voltage Warning', f"Target voltage is {target_v}V (above 3.5V).\n\nThis may cause PSU over-voltage protection to trip.\n\nContinue with ramping?", icon='warning')
            if not response:
                return

        def safe_ramping_thread():
            try:
                self.emergency_stop = False
                self.ramping_active = True
                self.start_ramping_btn.configure(state='disabled')
                self.stop_ramping_btn.configure(state='normal')

                self.data_manager.clear_data()

                # Determine which waveform to use for the active PSU channel.
                active_channel = int(self.psu_active_channel_var.get() or 1)
                # Prefer per-channel assignment if set (non-empty), else fall back to global waveform selection
                per_ch_var = self.psu_channel_waveform_vars.get(active_channel)
                channel_waveform = None
                try:
                    if per_ch_var and per_ch_var.get():
                        channel_waveform = per_ch_var.get()
                except Exception:
                    channel_waveform = None
                if not channel_waveform:
                    channel_waveform = self.waveform_type_var.get()
                self.data_manager.waveform_type = channel_waveform
                generator = WaveformGenerator(
                    waveform_type=channel_waveform,
                    target_voltage=self.target_voltage_var.get(),
                    cycles=self.cycles_var.get(),
                    points_per_cycle=self.points_per_cycle_var.get(),
                    cycle_duration=self.cycle_duration_var.get()
                )
                self.current_profile = generator.generate()
                total_points = len(self.current_profile)
                self.log_message(f"Starting safe ramping: {total_points} points over {self.cycles_var.get()} cycles", 'SUCCESS')

                try:
                    nplc = float(self.nplc_var.get())
                except Exception:
                    nplc = 1.0

                per_point_overhead = float(self.psu_settle_var.get()) + (nplc / 50.0)
                overhead_total = per_point_overhead * total_points

                profile_duration = float(self.current_profile[-1][0]) if self.current_profile else 0.0
                estimated_seconds = max(profile_duration, overhead_total)
                eta_h = int(estimated_seconds // 3600)
                eta_m = int((estimated_seconds % 3600) // 60)
                eta_s = int(estimated_seconds % 60)
                eta_str = f"{eta_h}h {eta_m}m {eta_s}s" if eta_h else f"{eta_m}m {eta_s}s"
                self.log_message(f"Estimated ramping duration: ~{eta_str} (points: {total_points}, NPLC: {nplc})", 'INFO')

                self.eta_total_seconds = float(estimated_seconds)
                self.eta_display_var.set(f"ETA: {eta_str}")
                self.current_operation_var.set('Safe ramping in progress')
                self.safety_status_var.set('Active ramping - Emergency stop available')

                current_limit_cfg = self.current_limit_var.get()
                max_set_voltage = max((v for _, v in self.current_profile), default=0.0)
                ovp_const = min(max_set_voltage + 2, 5.5)

                try:
                    # Pre-configure the selected active channel
                    self.power_supply.configure_channel(int(active_channel), 0.0, current_limit_cfg, ovp_const, True)
                except Exception as e:
                    self.status_queue.put(('error', f'PSU pre-configuration failed: {e}'))
                    return

                start_time = datetime.now()
                successful_points = 0

                for i, (time_point, voltage) in enumerate(self.current_profile):
                    if self.emergency_stop or not self.ramping_active:
                        self.log_message('Ramping stopped by user request', 'WARNING')
                        break

                    try:
                        if voltage > 5.0 or ovp_const > 5.5:
                            raise Exception(f"Voltage exceeds safety limits: {voltage}V (OVP: {ovp_const}V)")

                        success = self.power_supply.configure_channel(channel=int(active_channel), voltage=voltage, current_limit=current_limit_cfg, ovp_level=ovp_const, enable_output=True)
                        if not success:
                            raise Exception('PSU configuration command failed')

                        time.sleep(float(self.psu_settle_var.get()))

                        measured_voltage = self.dmm.measure(MeasurementFunction.DC_VOLTAGE, nplc=nplc)
                        if measured_voltage is None:
                            self.log_message(f'DMM reading failed at point {i}', 'WARNING')
                            measured_voltage = 0.0

                        voltage_error = abs(measured_voltage - voltage)
                        if voltage_error > 0.5 and voltage > 0.5:
                            self.log_message(f'Large voltage error at point {i}: Set={voltage:.3f}V, Measured={measured_voltage:.3f}V', 'WARNING')

                        self.set_voltage_display.set(f"{voltage:.4f} V")
                        self.measured_voltage_display.set(f"{measured_voltage:.4f} V")

                        current_cycle = i // max(1, self.points_per_cycle_var.get())
                        point_in_cycle = i % max(1, self.points_per_cycle_var.get())

                        self.data_manager.add_data_point(datetime.now(), voltage, measured_voltage, current_cycle, point_in_cycle)
                        successful_points += 1

                        progress = (i + 1) / total_points * 100 if total_points else 100
                        self.progress_var.set(progress)

                        elapsed = (datetime.now() - start_time).total_seconds()
                        remaining = max(0.0, self.eta_total_seconds - elapsed)
                        r_h = int(remaining // 3600)
                        r_m = int((remaining % 3600) // 60)
                        r_s = int(remaining % 60)
                        r_str = f"{r_h}h {r_m}m {r_s}s" if r_h else f"{r_m}m {r_s}s"
                        self.eta_display_var.set(f"ETA: {r_str}")

                        # Respect profile timing (sleep until next scheduled profile time)
                        if i + 1 < len(self.current_profile):
                            next_time = self.current_profile[i + 1][0]
                            current_time = (datetime.now() - start_time).total_seconds()
                            sleep_time = max(0.01, next_time - current_time)
                            time.sleep(sleep_time)

                    except Exception as e:
                        self.status_queue.put(('error', f'Error at point {i}: {str(e)}'))
                        self.log_message(f'Point {i} failed: {str(e)}', 'ERROR')
                        continue

                if self.power_supply and getattr(self.power_supply, 'is_connected', False):
                    self.log_message('Performing safe ramp-down to 0 V...', 'INFO')
                    try:
                        # Ramp down the active channel
                        self.power_supply.configure_channel(int(active_channel), 0.0, 0.05, 0.5, False)
                    except Exception:
                        pass
                    time.sleep(0.5)

                self.status_queue.put(('ramping_complete', f"Safe ramping completed successfully\nSuccessful points: {successful_points}/{total_points}\nData points collected: {len(self.data_manager.voltage_data)}"))

            except Exception as e:
                self.status_queue.put(('error', f'Critical ramping error: {str(e)}'))
                try:
                    if self.power_supply and getattr(self.power_supply, 'is_connected', False):
                        self.power_supply.configure_channel(1, 0.0, 0.01, 0.5, False)
                        self.log_message('PSU emergency disabled due to critical error', 'ERROR')
                except Exception:
                    self.log_message('Failed to emergency disable PSU', 'ERROR')

            finally:
                self.ramping_active = False
                self.start_ramping_btn.configure(state='normal')
                self.stop_ramping_btn.configure(state='disabled')
                self.safety_status_var.set('System safe - Ramping completed')
                self.eta_display_var.set('ETA: --')

        self.operation_thread = threading.Thread(target=safe_ramping_thread, daemon=True)
        self.operation_thread.start()

    def stop_ramping(self):
        if getattr(self, 'ramping_active', False):
            self.ramping_active = False
            self.log_message('Ramping stop requested - initiating safe shutdown', 'WARNING')
            if self.power_supply and getattr(self.power_supply, 'is_connected', False):
                try:
                    self.power_supply.configure_channel(1, 0.0, 0.1, 0.5, False)
                    self.set_voltage_display.set('0.000 V')
                    self.log_message('PSU safely set to 0 V and disabled', 'SUCCESS')
                    self.current_operation_var.set('Ramping stopped - PSU safely disabled')
                except Exception as e:
                    self.log_message(f'Error during PSU shutdown: {e}', 'ERROR')

    # Data Management Methods
    def generate_safe_graph(self):
        def graph_thread():
            try:
                if not self.data_manager.voltage_data:
                    self.status_queue.put(('error', 'No data available for graphing'))
                    return
                self.current_operation_var.set('Generating graphs (PSU, DMM, Comparator)...')
                graph_save_location = self.graphs_path_var.get()
                graph_title = self.graph_title_var.get() or 'SAFE Wave Voltage Ramping'

                # Create a single combined figure (PSU + DMM + Comparator + stats)
                combined_path, stats = self.data_manager.generate_combined_graph(custom_title=graph_title, save_location=graph_save_location if os.path.exists(graph_save_location) else None)

                summary = (
                    f'Combined graph saved:\n{combined_path}\n\n'
                    f"Points: {stats.get('points', 0)}\n"
                    f"Meas mean: {stats.get('meas_mean', 0):.4f} V\n"
                    f"Error RMS: {stats.get('error_rms', 0):.6f} V\n"
                )

                self.status_queue.put(('graph_generated', summary))
                try:
                    messagebox.showinfo('Graph Generated', summary)
                except Exception:
                    pass
            except Exception as e:
                self.status_queue.put(('error', f'Graph generation error: {str(e)}'))
        threading.Thread(target=graph_thread, daemon=True).start()

    def export_data(self):
        def export_thread():
            try:
                if not self.data_manager.voltage_data:
                    self.status_queue.put(('error', 'No data available for export'))
                    return
                self.current_operation_var.set('Exporting data to CSV...')
                csv_path = self.data_manager.export_to_csv()
                self.status_queue.put(('csv_exported', f'CSV data exported:\n{csv_path}'))
            except Exception as e:
                self.status_queue.put(('error', f'CSV export error: {str(e)}'))
        threading.Thread(target=export_thread, daemon=True).start()

    def capture_screenshot(self):
        def screenshot_thread():
            try:
                if not (self.oscilloscope and getattr(self.oscilloscope, 'is_connected', False)):
                    self.status_queue.put(('error', 'Oscilloscope not connected for screenshot'))
                    return

                self.current_operation_var.set('Capturing oscilloscope screenshot...')
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f'safe_ramping_screenshot_{timestamp}.png'
                screenshots_dir = Path(self.screenshots_path_var.get())
                screenshots_dir.mkdir(parents=True, exist_ok=True)
                filepath = screenshots_dir / filename

                # Before attempting capture, verify directory is writable
                dir_writable = False
                try:
                    test_file = screenshots_dir / f".__write_test_{timestamp}.tmp"
                    test_file.write_text('test')
                    test_file.unlink()
                    dir_writable = True
                except Exception as dir_test_exc:
                    self.log_message(f'Directory write test failed: {dir_test_exc}', 'ERROR')

                if not dir_writable:
                    self.status_queue.put(('error', f'Screenshot directory not writable: {screenshots_dir}'))
                    return

                attempted = []
                success = False

                # Candidate method names to try on the oscilloscope driver
                candidate_methods = [
                    'save_screenshot', 'capture_screen', 'capture_screenshot', 'get_screenshot',
                    'save_screen', 'save_screengrab', 'save_image', 'get_screen_image'
                ]

                for m in candidate_methods:
                    if hasattr(self.oscilloscope, m):
                        attempted.append(m)
                        method = getattr(self.oscilloscope, m)
                        try:
                            # Try calling with path first
                            result = None
                            try:
                                result = method(str(filepath))
                            except TypeError:
                                # Method may take no args or different signature
                                result = method()

                            # If driver method returned bytes, write them to file
                            if isinstance(result, (bytes, bytearray)):
                                try:
                                    with open(filepath, 'wb') as f:
                                        f.write(result)
                                    success = filepath.exists() and filepath.stat().st_size > 0
                                    if success:
                                        break
                                except Exception as write_exc:
                                    self.log_message(f'Failed to write bytes returned by {m}: {write_exc}', 'ERROR')
                                    continue

                            # If method returned a truthy value, check file existence
                            if result:
                                if filepath.exists() and filepath.stat().st_size > 0:
                                    success = True
                                    break
                                # Some drivers may save to a different default path; check common driver attributes
                                try:
                                    # If result is a path-like string
                                    if isinstance(result, str) and os.path.exists(result):
                                        fp_candidate = Path(result)
                                        if fp_candidate.exists() and fp_candidate.stat().st_size > 0:
                                            filepath = fp_candidate
                                            success = True
                                            break
                                except Exception:
                                    pass

                            # Otherwise, check file presence directly
                            if filepath.exists() and filepath.stat().st_size > 0:
                                success = True
                                break

                        except Exception as e:
                            self.log_message(f'Oscilloscope method {m} raised: {e}', 'ERROR')
                            continue

                # If none of the candidate methods worked, try a generic 'screenshot' attribute that may return raw bytes
                if not success and hasattr(self.oscilloscope, 'screenshot'):
                    attempted.append('screenshot')
                    try:
                        data = self.oscilloscope.screenshot()
                        if isinstance(data, (bytes, bytearray)):
                            with open(filepath, 'wb') as f:
                                f.write(data)
                            success = filepath.exists() and filepath.stat().st_size > 0
                    except Exception as e:
                        self.log_message(f'Oscilloscope.screenshot() failed: {e}', 'ERROR')

                if success and filepath.exists() and filepath.stat().st_size > 0:
                    self.status_queue.put(('screenshot_captured', f'Screenshot saved:\n{filepath}'))
                else:
                    # Detailed error report for debugging
                    error_details = f"Screenshot capture failed. Methods attempted: {attempted}\n"
                    error_details += f"Target file: {filepath}\n"
                    error_details += f"Directory exists: {screenshots_dir.exists()}\n"
                    error_details += f"Directory writable: {os.access(screenshots_dir, os.W_OK)}\n"
                    try:
                        test_file = screenshots_dir / f"test_write_{timestamp}.txt"
                        test_file.write_text(f"Test write at {datetime.now()}")
                        test_file.unlink()
                        error_details += "Directory write test: PASSED\n"
                    except Exception as dir_err:
                        error_details += f"Directory write test: FAILED ({dir_err})\n"

                    error_details += (
                        "Driver capabilities: " + ",".join([attr for attr in dir(self.oscilloscope) if 'save' in attr or 'capture' in attr or 'screenshot' in attr])
                    )

                    self.status_queue.put(('error', f'Screenshot capture failed: {error_details}'))

            except Exception as e:
                self.status_queue.put(('error', f'Screenshot error: {str(e)}'))
        threading.Thread(target=screenshot_thread, daemon=True).start()

    def clear_all_data(self):
        response = messagebox.askyesno('Confirm Data Clear', 'Clear all collected voltage ramping data?\n\nThis action cannot be undone.', icon='warning')
        if response:
            self.data_manager.clear_data()
            self.set_voltage_display.set('0.000 V')
            self.measured_voltage_display.set('0.000 V')
            self.progress_var.set(0)
            self.log_message('All data cleared successfully', 'SUCCESS')

    # Utility Methods
    def log_message(self, message: str, level: str = 'INFO'):
        timestamp = datetime.now().strftime('%H:%M:%S')
        log_entry = f"[{timestamp}] {message}\n"
        try:
            self.log_text.insert(tk.END, log_entry)
            if level == 'ERROR':
                color = '#ff6b6b'
            elif level == 'SUCCESS':
                color = '#51cf66'
            elif level == 'WARNING':
                color = '#ffd43b'
            else:
                color = '#74c0fc'
            tag = level + str(time.time())
            self.log_text.tag_add(tag, f"end-{len(log_entry)}c", "end-1c")
            self.log_text.tag_config(tag, foreground=color)
            self.log_text.see(tk.END)
        except Exception:
            pass

    def clear_log(self):
        try:
            self.log_text.delete(1.0, tk.END)
            self.log_message('Activity log cleared', 'SUCCESS')
        except Exception as e:
            print(f'Clear log error: {e}')

    def save_log(self):
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = filedialog.asksaveasfilename(defaultextension='.txt', filetypes=[('Text files', '*.txt'), ('All files', '*.*')], initialname=f'safe_ramping_log_{timestamp}.txt', title='Save Activity Log')
            if filename:
                log_content = self.log_text.get(1.0, tk.END)
                with open(filename, 'w') as f:
                    f.write('SAFE Voltage Ramping Activity Log\n')
                    f.write(f'Generated: {datetime.now()}\n')
                    f.write('=' * 60 + '\n\n')
                    f.write(log_content)
                self.log_message(f'Activity log saved: {filename}', 'SUCCESS')
        except Exception as e:
            self.log_message(f'Failed to save log: {e}', 'ERROR')

    def check_status_updates(self):
        try:
            while True:
                status_type, data = self.status_queue.get_nowait()
                if status_type == 'psu_connected':
                    self.psu_status_var.set('Connected')
                    self.psu_disconnect_btn.configure(state='normal')
                    self.log_message(data, 'SUCCESS')
                elif status_type == 'psu_disconnected':
                    self.psu_status_var.set('Disconnected')
                    self.psu_disconnect_btn.configure(state='disabled')
                    self.log_message(data, 'SUCCESS')
                elif status_type == 'dmm_connected':
                    self.dmm_status_var.set('Connected')
                    self.dmm_disconnect_btn.configure(state='normal')
                    self.log_message(data, 'SUCCESS')
                elif status_type == 'dmm_disconnected':
                    self.dmm_status_var.set('Disconnected')
                    self.dmm_disconnect_btn.configure(state='disabled')
                    self.log_message(data, 'SUCCESS')
                elif status_type == 'scope_connected':
                    self.scope_status_var.set('Connected')
                    self.scope_disconnect_btn.configure(state='normal')
                    self.log_message(data, 'SUCCESS')
                elif status_type == 'scope_disconnected':
                    self.scope_status_var.set('Disconnected')
                    self.scope_disconnect_btn.configure(state='disabled')
                    self.log_message(data, 'SUCCESS')
                elif status_type == 'all_connected':
                    self.current_operation_var.set('All instruments connection process completed')
                    self.log_message(data, 'SUCCESS')
                elif status_type == 'connection_test':
                    self.current_operation_var.set(data)
                    self.log_message(data, 'INFO')
                elif status_type == 'ramping_complete':
                    self.current_operation_var.set('Ramping complete - Ready for next operation')
                    self.log_message(data, 'SUCCESS')
                elif status_type == 'graph_generated':
                    self.current_operation_var.set('Graph generated and saved')
                    self.log_message(data, 'SUCCESS')
                elif status_type == 'csv_exported':
                    self.current_operation_var.set('CSV exported successfully')
                    self.log_message(data, 'SUCCESS')
                elif status_type == 'screenshot_captured':
                    self.current_operation_var.set('Screenshot captured successfully')
                    self.log_message(data, 'SUCCESS')
                elif status_type == 'error':
                    self.current_operation_var.set(data)
                    self.log_message(data, 'ERROR')
                    self.current_operation_var.set('ERROR OCCURRED - Check activity log')
        except queue.Empty:
            pass
        except Exception as e:
            print(f'Status update error: {e}')
        self.root.after(100, self.check_status_updates)

    def safe_shutdown(self):
        try:
            self.log_message('🔄 Initiating safe application shutdown...', 'INFO')
            if getattr(self, 'ramping_active', False):
                self.emergency_psu_stop()
                time.sleep(1)
            self.disconnect_all_instruments()
            time.sleep(1)
            self.log_message('✅ Safe shutdown completed', 'SUCCESS')
            time.sleep(0.5)
        except Exception as e:
            print(f'Shutdown error: {e}')
        finally:
            # Safely destroy the Tk root if it still exists
            try:
                # Ensure PSU outputs are disabled before exit if still connected
                try:
                    if getattr(self, 'power_supply', None) and getattr(self.power_supply, 'is_connected', False):
                        for ch in range(1, 5):
                            try:
                                self.power_supply.configure_channel(ch, 0.0, 0.01, 0.5, False)
                            except Exception:
                                pass
                except Exception:
                    pass

                if getattr(self, 'root', None):
                    try:
                        exists = bool(self.root.winfo_exists())
                    except Exception:
                        exists = False

                    if exists:
                        try:
                            self.root.destroy()
                        except Exception as destroy_exc:
                            print(f"Warning: root.destroy() failed during safe_shutdown: {destroy_exc}")
            except Exception:
                pass

    def run(self):
        try:
            self.log_message('🟢 SAFE Voltage Ramping Application Started - Improved UI v2.3', 'SUCCESS')
            self.log_message('🛡️ PSU SAFETY PRIORITY - Emergency stop available at all times', 'WARNING')
            self.log_message('📜 Scrollable interface - use mouse wheel to scroll', 'INFO')
            self.root.mainloop()
        except Exception as e:
            self.log_message(f'❌ Application error: {e}', 'ERROR')
        finally:
            self.safe_shutdown()


def main():
    try:
        app = ImprovedSafeVoltageRampingGUI()
        app.run()
    except Exception as e:
        print(f'❌ FATAL APPLICATION ERROR: {e}')
        sys.exit(1)


if __name__ == '__main__':
    main()
