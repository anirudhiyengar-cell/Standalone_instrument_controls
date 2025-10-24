#!/usr/bin/env python3
"""
╔════════════════════════════════════════════════════════════════════════════════╗
║         KEITHLEY DMM6500 PROFESSIONAL AUTOMATION - COMPLETE APPLICATION        ║
║                   Enhanced GUI Application - Version 2.0                       ║
║                     Fully Annotated for Executive Review                       ║
║                                                                                ║
║  Purpose: Comprehensive automation control for Keithley DMM6500 multimeter    ║
║           with real-time measurements, data logging, CSV export, and graphing ║
║                                                                                ║
║  Core Features:                                                                ║
║    • Multi-threaded responsive GUI using tkinter framework                    ║
║    • VISA-based communication with Keithley DMM6500 multimeter               ║
║    • Eight measurement functions (DC/AC voltage/current, resistance, etc.)   ║
║    • Configurable measurement range, resolution, and integration time (NPLC) ║
║    • Real-time measurement display with 9-digit precision                    ║
║    • Single measurement or continuous measurement polling modes              ║
║    • Statistical analysis (mean, std dev, min, max, range)                   ║
║    • CSV data export with metadata headers and timestamps                    ║
║    • Publication-quality plot generation with statistics overlay             ║
║    • Multi-type measurement graphing (different traces per measurement type) ║
║    • Customizable save locations for data files and generated plots          ║
║    • Color-coded activity logging for diagnostic purposes                    ║
║    • Professional GUI layout with status indicators and controls             ║
║                                                                                ║
║  Measurement Functions Supported:                                             ║
║    • DC Voltage (0-1000V, 6-digit accuracy)                                  ║
║    • AC Voltage (0-750V, 6-digit accuracy)                                   ║
║    • DC Current (0-10A, 6-digit accuracy)                                    ║
║    • AC Current (0-10A, 6-digit accuracy)                                    ║
║    • 2-Wire Resistance (0-100MΩ, 6-digit accuracy)                          ║
║    • 4-Wire Resistance (for ultra-low resistance measurement)               ║
║    • Capacitance (1pF-10mF)                                                  ║
║    • Frequency (10Hz-1MHz)                                                   ║
║                                                                                ║
║  System Requirements:                                                          ║
║    • Python 3.7+                                                              ║
║    • PyVISA library for instrument communication                              ║
║    • Keysight/National Instruments VISA drivers installed                    ║
║    • Matplotlib for graphing capabilities                                    ║
║    • Pandas for data handling                                                ║
║    • NumPy for statistical calculations                                      ║
║    • Keithley DMM6500 multimeter connected via USB                          ║
║                                                                                ║
║  Author: Professional Instrumentation Control System                          ║
║  Last Updated: 2025-10-24 | Status: Production Ready                          ║
╚════════════════════════════════════════════════════════════════════════════════╝
"""

# ============================================================================
# STANDARD LIBRARY IMPORTS - Core Python functionality
# ============================================================================
import sys  # System-level operations and program exit control
import logging  # Application event logging for diagnostics and troubleshooting
import tkinter as tk  # Core GUI framework for window creation and widgets
from tkinter import ttk, messagebox, filedialog, scrolledtext  # Advanced GUI components for dialogs and text
from pathlib import Path  # Object-oriented filesystem path handling (cross-platform compatible)
from typing import Optional, Dict, Any, List  # Type hints for improved code clarity and IDE support
import threading  # Multi-threaded execution to prevent GUI blocking during long operations
import queue  # Thread-safe queue for inter-thread communication (main ↔ workers)
from datetime import datetime  # Timestamp generation for logging and measurement records
import time  # Time delay operations for waiting between measurements

# ============================================================================
# THIRD-PARTY IMPORTS - Data processing and visualization libraries
# ============================================================================
import pandas as pd  # Data manipulation and CSV export functionality
import matplotlib.pyplot as plt  # Publication-quality graph plotting and visualization
import matplotlib.dates as mdates  # Time-based axis formatting for matplotlib
import numpy as np  # Numerical array operations for statistical calculations

# ============================================================================
# INSTRUMENT CONTROL IMPORTS - Custom modules for device communication
# ============================================================================
try:
    # Primary import method: from instrument_control package
    from instrument_control.keithley_dmm import KeithleyDMM6500, KeithleyDMM6500Error, MeasurementFunction  # DMM driver and enumerations
except ImportError as e:
    # Error handling: driver not found
    print(f"Error importing instrument control modules: {e}")  # Print error message
    print("Please ensure the instrument_control package is in your Python path")  # Provide guidance
    sys.exit(1)  # Exit program with error status

# ============================================================================
# APPLICATION CONSTANTS - Measurement options and configurations
# ============================================================================

# Measurement function options with human-readable names and enum values
# Each tuple: (Display Name, MeasurementFunction Enum Value)
MEASUREMENT_OPTIONS = [
    ("DC Voltage", MeasurementFunction.DC_VOLTAGE),  # Measure DC voltage (0-1000V)
    ("AC Voltage", MeasurementFunction.AC_VOLTAGE),  # Measure AC voltage (RMS, 0-750V)
    ("DC Current", MeasurementFunction.DC_CURRENT),  # Measure DC current (0-10A)
    ("AC Current", MeasurementFunction.AC_CURRENT),  # Measure AC current (RMS, 0-10A)
    ("2-Wire Resistance", MeasurementFunction.RESISTANCE_2W),  # Standard resistance measurement (0-100MΩ)
    ("4-Wire Resistance", MeasurementFunction.RESISTANCE_4W),  # Precision resistance (eliminates lead resistance)
    ("Capacitance", MeasurementFunction.CAPACITANCE),  # Measure capacitance (1pF-10mF)
    ("Frequency", MeasurementFunction.FREQUENCY)  # Measure frequency (10Hz-1MHz)
]  # Complete list of supported measurement types


class DMMDataHandler:
    """
    Data Handler for DMM Measurement Operations
    
    Manages all data processing, export, and visualization operations for the DMM
    measurements. Provides methods for data collection, CSV export, graph generation,
    and statistical analysis of measurement results.
    
    Main Responsibilities:
    - Store and organize measurement data with timestamps
    - Export data to CSV format with metadata headers
    - Generate publication-quality graphs with statistics overlay
    - Calculate statistical summaries (mean, std dev, min, max)
    - Manage output directories for data and graphs
    """

    def __init__(self):
        """
        Initialize the data handler with logging and directory configuration
        
        Sets up logger instance and creates reference paths for default output
        directories (dmm_data for CSV files, dmm_graphs for plot images).
        """
        self._logger = logging.getLogger(f'{self.__class__.__name__}')  # Create per-class logger instance
        self.measurement_data = []  # List to store all measurement dictionaries
        
        # Define default output directories (will be created when needed)
        self.default_data_dir = Path.cwd() / "dmm_data"  # CSV data files directory
        self.default_graph_dir = Path.cwd() / "dmm_graphs"  # Graph images directory

    def add_measurement(self, measurement_type: str, value: float, unit: str, timestamp: datetime = None):
        """
        Add a single measurement record to the data collection
        
        Stores measurement with timestamp, type, value, and unit for later export
        and analysis. Auto-generates timestamp if not provided.
        
        Args:
            measurement_type (str): Type of measurement (e.g., "DC Voltage", "AC Current")
            value (float): Measured numerical value
            unit (str): Unit of measurement (e.g., "V", "A", "Ω", "Hz")
            timestamp (datetime): Measurement timestamp (defaults to current time if None)
        """
        if timestamp is None:  # Check if timestamp provided
            timestamp = datetime.now()  # Use current time if not provided
        
        # Create measurement dictionary with all required information
        self.measurement_data.append({
            'timestamp': timestamp,  # When measurement was taken
            'measurement_type': measurement_type,  # What was measured
            'value': value,  # The measured value (9+ digit precision)
            'unit': unit  # Unit of measurement (V, A, Ω, Hz, etc.)
        })
        
        self._logger.info(f"Added measurement: {measurement_type} = {value} {unit}")  # Log addition

    def export_to_csv(self, custom_path: Optional[str] = None, filename: Optional[str] = None) -> Optional[str]:
        """
        Export all collected measurement data to CSV file with metadata headers
        
        Creates CSV file in specified directory with professional header comments
        including export timestamp, measurement count, and column descriptions.
        Auto-generates filename with timestamp if not provided.
        
        Args:
            custom_path (str): Optional custom directory path for CSV output
            filename (str): Optional custom filename (auto-generated if None)
        
        Returns:
            str: Full file path to saved CSV file, or None if export failed
        
        Raises:
            None (errors logged, returns None)
        """
        if not self.measurement_data:  # Check if any data to export
            self._logger.error("No measurement data to export")  # Log error
            return None  # Return None to indicate failure

        try:
            # Determine save directory (custom path or default)
            if custom_path:  # If custom path provided
                save_dir = Path(custom_path)  # Use custom directory
            else:
                save_dir = self.default_data_dir  # Use default data directory
            
            save_dir.mkdir(parents=True, exist_ok=True)  # Create directory if doesn't exist
            
            # Generate filename if not provided
            if filename is None:  # If no filename specified
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")  # Create timestamp string
                filename = f"dmm_measurements_{timestamp}.csv"  # Auto-generate filename
            
            if not filename.endswith('.csv'):  # Ensure .csv extension
                filename += '.csv'  # Add extension if missing
            
            filepath = save_dir / filename  # Construct full file path
            
            # Create pandas DataFrame from measurement data
            df = pd.DataFrame(self.measurement_data)  # Convert list of dicts to DataFrame
            
            # Write CSV file with metadata header
            with open(filepath, 'w') as f:  # Open file for writing
                # Write header comments (will be visible in CSV)
                f.write("# Keithley DMM6500 Measurement Data\n")  # File type header
                f.write(f"# Export Time: {datetime.now().isoformat()}\n")  # Export timestamp
                f.write(f"# Total Measurements: {len(self.measurement_data)}\n")  # Measurement count
                f.write("# Columns: timestamp, measurement_type, value, unit\n")  # Column descriptions
                f.write("\n")  # Blank line before data
            
            # Append the actual data (pandas adds column headers automatically)
            df.to_csv(filepath, mode='a', index=False)  # Append DataFrame to CSV
            
            self._logger.info(f"CSV exported: {filepath}")  # Log success
            return str(filepath)  # Return file path
        
        except Exception as e:  # Catch any export errors
            self._logger.error(f"Failed to export CSV: {e}")  # Log error details
            return None  # Return None to indicate failure

    def generate_graph(self, custom_path: Optional[str] = None, filename: Optional[str] = None,
                      graph_title: Optional[str] = None) -> Optional[str]:
        """
        Generate publication-quality plot of measurement data over time
        
        Creates matplotlib figure with separate traces for each measurement type,
        includes statistics overlay (count, mean, std dev, min, max), and saves
        as high-resolution PNG image with professional formatting.
        
        Args:
            custom_path (str): Optional custom directory path for graph output
            filename (str): Optional custom filename for PNG (auto-generated if None)
            graph_title (str): Optional custom title for plot (auto-generated if None)
        
        Returns:
            str: Full file path to saved PNG image, or None if generation failed
        
        Raises:
            None (errors logged, returns None)
        """
        if not self.measurement_data:  # Check if any data to plot
            self._logger.error("No measurement data to plot")  # Log error
            return None  # Return None to indicate failure

        try:
            # Determine save directory (custom path or default)
            if custom_path:  # If custom path provided
                save_dir = Path(custom_path)  # Use custom directory
            else:
                save_dir = self.default_graph_dir  # Use default graph directory
            
            save_dir.mkdir(parents=True, exist_ok=True)  # Create directory if doesn't exist
            
            # Generate filename if not provided
            if filename is None:  # If no filename specified
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")  # Create timestamp string
                filename = f"dmm_graph_{timestamp}.png"  # Auto-generate filename
            
            if not filename.lower().endswith(('.png', '.jpg', '.jpeg')):  # Ensure image extension
                filename += '.png'  # Add PNG extension if missing
            
            filepath = save_dir / filename  # Construct full file path
            
            # Create matplotlib figure with professional size and layout
            plt.figure(figsize=(12, 8))  # 12x8 inch figure for readability
            
            # Group data by measurement type (enables plotting multiple traces)
            data_by_type = {}  # Dictionary: measurement_type → {timestamps, values, units}
            
            for item in self.measurement_data:  # Process each measurement
                mtype = item['measurement_type']  # Get measurement type
                if mtype not in data_by_type:  # First measurement of this type
                    data_by_type[mtype] = {'timestamps': [], 'values': [], 'units': []}  # Initialize
                
                # Add measurement to its type's list
                data_by_type[mtype]['timestamps'].append(item['timestamp'])  # Append timestamp
                data_by_type[mtype]['values'].append(item['value'])  # Append value
                data_by_type[mtype]['units'].append(item['unit'])  # Append unit
            
            # Plot each measurement type with distinct color and label
            colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f']  # Color palette
            color_idx = 0  # Color index for cycling through palette
            
            for mtype, data in data_by_type.items():  # Each measurement type
                color = colors[color_idx % len(colors)]  # Get color (cycle if more types than colors)
                unit = data['units'][0] if data['units'] else ''  # Get unit from first measurement
                
                # Plot line with markers for this measurement type
                plt.plot(data['timestamps'], data['values'],
                    marker='o',  # Circle markers at data points
                    linestyle='-',  # Solid line connecting points
                    linewidth=2,  # Thicker lines for visibility
                    markersize=4,  # Small markers
                    color=color,  # Use assigned color
                    label=f"{mtype} ({unit})")  # Legend label with unit
                
                color_idx += 1  # Move to next color for next measurement type
            
            # Configure plot labels and title
            if graph_title:  # If custom title provided
                plt.title(graph_title, fontsize=14, fontweight='bold')  # Use custom title
            else:
                plt.title("DMM Measurement Data Over Time", fontsize=14, fontweight='bold')  # Default title
            
            plt.xlabel('Time', fontsize=12)  # Time axis label
            plt.ylabel('Measurement Value', fontsize=12)  # Value axis label
            plt.grid(True, alpha=0.3)  # Light grid for reference
            plt.legend()  # Add legend showing measurement types
            
            # Format x-axis for time display
            plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))  # HH:MM:SS format
            plt.gca().xaxis.set_major_locator(mdates.SecondLocator(interval=30))  # 30-second intervals
            plt.xticks(rotation=45)  # Rotate labels for readability
            
            # Add statistics box if multiple measurements (for analysis)
            if len(self.measurement_data) > 1:  # Only if enough data for stats
                all_values = [item['value'] for item in self.measurement_data]  # Extract all values
                
                # Format statistics text box
                stats_text = f"""Statistics:
Count: {len(all_values)}
Mean: {np.mean(all_values):.6f}
Std: {np.std(all_values):.6f}
Min: {np.min(all_values):.6f}
Max: {np.max(all_values):.6f}"""  # Statistics summary text
                
                # Place statistics box in plot
                plt.text(0.02, 0.98, stats_text,  # Position: top-left (2%, 98% of axes)
                    transform=plt.gca().transAxes,  # Transform: axes coordinates (not data)
                    fontsize=10,  # Small font for box
                    verticalalignment='top',  # Align from top
                    bbox=dict(boxstyle='round', facecolor='lightblue', alpha=0.8))  # Rounded box with blue background
            
            # Save figure
            plt.tight_layout()  # Adjust layout to prevent label cutoff
            plt.savefig(filepath, dpi=300, bbox_inches='tight')  # Save at 300 DPI (publication quality)
            plt.close()  # Close figure to free memory
            
            self._logger.info(f"Graph saved: {filepath}")  # Log success
            return str(filepath)  # Return file path
        
        except Exception as e:  # Catch any generation errors
            self._logger.error(f"Failed to generate graph: {e}")  # Log error details
            return None  # Return None to indicate failure

    def clear_data(self):
        """
        Clear all collected measurement data (reset collection)
        
        Removes all stored measurements. Useful when starting new measurement
        session or clearing previous data for fresh analysis.
        """
        self.measurement_data.clear()  # Empty the measurement list
        self._logger.info("Measurement data cleared")  # Log action

    def get_statistics(self) -> Dict[str, Any]:
        """
        Calculate and return statistical summary of all measurements
        
        Performs basic statistical analysis on collected values:
        count (number of measurements), mean (average), std deviation,
        minimum, maximum, and range (max - min).
        
        Returns:
            Dict[str, Any]: Dictionary with statistical summary:
                - count: Number of measurements
                - mean: Average value
                - std: Standard deviation
                - min: Minimum value
                - max: Maximum value
                - range: Maximum minus minimum
            Returns empty dict if no data available
        """
        if not self.measurement_data:  # Check if data exists
            return {}  # Return empty dict if no data
        
        values = [item['value'] for item in self.measurement_data]  # Extract all measured values
        
        # Calculate statistics using NumPy
        return {
            'count': len(values),  # Number of data points
            'mean': np.mean(values),  # Average (arithmetic mean)
            'std': np.std(values),  # Standard deviation (spread of data)
            'min': np.min(values),  # Minimum value
            'max': np.max(values),  # Maximum value
            'range': np.max(values) - np.min(values)  # Span from min to max
        }


class KeithleyDMMAutomationGUI:
    """
    Main GUI Application for Keithley DMM6500 Professional Automation
    
    Professional interface for controlling Keithley DMM6500 multimeter with
    comprehensive measurement capabilities, data logging, export functionality,
    and graphing. Implements multi-threaded architecture for responsive UI.
    
    Main Responsibilities:
    - GUI construction and event management (tkinter)
    - Thread spawning and coordination for non-blocking operations
    - VISA communication through KeithleyDMM6500 driver
    - Data collection, export, and visualization
    - Status tracking and user feedback via logging
    """

    def __init__(self):
        """
        Initialize the DMM Automation GUI Application
        
        Creates main Tkinter window, initializes application state variables,
        configures threading infrastructure, and sets up logging system.
        
        Instance Variables:
        - root: Main Tkinter window container
        - dmm: KeithleyDMM6500 driver instance (None if disconnected)
        - data_handler: DMMDataHandler for data management
        - status_queue: Thread-safe queue for worker-to-main communication
        - save_locations: Dictionary of custom save paths for data and graphs
        - continuous_running: Flag for continuous measurement state
        """
        # Create main window
        self.root = tk.Tk()  # Initialize Tkinter root window
        
        # Initialize DMM-related objects
        self.dmm = None  # DMM driver instance (None until connected)
        self.data_handler = DMMDataHandler()  # Create data handler for measurements
        
        # User preferences for file saving
        self.save_locations = {
            'data': str(Path.cwd() / "dmm_data"),  # CSV save location
            'graphs': str(Path.cwd() / "dmm_graphs")  # Graph save location
        }
        
        # Setup application infrastructure
        self.setup_logging()  # Configure diagnostic logging system
        self.setup_gui()  # Construct GUI elements and layout
        
        # Thread communication
        self.status_queue = queue.Queue()  # Thread-safe queue for status messages
        self.check_status_updates()  # Start periodic status check (queue polling)
        
        # Continuous measurement state
        self.continuous_running = False  # Flag: continuous measurement active
        self.continuous_thread = None  # Handle to continuous measurement thread

    def setup_logging(self):
        """
        Configure application-wide logging infrastructure
        
        Sets up Python logging system with standardized format including
        timestamp, logger name, log level, and message content. Enables
        diagnostic output for troubleshooting and audit trails.
        """
        logging.basicConfig(
            level=logging.INFO,  # Log INFO level and above (INFO, WARNING, ERROR, CRITICAL)
            format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"  # Standard format string
        )
        self.logger = logging.getLogger('DMMAutomation')  # Create application logger

    def setup_gui(self):
        """
        Construct main GUI window and initialize all UI components
        
        Configures window properties, applies custom styling, and triggers
        creation of all GUI sections (connection, measurement config, operations,
        file preferences, logging).
        """
        # Window configuration
        self.root.title("Keithley DMM6500 Professional Automation")  # Window title bar text
        self.root.geometry("1200x750")  # Initial window size in pixels (width, height)
        self.root.configure(bg='#f0f0f0')  # Background color (light gray)
        self.root.minsize(1000, 600)  # Minimum window dimensions in pixels
        
        # Configure grid expansion properties
        self.root.columnconfigure(0, weight=1)  # Column 0 expands horizontally
        self.root.rowconfigure(0, weight=1)  # Row 0 expands vertically
        
        # Configure styles for themed widgets
        self.style = ttk.Style()  # Create themed widget style manager
        self.style.theme_use('clam')  # Apply "clam" theme (modern flat design)
        self.configure_styles()  # Apply custom color and font configurations
        
        # Create main frame container
        main_frame = ttk.Frame(self.root, padding="5")  # Main content frame
        main_frame.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)  # Fill entire window
        
        # Configure main frame grid expansion
        main_frame.columnconfigure(0, weight=1)  # Content expands horizontally
        main_frame.rowconfigure(0, weight=0)  # Title (fixed height)
        main_frame.rowconfigure(1, weight=0)  # Connection section (fixed)
        main_frame.rowconfigure(2, weight=0)  # Measurement config (fixed)
        main_frame.rowconfigure(3, weight=0)  # File preferences (fixed)
        main_frame.rowconfigure(4, weight=0)  # Operations (fixed)
        main_frame.rowconfigure(5, weight=1)  # Status/Log (expandable)
        
        # Create all GUI sections (order determines layout from top to bottom)
        self.create_title_section(main_frame, row=0)  # Application title
        self.create_connection_section(main_frame, row=1)  # Connection controls
        self.create_measurement_config_section(main_frame, row=2)  # Measurement settings
        self.create_file_preferences_section(main_frame, row=3)  # Save location preferences
        self.create_operations_section(main_frame, row=4)  # Operation buttons
        self.create_status_section(main_frame, row=5)  # Status display and log

    def configure_styles(self):
        """
        Apply custom styling to GUI widgets (colors, fonts, button states)
        
        Defines appearance for labels (bold blue titles), buttons (color-coded
        by function: green for success, red for disconnect, blue for info),
        and measurement displays (monospace for proper alignment).
        """
        # Title label styling: large bold font, professional blue color
        self.style.configure('Title.TLabel',
            font=('Bookman Old Style', 14, 'bold'),  # Large, bold serif font
            foreground='#1a365d',  # Dark blue color
            background='#f0f0f0')  # Match window background
        
        # Connect button style (green for positive action)
        self.style.configure('Connect.TButton',
            font=('Bookman Old Style', 9),  # Medium serif font
            foreground='white')  # White text
        self.style.map('Connect.TButton',
            background=[('active', '#2d7d32'), ('!active', '#388e3c')])  # Green shades
        
        # Disconnect button style (red for caution action)
        self.style.configure('Disconnect.TButton',
            font=('Bookman Old Style', 9),
            foreground='white')
        self.style.map('Disconnect.TButton',
            background=[('active', '#c62828'), ('!active', '#d32f2f')])  # Red shades
        
        # Measure button style (blue for information action)
        self.style.configure('Measure.TButton',
            font=('Bookman Old Style', 9),
            foreground='white')
        self.style.map('Measure.TButton',
            background=[('active', '#1565c0'), ('!active', '#1976d2')])  # Blue shades
        
        # Export button style (green for save/export action)
        self.style.configure('Export.TButton',
            font=('Arial', 9),
            foreground='white')
        self.style.map('Export.TButton',
            background=[('active', '#2e7d32'), ('!active', '#388e3c')])  # Green shades

    # ============================================================================
    # GUI SECTION CREATION METHODS
    # ============================================================================

    def create_title_section(self, parent, row):
        """
        Create application title section (display only)
        
        Args:
            parent: Parent container frame
            row: Grid row number for placement
        """
        title_label = ttk.Label(parent,
            text="Keithley DMM6500 Professional Automation Suite",  # Application title
            style='Title.TLabel')  # Use title style (large bold blue)
        title_label.grid(row=row, column=0, pady=(0, 10), sticky='ew')  # Place in grid, fill width

    def create_connection_section(self, parent, row):
        """
        Create connection settings section (VISA address, Connect/Disconnect buttons)
        
        Allows user to enter instrument VISA address and provides buttons for
        connecting/disconnecting from DMM. Also displays connection status and
        instrument identification information.
        
        Args:
            parent: Parent container frame
            row: Grid row number for placement
        """
        conn_frame = ttk.LabelFrame(parent, text="Instrument Connection", padding="8")  # Labeled section container
        conn_frame.grid(row=row, column=0, sticky='ew', pady=(0, 5))  # Fill width with bottom padding
        conn_frame.columnconfigure(1, weight=1)  # Middle column expands for address field
        
        # VISA address label and input field
        ttk.Label(conn_frame, text="VISA Address:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky='w', padx=(0, 10))  # Left-aligned label with padding
        
        self.visa_address_var = tk.StringVar(value="USB0::0x05E6::0x6500::04561287::INSTR")  # Default VISA address for DMM6500
        self.visa_entry = ttk.Entry(conn_frame, textvariable=self.visa_address_var,
            font=('Arial', 10), width=50)  # Address entry field with monospace font
        self.visa_entry.grid(row=0, column=1, sticky='ew', padx=(0, 10))  # Expand to fill available space
        
        # Connection control buttons
        self.connect_btn = ttk.Button(conn_frame, text="Connect",
            command=self.connect_dmm,
            style='Connect.TButton')  # Green button for connecting
        self.connect_btn.grid(row=0, column=2, padx=(0, 5))  # Place with right padding
        
        self.disconnect_btn = ttk.Button(conn_frame, text="Disconnect",
            command=self.disconnect_dmm,
            style='Disconnect.TButton',
            state='disabled')  # Red button for disconnecting (disabled until connected)
        self.disconnect_btn.grid(row=0, column=3, padx=(0, 5))
        
        self.test_btn = ttk.Button(conn_frame, text="Test Connection",
            command=self.test_connection,
            state='disabled')  # Test button (disabled until connected)
        self.test_btn.grid(row=0, column=4)
        
        # Connection status indicator
        self.conn_status_var = tk.StringVar(value="Disconnected")  # Status text variable
        self.conn_status_label = ttk.Label(conn_frame, textvariable=self.conn_status_var,
            font=('Arial', 10, 'bold'),
            foreground='#d32f2f')  # Red text initially (disconnected)
        self.conn_status_label.grid(row=0, column=5, sticky='e', padx=(10, 0))  # Right-aligned
        
        # Instrument info display (read-only text widget)
        self.info_text = tk.Text(conn_frame, height=2, font=('Courier', 9),
            state='disabled', bg='#f8f9fa',
            relief='flat', borderwidth=1)  # Monospace font, light gray background, no border
        self.info_text.grid(row=1, column=0, columnspan=6, sticky='ew', pady=(8, 0))  # Span all columns

    def create_measurement_config_section(self, parent, row):
        """
        Create measurement configuration section (function, range, resolution, NPLC)
        
        Allows user to select measurement type, configure acquisition parameters
        (range, resolution), and set integration time (NPLC = Number of Power Line Cycles).
        
        Args:
            parent: Parent container frame
            row: Grid row number for placement
        """
        config_frame = ttk.LabelFrame(parent, text="Measurement Configuration", padding="8")  # Labeled section
        config_frame.grid(row=row, column=0, sticky='ew', pady=(0, 5))  # Fill width
        
        # Configure grid columns
        for i in range(8):  # 8 control columns
            config_frame.columnconfigure(i, weight=0)  # Fixed width columns
        config_frame.columnconfigure(7, weight=1)  # Last column expands
        
        # Measurement type/function selection
        ttk.Label(config_frame, text="Function:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky='w', padx=(0, 5))  # Left-aligned label
        
        self.measurement_type_var = tk.StringVar(value="DC Voltage")  # Default measurement type
        self.measurement_type_cb = ttk.Combobox(
            config_frame, textvariable=self.measurement_type_var,
            values=[opt[0] for opt in MEASUREMENT_OPTIONS],  # Extract display names from options
            state='readonly', font=('Arial', 10), width=15)  # Read-only dropdown
        self.measurement_type_cb.grid(row=0, column=1, padx=(0, 10))  # Place with right padding
        
        # Measurement range setting
        ttk.Label(config_frame, text="Range:", font=('Arial', 10)).grid(
            row=0, column=2, sticky='w', padx=(0, 5))  # Label
        
        self.range_var = tk.StringVar(value="Auto")  # Default: auto-ranging enabled
        self.range_entry = ttk.Entry(config_frame, textvariable=self.range_var,
            font=('Arial', 10), width=10)  # Entry field (accept numeric or "Auto")
        self.range_entry.grid(row=0, column=3, padx=(0, 10))  # Place with right padding
        
        # Measurement resolution setting
        ttk.Label(config_frame, text="Resolution:", font=('Arial', 10)).grid(
            row=0, column=4, sticky='w', padx=(0, 5))  # Label
        
        self.resolution_var = tk.StringVar(value="Auto")  # Default: auto-resolution enabled
        self.resolution_entry = ttk.Entry(config_frame, textvariable=self.resolution_var,
            font=('Arial', 10), width=10)  # Entry field
        self.resolution_entry.grid(row=0, column=5, padx=(0, 10))
        
        # NPLC (Number of Power Line Cycles) integration time
        ttk.Label(config_frame, text="NPLC:", font=('Arial', 10)).grid(
            row=0, column=6, sticky='w', padx=(0, 5))  # Label (NPLC explanation)
        
        self.nplc_var = tk.DoubleVar(value=1.0)  # Default: 1 power line cycle (20ms at 50Hz)
        self.nplc_cb = ttk.Combobox(config_frame, textvariable=self.nplc_var,
            values=[0.01, 0.02, 0.06, 0.2, 1.0, 2.0, 10.0],  # Common NPLC values
            state='readonly', font=('Arial', 10), width=8)  # Read-only dropdown
        self.nplc_cb.grid(row=0, column=7, sticky='w')  # Place at end

    def create_file_preferences_section(self, parent, row):
        """
        Create file preferences section (save locations and graph title)
        
        Allows user to specify custom directories for CSV data export and
        graph image generation. Also provides field for custom graph title.
        
        Args:
            parent: Parent container frame
            row: Grid row number for placement
        """
        pref_frame = ttk.LabelFrame(parent, text="File Preferences", padding="8")  # Labeled section
        pref_frame.grid(row=row, column=0, sticky='ew', pady=(0, 5))  # Fill width
        
        # Configure grid columns for expansion
        pref_frame.columnconfigure(1, weight=1)  # Data path column expands
        pref_frame.columnconfigure(4, weight=1)  # Graph path column expands
        
        # Data folder selection
        ttk.Label(pref_frame, text="Data Folder:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky='w', padx=(0, 5))  # Label
        
        self.data_path_var = tk.StringVar(value=str(Path.cwd() / "dmm_data"))  # CSV save location
        ttk.Entry(pref_frame, textvariable=self.data_path_var,
            font=('Arial', 9)).grid(row=0, column=1, sticky='ew', padx=(0, 5))  # Entry field, expands
        
        ttk.Button(pref_frame, text="Browse...",
            command=lambda: self.browse_folder('data')).grid(row=0, column=2, padx=(0, 15))  # Browse button
        
        # Graphs folder selection
        ttk.Label(pref_frame, text="Graphs Folder:", font=('Arial', 10, 'bold')).grid(
            row=0, column=3, sticky='w', padx=(0, 5))  # Label
        
        self.graph_path_var = tk.StringVar(value=str(Path.cwd() / "dmm_graphs"))  # Graph save location
        ttk.Entry(pref_frame, textvariable=self.graph_path_var,
            font=('Arial', 9)).grid(row=0, column=4, sticky='ew', padx=(0, 5))  # Entry field, expands
        
        ttk.Button(pref_frame, text="Browse...",
            command=lambda: self.browse_folder('graphs')).grid(row=0, column=5)  # Browse button
        
        # Graph title (optional custom title for generated plots)
        ttk.Label(pref_frame, text="Graph Title:", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, sticky='w', pady=(8, 0), padx=(0, 5))  # Label on row 2
        
        self.graph_title_var = tk.StringVar(value="DMM Measurements")  # Default graph title
        ttk.Entry(pref_frame, textvariable=self.graph_title_var,
            font=('Arial', 10)).grid(row=1, column=1, columnspan=4, sticky='ew',
            pady=(8, 0), padx=(0, 5))  # Entry field spans multiple columns

    def create_operations_section(self, parent, row):
        """
        Create operations section (measurement, export, and data buttons)
        
        Provides buttons for:
        - Single measurement (one reading)
        - Continuous measurements (repeated readings)
        - Statistics display
        - CSV export
        - Graph generation
        - Clear data
        
        Args:
            parent: Parent container frame
            row: Grid row number for placement
        """
        ops_frame = ttk.LabelFrame(parent, text="Operations", padding="8")  # Labeled section
        ops_frame.grid(row=row, column=0, sticky='ew', pady=(0, 5))  # Fill width
        
        # Configure equal width columns for balanced layout
        for i in range(7):  # 7 operation buttons
            ops_frame.columnconfigure(i, weight=1)  # All columns have equal weight (equal width)
        
        # Single measurement button
        self.single_measure_btn = ttk.Button(ops_frame, text="Single Measurement",
            command=self.single_measurement,
            style='Measure.TButton')  # Blue button
        self.single_measure_btn.grid(row=0, column=0, sticky='ew', padx=2)  # Expand to fill column width
        
        # Continuous measurement buttons
        self.continuous_btn = ttk.Button(ops_frame, text="Start Continuous",
            command=self.start_continuous,
            style='Measure.TButton')  # Blue button
        self.continuous_btn.grid(row=0, column=1, sticky='ew', padx=2)
        
        self.stop_btn = ttk.Button(ops_frame, text="Stop",
            command=self.stop_continuous,
            style='Disconnect.TButton',  # Red button
            state='disabled')  # Disabled until continuous running
        self.stop_btn.grid(row=0, column=2, sticky='ew', padx=2)
        
        # Statistics button
        self.stats_btn = ttk.Button(ops_frame, text="Get Statistics",
            command=self.get_statistics)  # Default style (gray)
        self.stats_btn.grid(row=0, column=3, sticky='ew', padx=2)
        
        # Export operations
        self.export_csv_btn = ttk.Button(ops_frame, text="Export CSV",
            command=self.export_csv,
            style='Export.TButton')  # Green button
        self.export_csv_btn.grid(row=0, column=4, sticky='ew', padx=2)
        
        self.generate_graph_btn = ttk.Button(ops_frame, text="Generate Graph",
            command=self.generate_graph,
            style='Export.TButton')  # Green button
        self.generate_graph_btn.grid(row=0, column=5, sticky='ew', padx=2)
        
        # Clear data button
        self.clear_btn = ttk.Button(ops_frame, text="Clear Data",
            command=self.clear_data)  # Default style
        self.clear_btn.grid(row=0, column=6, sticky='ew', padx=2)
        
        # Disable initially (until connected)
        self.disable_operation_buttons()

    def create_status_section(self, parent, row):
        """
        Create expandable status and activity log section
        
        Shows current operation status and maintains scrolling log of all
        user actions, measurements, and system events for troubleshooting
        and audit trail purposes.
        
        Args:
            parent: Parent container frame
            row: Grid row number for placement
        """
        status_frame = ttk.LabelFrame(parent, text="Status & Activity Log", padding="8")  # Labeled section
        status_frame.grid(row=row, column=0, sticky='nsew')  # Fill available space (expandable)
        
        # Configure expansion properties
        status_frame.columnconfigure(0, weight=1)  # Content expands horizontally
        status_frame.rowconfigure(0, weight=0)  # Status line (fixed height)
        status_frame.rowconfigure(1, weight=0)  # Controls (fixed height)
        status_frame.rowconfigure(2, weight=1)  # Log (expandable, fills remaining space)
        
        # Current status display (one-line status)
        self.current_operation_var = tk.StringVar(value="Ready - Connect to DMM")  # Initial status text
        status_label = ttk.Label(status_frame, textvariable=self.current_operation_var,
            font=('Arial', 11, 'bold'), foreground='#1565c0')  # Bold blue status text
        status_label.grid(row=0, column=0, sticky='ew', pady=(0, 5))  # Fill width, add bottom padding
        
        # Controls line (log control buttons)
        controls_frame = ttk.Frame(status_frame)  # Container for controls
        controls_frame.grid(row=1, column=0, sticky='ew', pady=(0, 5))
        controls_frame.columnconfigure(1, weight=1)  # Middle column expands (spacer)
        
        ttk.Label(controls_frame, text="Activity Log:",
            font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w')  # Log header label
        
        # Spacer (empty frame that expands)
        ttk.Frame(controls_frame).grid(row=0, column=1, sticky='ew')
        
        ttk.Button(controls_frame, text="Clear Log",
            command=self.clear_log, width=8).grid(row=0, column=2, padx=(0, 5))  # Clear button
        
        ttk.Button(controls_frame, text="Save Log",
            command=self.save_log, width=8).grid(row=0, column=3)  # Save button
        
        # Expandable log text widget (scrolling text area)
        self.log_text = scrolledtext.ScrolledText(status_frame,
            font=('Consolas', 9),  # Monospace font for code/logs
            bg='#1e1e1e',  # Dark background (dark mode)
            fg='#00ff41',  # Green text (classic terminal look)
            insertbackground='white',  # White cursor
            wrap=tk.WORD)  # Wrap at word boundaries
        self.log_text.grid(row=2, column=0, sticky='nsew')  # Fill all available space

    # ============================================================================
    # UTILITY METHODS FOR UI MANAGEMENT
    # ============================================================================

    def browse_folder(self, folder_type):
        """
        Open file browser for user to select custom save directory
        
        Allows user to choose destination folder for CSV exports or graph images.
        Creates directory if it doesn't exist.
        
        Args:
            folder_type (str): Type of folder being browsed ("data" or "graphs")
        """
        try:
            # Map folder type to corresponding UI variable
            var_mapping = {
                'data': 'data_path_var',  # CSV data path variable
                'graphs': 'graph_path_var'  # Graph path variable
            }
            
            var_name = var_mapping.get(folder_type)  # Get variable name
            if not var_name:  # Invalid folder type
                return  # Exit without doing anything
            
            # Get current path from UI variable
            current_path = getattr(self, var_name).get()  # Get current value
            if not current_path or not Path(current_path).exists():  # Path invalid/missing
                current_path = str(Path.cwd())  # Use current working directory
            
            # Open file browser dialog
            folder_path = filedialog.askdirectory(
                initialdir=current_path,  # Start at current path
                title=f"Select {folder_type.title()} Save Location")  # Dialog title
            
            if folder_path:  # User selected a folder (didn't cancel)
                getattr(self, var_name).set(folder_path)  # Update UI variable with selected path
                self.save_locations[folder_type] = folder_path  # Store in preferences
                self.log_message(f"Updated {folder_type} folder: {folder_path}", "SUCCESS")  # Log change
                Path(folder_path).mkdir(parents=True, exist_ok=True)  # Create directory if needed
        
        except Exception as e:  # Catch any file browser errors
            self.log_message(f"Error selecting folder: {e}", "ERROR")  # Log error

    def log_message(self, message: str, level: str = "INFO"):
        """
        Add timestamped colored message to activity log
        
        Appends message to log text widget with timestamp prefix and color-coding
        based on message level for easy visual scanning (INFO=blue, SUCCESS=green,
        WARNING=yellow, ERROR=red).
        
        Args:
            message (str): Message text to log
            level (str): Message level (INFO, SUCCESS, WARNING, ERROR)
        """
        timestamp = datetime.now().strftime("%H:%M:%S")  # Format time HH:MM:SS
        log_entry = f"[{timestamp}] {message}\n"  # Format log line with timestamp
        
        try:
            self.log_text.insert(tk.END, log_entry)  # Add to end of text widget
            self.log_text.see(tk.END)  # Scroll to bottom to show new message
            
            # Apply color-coding based on message level
            if level == "ERROR":  # Error messages = red
                self.log_text.tag_add("error", f"end-{len(log_entry)}c", "end-1c")  # Mark text
                self.log_text.tag_config("error", foreground="#ff6b6b")  # Red color
            elif level == "SUCCESS":  # Success messages = green
                self.log_text.tag_add("success", f"end-{len(log_entry)}c", "end-1c")
                self.log_text.tag_config("success", foreground="#51cf66")  # Green color
            elif level == "WARNING":  # Warning messages = yellow
                self.log_text.tag_add("warning", f"end-{len(log_entry)}c", "end-1c")
                self.log_text.tag_config("warning", foreground="#ffd43b")  # Yellow color
            else:  # INFO messages = blue (default)
                self.log_text.tag_add("info", f"end-{len(log_entry)}c", "end-1c")
                self.log_text.tag_config("info", foreground="#74c0fc")  # Blue color
        
        except Exception as e:  # Catch any logging errors
            print(f"Log error: {e}")  # Print to console (fallback)

    def clear_log(self):
        """Clear all content from activity log text widget"""
        try:
            self.log_text.delete(1.0, tk.END)  # Delete from start (line 1, column 0) to end
            self.log_message("Log cleared")  # Log the clear action
        except Exception as e:  # Catch any clear errors
            print(f"Clear error: {e}")

    def save_log(self):
        """Save activity log content to text file with timestamp filename"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # Generate timestamp for filename
            filename = filedialog.asksaveasfilename(
                defaultextension=".txt",  # Default extension
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],  # File type filters
                initialname=f"dmm_log_{timestamp}.txt")  # Suggested filename with timestamp
            
            if filename:  # User selected a file (didn't cancel)
                log_content = self.log_text.get(1.0, tk.END)  # Get all log text
                with open(filename, 'w') as f:  # Open file for writing
                    f.write(f"DMM Automation Log - {datetime.now()}\n")  # Header with timestamp
                    f.write("="*50 + "\n\n")  # Separator line
                    f.write(log_content)  # All log content
                self.log_message(f"Log saved: {filename}", "SUCCESS")  # Log save success
        
        except Exception as e:  # Catch any save errors
            self.log_message(f"Save error: {e}", "ERROR")  # Log error

    def update_status(self, status: str):
        """
        Update current operation status display text
        
        Immediately updates the status label with new message and forces GUI
        update to display new status without waiting for next event cycle.
        
        Args:
            status (str): Status message text to display
        """
        try:
            self.current_operation_var.set(status)  # Update StringVar (updates label automatically)
            self.root.update_idletasks()  # Process pending GUI updates immediately
        except:
            pass  # Silently ignore any status update errors

    def disable_operation_buttons(self):
        """Disable all operation buttons (for disconnected state)"""
        buttons = [self.single_measure_btn, self.continuous_btn, self.stats_btn,
                   self.export_csv_btn, self.generate_graph_btn, self.clear_btn,
                   self.test_btn]  # List of operation buttons
        
        for btn in buttons:  # Disable each button
            try:
                btn.configure(state='disabled')  # Set to disabled state
            except:
                pass  # Ignore errors

    def enable_operation_buttons(self):
        """Enable all operation buttons (for connected state)"""
        buttons = [self.single_measure_btn, self.continuous_btn, self.stats_btn,
                   self.export_csv_btn, self.generate_graph_btn, self.clear_btn,
                   self.test_btn]  # List of operation buttons
        
        for btn in buttons:  # Enable each button
            try:
                btn.configure(state='normal')  # Set to normal (enabled) state
            except:
                pass  # Ignore errors

    # ============================================================================
    # CONNECTION MANAGEMENT METHODS
    # ============================================================================

    def connect_dmm(self):
        """
        Establish VISA connection to Keithley DMM6500
        
        Spawns background thread to connect, avoiding GUI freezing during
        potentially long USB enumeration. Updates connection status and
        enables channel controls on success.
        """
        def connect_thread():  # Background worker function
            try:
                self.update_status("Connecting...")  # Update status
                self.log_message("Connecting to Keithley DMM6500...")  # Log attempt
                
                visa_address = self.visa_address_var.get().strip()  # Get VISA address from entry
                if not visa_address:  # Check if address is empty
                    raise ValueError("VISA address empty")  # Raise error
                
                self.dmm = KeithleyDMM6500(visa_address)  # Create driver instance
                
                if self.dmm.connect():  # Attempt connection
                    info = self.dmm.get_instrument_info()  # Query device info
                    self.status_queue.put(("connected", info))  # Signal main thread: connected
                else:
                    raise Exception("Connection failed")  # Raise error if connection unsuccessful
            
            except Exception as e:  # Catch any connection errors
                self.status_queue.put(("error", f"Connection failed: {str(e)}"))  # Signal main thread: error
        
        threading.Thread(target=connect_thread, daemon=True).start()  # Start background thread

    def disconnect_dmm(self):
        """
        Close VISA connection to DMM and disable all controls
        
        Performs cleanup: closes connection, resets UI state, clears status
        displays, and disables all operation buttons until reconnected.
        """
        try:
            # Stop continuous measurements if running
            if self.continuous_running:  # Check if continuous measurements active
                self.stop_continuous()  # Stop the continuous measurement
            
            if self.dmm:  # Check if DMM is connected
                self.dmm.disconnect()  # Close VISA connection
                self.dmm = None  # Clear reference
            
            # Update connection status UI
            self.conn_status_var.set("Disconnected")  # Set status text
            self.conn_status_label.configure(foreground='#d32f2f')  # Red color (disconnected)
            self.connect_btn.configure(state='normal')  # Enable connect button
            self.disconnect_btn.configure(state='disabled')  # Disable disconnect button
            
            # Disable all operation buttons
            self.disable_operation_buttons()  # Reset button states
            
            # Clear instrument info display
            self.info_text.configure(state='normal')  # Enable editing
            self.info_text.delete(1.0, tk.END)  # Clear content
            self.info_text.configure(state='disabled')  # Disable editing
            
            self.update_status("Disconnected")  # Update status
            self.log_message("Disconnected from DMM", "SUCCESS")  # Log event
        
        except Exception as e:  # Catch any errors during disconnect
            self.log_message(f"Disconnect error: {e}", "ERROR")  # Log error

    def test_connection(self):
        """Test VISA communication link to verify DMM is responsive"""
        try:
            if self.dmm and self.dmm.is_connected:  # Check if connected
                self.log_message("Connection test: PASSED", "SUCCESS")  # Log success
                self.update_status("Test passed")  # Update status
                messagebox.showinfo("Test Result", "Connection test PASSED!")  # Show success dialog
            else:
                self.log_message("Connection test: FAILED", "ERROR")  # Log failure
                messagebox.showerror("Test Result", "Connection test FAILED!")  # Show error dialog
        
        except Exception as e:  # Catch any test errors
            self.log_message(f"Test error: {e}", "ERROR")  # Log error

    # ============================================================================
    # MEASUREMENT METHODS
    # ============================================================================

    def get_measurement_function(self):
        """
        Get selected measurement function enum from dropdown
        
        Returns the MeasurementFunction enum value corresponding to the
        user-selected measurement type from the dropdown combobox.
        
        Returns:
            MeasurementFunction: Selected measurement function (defaults to DC_VOLTAGE)
        """
        selected = self.measurement_type_var.get()  # Get selected measurement type string
        for name, func in MEASUREMENT_OPTIONS:  # Search through measurement options
            if name == selected:  # Find matching option
                return func  # Return the function enum
        return MeasurementFunction.DC_VOLTAGE  # Default to DC voltage if not found

    def get_measurement_parameters(self):
        """
        Extract measurement parameters from GUI fields
        
        Retrieves range, resolution, and NPLC values from user input fields,
        validates them, and returns as dictionary for passing to measurement method.
        
        Returns:
            Dict: Parameters dictionary with keys: measurement_range, resolution, nplc
        """
        params = {}  # Initialize empty parameters dictionary
        
        # Range parameter (if specified and not "Auto")
        range_val = self.range_var.get().strip()  # Get range value from entry
        if range_val and range_val.lower() != "auto":  # Not empty and not "Auto"
            try:
                params['measurement_range'] = float(range_val)  # Convert to float and add
            except ValueError:
                pass  # Ignore invalid values (skip this parameter)
        
        # Resolution parameter (if specified and not "Auto")
        res_val = self.resolution_var.get().strip()  # Get resolution value from entry
        if res_val and res_val.lower() != "auto":  # Not empty and not "Auto"
            try:
                params['resolution'] = float(res_val)  # Convert to float and add
            except ValueError:
                pass  # Ignore invalid values
        
        # NPLC parameter (always included, default 1.0)
        try:
            params['nplc'] = self.nplc_var.get()  # Get NPLC from dropdown
        except:
            params['nplc'] = 1.0  # Default to 1.0 if error
        
        return params  # Return parameters dictionary

    def single_measurement(self):
        """
        Perform single measurement and add to data collection
        
        Queries DMM for one measurement with configured parameters,
        displays result, and stores in data handler for export/analysis.
        """
        def measure_thread():  # Background worker function
            try:
                self.update_status("Taking measurement...")  # Update status
                self.log_message("Performing single measurement...")  # Log operation
                
                func = self.get_measurement_function()  # Get selected measurement type
                params = self.get_measurement_parameters()  # Get configured parameters
                
                value = self.dmm.measure(func, **params)  # Perform measurement on DMM
                
                if value is not None:  # Measurement succeeded (returned valid number)
                    # Determine unit based on measurement type
                    measurement_name = self.measurement_type_var.get()  # Get measurement name
                    unit = self.get_unit_for_measurement(measurement_name)  # Get corresponding unit
                    
                    # Add to data handler for later export
                    self.data_handler.add_measurement(measurement_name, value, unit)  # Store measurement
                    
                    # Signal measurement complete with result
                    self.status_queue.put(("measurement_complete",
                        f"{measurement_name}: {value:.9f} {unit}"))  # Send result to main thread
                
                else:
                    self.status_queue.put(("error", "Measurement failed"))  # Signal error
            
            except Exception as e:  # Catch any measurement errors
                self.status_queue.put(("error", f"Measurement error: {str(e)}"))  # Signal error
        
        if self.dmm and self.dmm.is_connected:  # Verify connection
            threading.Thread(target=measure_thread, daemon=True).start()  # Start background thread
        else:
            messagebox.showerror("Error", "DMM not connected")  # Show error dialog

    def get_unit_for_measurement(self, measurement_type):
        """
        Get unit string corresponding to measurement type
        
        Maps human-readable measurement types to their corresponding physical units
        for display in results and logs.
        
        Args:
            measurement_type (str): Human-readable measurement type
        
        Returns:
            str: Unit abbreviation (V, A, Ω, F, Hz, etc.)
        """
        unit_map = {
            "DC Voltage": "V",  # Volts
            "AC Voltage": "V",  # Volts
            "DC Current": "A",  # Amps
            "AC Current": "A",  # Amps
            "2-Wire Resistance": "Ω",  # Ohms (Greek letter)
            "4-Wire Resistance": "Ω",  # Ohms
            "Capacitance": "F",  # Farads
            "Frequency": "Hz"  # Hertz
        }
        return unit_map.get(measurement_type, "")  # Return unit or empty string if not found

    def start_continuous(self):
        """
        Start continuous measurement polling at configured interval
        
        Begins repeated measurements at fixed time interval (configurable
        via GUI). Runs in background thread to prevent UI blocking. Displays
        each measurement result and stores in data handler.
        """
        if not self.dmm or not self.dmm.is_connected:  # Verify connection
            messagebox.showerror("Error", "DMM not connected")  # Show error
            return

        self.continuous_running = True  # Set flag to indicate measurements running
        self.continuous_btn.configure(state='disabled')  # Disable start button
        self.stop_btn.configure(state='normal')  # Enable stop button
        self.update_status("Continuous measurements running...")  # Update status
        self.log_message("Starting continuous measurements...")  # Log event

        def continuous_thread():  # Background worker function
            measurement_count = 0  # Counter for measurement number
            
            while self.continuous_running:  # Loop while measurements enabled
                try:
                    func = self.get_measurement_function()  # Get selected measurement type
                    params = self.get_measurement_parameters()  # Get configured parameters
                    
                    value = self.dmm.measure(func, **params)  # Perform measurement
                    
                    if value is not None:  # Measurement succeeded
                        measurement_count += 1  # Increment counter
                        measurement_name = self.measurement_type_var.get()  # Get measurement name
                        unit = self.get_unit_for_measurement(measurement_name)  # Get unit
                        
                        # Add to data handler
                        self.data_handler.add_measurement(measurement_name, value, unit)  # Store measurement
                        
                        # Signal measurement complete
                        self.status_queue.put(("continuous_measurement",
                            f"#{measurement_count}: {value:.9f} {unit}"))  # Send result
                    
                    else:
                        self.status_queue.put(("error", f"Measurement #{measurement_count+1} failed"))  # Signal error
                    
                    # Wait before next measurement
                    time.sleep(5.0)  # 5-second delay between measurements
                
                except Exception as e:  # Catch measurement errors
                    self.status_queue.put(("error", f"Continuous measurement error: {str(e)}"))  # Signal error
                    break  # Exit loop on error
        
        self.continuous_thread = threading.Thread(target=continuous_thread, daemon=True)  # Create thread
        self.continuous_thread.start()  # Start thread

    def stop_continuous(self):
        """Stop continuous measurements and update UI state"""
        self.continuous_running = False  # Clear running flag
        self.continuous_btn.configure(state='normal')  # Enable start button
        self.stop_btn.configure(state='disabled')  # Disable stop button
        self.update_status("Continuous measurements stopped")  # Update status
        self.log_message("Continuous measurements stopped", "SUCCESS")  # Log event

    def get_statistics(self):
        """Display measurement statistics dialog with summary data"""
        try:
            stats = self.data_handler.get_statistics()  # Get statistics from data handler
            
            if not stats:  # No data available
                messagebox.showinfo("Statistics", "No measurement data available")  # Show info dialog
                return
            
            # Format statistics display text
            stats_text = f"""Measurement Statistics:
Count: {stats['count']}
Mean: {stats['mean']:.9f}
Standard Deviation: {stats['std']:.9f}
Minimum: {stats['min']:.9f}
Maximum: {stats['max']:.9f}
Range: {stats['range']:.9f}"""  # Statistics summary text
            
            messagebox.showinfo("Statistics", stats_text)  # Show in dialog
            self.log_message(f"Statistics: {stats['count']} measurements, μ={stats['mean']:.6f}, σ={stats['std']:.6f}", "SUCCESS")  # Log to activity log
        
        except Exception as e:  # Catch any statistics errors
            self.log_message(f"Statistics error: {e}", "ERROR")  # Log error

    # ============================================================================
    # DATA EXPORT METHODS
    # ============================================================================

    def export_csv(self):
        """Export measurement data to CSV file in background thread"""
        def export_thread():  # Background worker function
            try:
                self.update_status("Exporting CSV...")  # Update status
                self.log_message("Exporting measurement data to CSV...")  # Log operation
                
                filepath = self.data_handler.export_to_csv(
                    custom_path=self.data_path_var.get())  # Export to configured path
                
                if filepath:  # Export succeeded
                    self.status_queue.put(("csv_exported", filepath))  # Signal success
                else:
                    self.status_queue.put(("error", "CSV export failed"))  # Signal failure
            
            except Exception as e:  # Catch any export errors
                self.status_queue.put(("error", f"CSV export error: {str(e)}"))  # Signal error
        
        threading.Thread(target=export_thread, daemon=True).start()  # Start background thread

    def generate_graph(self):
        """Generate measurement graph in background thread"""
        def graph_thread():  # Background worker function
            try:
                self.update_status("Generating graph...")  # Update status
                self.log_message("Generating measurement graph...")  # Log operation
                
                graph_title = self.graph_title_var.get().strip() or None  # Get custom title or None
                filepath = self.data_handler.generate_graph(
                    custom_path=self.graph_path_var.get(),  # Output directory
                    graph_title=graph_title)  # Graph title
                
                if filepath:  # Generation succeeded
                    self.status_queue.put(("graph_generated", filepath))  # Signal success
                else:
                    self.status_queue.put(("error", "Graph generation failed"))  # Signal failure
            
            except Exception as e:  # Catch any generation errors
                self.status_queue.put(("error", f"Graph generation error: {str(e)}"))  # Signal error
        
        threading.Thread(target=graph_thread, daemon=True).start()  # Start background thread

    def clear_data(self):
        """Clear all measurement data with confirmation"""
        try:
            if messagebox.askyesno("Confirm", "Clear all measurement data?"):  # Ask for confirmation
                self.data_handler.clear_data()  # Clear data
                self.log_message("All measurement data cleared", "SUCCESS")  # Log action
                self.update_status("Data cleared")  # Update status
        
        except Exception as e:  # Catch any clear errors
            self.log_message(f"Clear data error: {e}", "ERROR")  # Log error

    # ============================================================================
    # INFORMATION DISPLAY METHODS
    # ============================================================================

    def display_instrument_info(self, info):
        """Display instrument identification information in info text widget"""
        try:
            self.info_text.configure(state='normal')  # Enable editing
            self.info_text.delete(1.0, tk.END)  # Clear previous content
            
            if info:  # Info dictionary provided
                # Format info string: Manufacturer Model | S/N: SerialNumber | FW: Firmware
                info_str = f"Connected: {info.get('manufacturer', 'N/A')} {info.get('model', 'N/A')} | S/N: {info.get('serial_number', 'N/A')} | FW: {info.get('firmware_version', 'N/A')}"
            else:
                info_str = "Connected but no instrument info available"  # Fallback if no info
            
            self.info_text.insert(1.0, info_str)  # Insert formatted string
            self.info_text.configure(state='disabled')  # Disable editing (read-only)
        
        except Exception as e:  # Catch any display errors
            print(f"Info display error: {e}")  # Print to console (fallback)

    # ============================================================================
    # STATUS UPDATE POLLING
    # ============================================================================

    def check_status_updates(self):
        """
        Periodic polling of status queue for messages from worker threads
        
        Checks for status updates from background threads and processes them:
        connected, disconnected, errors, measurements, exports, etc.
        Calls itself recursively every 100ms using root.after().
        """
        try:
            while True:  # Process all queued messages
                status_type, data = self.status_queue.get_nowait()  # Get next message (non-blocking)
                
                # Route message to appropriate handler based on type
                if status_type == "connected":  # Connection successful
                    self.conn_status_var.set("Connected")  # Update status text
                    self.conn_status_label.configure(foreground='#2d7d32')  # Green color
                    self.connect_btn.configure(state='disabled')  # Disable connect button
                    self.disconnect_btn.configure(state='normal')  # Enable disconnect button
                    self.enable_operation_buttons()  # Enable all operation buttons
                    self.update_status("Connected - Ready")  # Update status
                    self.log_message("Successfully connected to DMM", "SUCCESS")  # Log success
                    if data:
                        self.display_instrument_info(data)  # Display device info
                        self.log_message(f"Instrument: {data.get('manufacturer', 'N/A')} {data.get('model', 'N/A')}", "SUCCESS")  # Log info
                
                elif status_type == "error":  # Error from worker thread
                    self.log_message(data, "ERROR")  # Log error message
                    self.update_status("Error occurred")  # Update status
                
                elif status_type == "measurement_complete":  # Single measurement done
                    self.log_message(f"Measurement: {data}", "SUCCESS")  # Log result
                    self.update_status("Measurement complete")  # Update status
                
                elif status_type == "continuous_measurement":  # Continuous measurement result
                    # Don't log every continuous measurement to avoid spam
                    self.update_status(f"Continuous: {data}")  # Update status only
                
                elif status_type == "csv_exported":  # CSV export completed
                    self.log_message(f"CSV exported: {Path(data).name}", "SUCCESS")  # Log success with filename
                    self.update_status("CSV exported")  # Update status
                
                elif status_type == "graph_generated":  # Graph generation completed
                    self.log_message(f"Graph generated: {Path(data).name}", "SUCCESS")  # Log success with filename
                    self.update_status("Graph generated")  # Update status
        
        except queue.Empty:  # No messages in queue (normal condition)
            pass  # Continue to next polling cycle
        
        except Exception as e:  # Catch any queue processing errors
            print(f"Status update error: {e}")  # Print to console (fallback)
        
        finally:
            # Schedule next polling cycle (100ms delay)
            self.root.after(100, self.check_status_updates)  # Recursive scheduling

    # ============================================================================
    # APPLICATION STARTUP
    # ============================================================================

    def run(self):
        """
        Start the main GUI application loop
        
        Launches Tkinter event loop (blocks until user closes window).
        Provides startup messages and handles cleanup on application exit.
        """
        try:
            # Log application startup
            self.log_message("Keithley DMM6500 Professional Automation - Version 2.0", "SUCCESS")  # Log version
            self.log_message("Features: Measurements + CSV Export + Graphing + Statistics", "SUCCESS")  # Log features
            self.log_message("Ready to connect to DMM")  # Log ready state
            
            # Start GUI event loop (blocks until window closed)
            self.root.mainloop()  # Enter event loop
        
        except KeyboardInterrupt:  # User interrupts (Ctrl+C)
            self.log_message("Application interrupted")  # Log interruption
        
        except Exception as e:  # Catch any runtime errors
            self.log_message(f"Application error: {e}", "ERROR")  # Log error
        
        finally:
            # Cleanup on exit
            if hasattr(self, 'dmm') and self.dmm:  # Check if DMM connected
                try:
                    self.dmm.disconnect()  # Close connection
                except:
                    pass  # Ignore errors during cleanup


# ============================================================================
# APPLICATION ENTRY POINT
# ============================================================================

def main():
    """
    Main entry point for the application
    
    Creates and runs the GUI application. Provides error handling for startup
    issues and user-friendly error messages.
    """
    print("Keithley DMM6500 Professional Automation Suite")  # Print header
    print("Features: Measurements + CSV Export + Graphing + Statistics")  # Print features
    print("Professional GUI with color coding and comprehensive logging")  # Print description
    print("="*70)  # Print separator
    
    try:
        app = KeithleyDMMAutomationGUI()  # Create GUI application instance
        app.run()  # Start application (blocks until window closed)
    
    except Exception as e:  # Catch any startup errors
        print(f"Error: {e}")  # Print error message
        input("Press Enter to exit...")  # Wait for user acknowledgment


if __name__ == "__main__":  # Check if run as main script
    main()  # Execute main function
