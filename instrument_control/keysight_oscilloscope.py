"""
MEASUREMENT FEATURE: Keysight DSOX6004A Oscilloscope Measurement Functions
Provides automatic waveform analysis with multiple measurement types
"""

import logging
import time
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any, List, Tuple, Union
import numpy as np
from .scpi_wrapper import SCPIWrapper

class KeysightDSOX6004AError(Exception):
    """Custom exception for Keysight DSOX6004A oscilloscope errors."""
    pass

class KeysightDSOX6004A:
    """Keysight DSOX6004A Oscilloscope Control Class with Measurement Features"""

    def __init__(self, visa_address: str, timeout_ms: int = 10000) -> None:
        """Initialize oscilloscope connection parameters"""
        self._scpi_wrapper = SCPIWrapper(visa_address, timeout_ms)
        self._logger = logging.getLogger(f'{self.__class__.__name__}.{id(self)}')
        self.max_channels = 4
        self.max_sample_rate = 20e9
        self.max_memory_depth = 16e6
        self.bandwidth_hz = 1e9
        self._valid_vertical_scales = [
            1e-3, 2e-3, 5e-3, 10e-3, 20e-3, 50e-3,
            100e-3, 200e-3, 500e-3, 1.0, 2.0, 5.0, 10.0
        ]
        self._valid_timebase_scales = [
            1e-12, 2e-12, 5e-12, 10e-12, 20e-12, 50e-12,
            100e-12, 200e-12, 500e-12, 1e-9, 2e-9, 5e-9,
            10e-9, 20e-9, 50e-9, 100e-9, 200e-9, 500e-9,
            1e-6, 2e-6, 5e-6, 10e-6, 20e-6, 50e-6,
            100e-6, 200e-6, 500e-6, 1e-3, 2e-3, 5e-3,
            10e-3, 20e-3, 50e-3, 100e-3, 200e-3, 500e-3,
            1.0, 2.0, 5.0, 10.0, 20.0, 50.0
        ]
        # self._measurement_types = [
        #     "FREQ", "PERiod", "VTOP", "VBASe", "VAMP", "VAVG", "VRMS",
        #     "VMIN", "VMAX", "VPPK", "VRISE", "VFALL", "PWIDTH", "NWIDTH",
        #     "PDUTy", "NDUTy", "RISE", "FALL", "DELAY"
        # ]
        self._measurement_types = [
            "FREQ",      # Frequency
            "PERiod",    # Period (correct: PER-iod)
            "VAMP",      # Amplitude
            "VAVG",      # Average
            "VRMS",      # RMS
            "VMAX",      # Maximum
            "VMIN",      # Minimum
            "RISE",      # Rise time
            "FALL",      # Fall time
            "PDUTy",     # Positive duty (correct: PDU-Ty)
            "NDUTy"      # Negative duty (correct: NDU-Ty)
        ]

    def connect(self) -> bool:
        """Establish VISA connection to oscilloscope"""
        if self._scpi_wrapper.connect():
            try:
                identification = self._scpi_wrapper.query("*IDN?")
                self._logger.info(f"Instrument identification: {identification.strip()}")
                self._scpi_wrapper.write("*CLS")
                time.sleep(0.5)
                self._scpi_wrapper.query("*OPC?")
                self._logger.info("Successfully connected to Keysight DSOX6004A")
                return True
            except Exception as e:
                self._logger.error(f"Error during instrument identification: {e}")
                self._scpi_wrapper.disconnect()
                return False
        return False

    def disconnect(self) -> None:
        """Close connection to oscilloscope"""
        self._scpi_wrapper.disconnect()
        self._logger.info("Disconnection completed")

    @property
    def is_connected(self) -> bool:
        """Check if oscilloscope is currently connected"""
        return self._scpi_wrapper.is_connected

    def get_instrument_info(self) -> Optional[Dict[str, Any]]:
        """Query instrument identification and specifications"""
        if not self.is_connected:
            return None
        try:
            idn = self._scpi_wrapper.query("*IDN?").strip()
            parts = idn.split(',')
            return {
                'manufacturer': parts[0] if len(parts) > 0 else 'Unknown',
                'model': parts[1] if len(parts) > 1 else 'Unknown',
                'serial_number': parts[2] if len(parts) > 2 else 'Unknown',
                'firmware_version': parts[3] if len(parts) > 3 else 'Unknown',
                'max_channels': self.max_channels,
                'bandwidth_hz': self.bandwidth_hz,
                'max_sample_rate': self.max_sample_rate,
                'max_memory_depth': self.max_memory_depth,
                'identification': idn
            }
        except Exception as e:
            self._logger.error(f"Failed to get instrument info: {e}")
            return None

    def configure_channel(self, channel: int, vertical_scale: float, vertical_offset: float = 0.0, 
                         coupling: str = "DC", probe_attenuation: float = 1.0) -> bool:
        """Configure vertical parameters for specified channel"""
        if not self.is_connected:
            raise KeysightDSOX6004AError("Oscilloscope not connected")
        if not (1 <= channel <= self.max_channels):
            raise ValueError(f"Channel must be 1-{self.max_channels}, got {channel}")
        try:
            self._scpi_wrapper.write(f":CHANnel{channel}:DISPlay ON")
            time.sleep(0.05)
            self._scpi_wrapper.write(f":CHANnel{channel}:SCALe {vertical_scale}")
            time.sleep(0.05)
            self._scpi_wrapper.write(f":CHANnel{channel}:OFFSet {vertical_offset}")
            time.sleep(0.05)
            self._scpi_wrapper.write(f":CHANnel{channel}:COUPling {coupling}")
            time.sleep(0.05)
            self._scpi_wrapper.write(f":CHANnel{channel}:PROBe {probe_attenuation}")
            time.sleep(0.05)
            self._logger.info(f"Channel {channel} configured: Scale={vertical_scale}V/div, Offset={vertical_offset}V, Coupling={coupling}, Probe={probe_attenuation}x")
            return True
        except Exception as e:
            self._logger.error(f"Failed to configure channel {channel}: {e}")
            return False

    def configure_timebase(self, time_scale: float, time_offset: float = 0.0) -> bool:
        """Configure horizontal timebase settings"""
        if not self.is_connected:
            self._logger.error("Cannot configure timebase: oscilloscope not connected")
            return False
        if time_scale not in self._valid_timebase_scales:
            closest_scale = min(self._valid_timebase_scales, key=lambda x: abs(x - time_scale))
            self._logger.warning(f"Invalid timebase scale {time_scale}s, using {closest_scale}s")
            time_scale = closest_scale
        try:
            self._scpi_wrapper.write(f":TIMebase:SCALe {time_scale}")
            time.sleep(0.1)
            self._scpi_wrapper.write(f":TIMebase:OFFSet {time_offset}")
            time.sleep(0.1)
            self._logger.info(f"Timebase configured: Scale={time_scale}s/div, Offset={time_offset}s")
            return True
        except Exception as e:
            self._logger.error(f"Failed to configure timebase: {type(e).__name__}: {e}")
            return False

    def configure_trigger(self, channel: int, trigger_level: float, trigger_slope: str = "POS") -> bool:
        """Configure trigger settings"""
        if not self.is_connected:
            self._logger.error("Cannot configure trigger: oscilloscope not connected")
            return False
        if not (1 <= channel <= self.max_channels):
            raise ValueError(f"Channel must be 1-{self.max_channels}, got {channel}")
        valid_slopes = ["POS", "NEG"]
        if trigger_slope.upper() not in valid_slopes:
            raise ValueError(f"Trigger slope must be one of {valid_slopes}, got {trigger_slope}")
        try:
            self._scpi_wrapper.write(":TRIGger:MODE EDGE")
            time.sleep(0.1)
            self._scpi_wrapper.write(f":TRIGger:EDGE:SOURce CHANnel{channel}")
            time.sleep(0.1)
            self._scpi_wrapper.write(f":TRIGger:LEVel {trigger_level}")
            time.sleep(0.1)
            self._scpi_wrapper.write(f":TRIGger:EDGE:SLOPe {trigger_slope.upper()}")
            time.sleep(0.1)
            self._logger.info(f"Trigger configured: Channel={channel}, Level={trigger_level}V, Slope={trigger_slope}")
            return True
        except Exception as e:
            self._logger.error(f"Failed to configure trigger: {type(e).__name__}: {e}")
            return False

    # ============================================================================
    # NEW: MEASUREMENT FUNCTIONS
    # ============================================================================

    def measure_single(self, channel: int, measurement_type: str) -> Optional[float]:
        """
        Perform a single measurement on specified channel
        
        Args:
            channel (int): Channel number (1-4)
            measurement_type (str): Measurement type (FREQ, VAMP, VRMS, VMAX, VMIN, PERiod, etc.)
        
        Returns:
            float: Measured value or None if measurement fails
        
        Available Measurements:
            - FREQ: Frequency in Hz
            - PERiod: Period in seconds
            - VTOP: Top voltage (maximum)
            - VBASe: Base voltage (minimum)
            - VAMP: Amplitude (peak-to-peak)
            - VAVG: Average voltage
            - VRMS: RMS voltage
            - VMIN: Minimum voltage
            - VMAX: Maximum voltage
            - VPPK: Peak-to-peak voltage
            - RISE: Rise time
            - FALL: Fall time
            - PDUTy: Positive duty cycle
            - NDUTy: Negative duty cycle
        """
        if not self.is_connected:
            self._logger.error("Cannot measure: oscilloscope not connected")
            return None
        if not (1 <= channel <= self.max_channels):
            self._logger.error(f"Invalid channel: {channel}")
            return None
        
        # measurement_type = measurement_type.upper()
        # if measurement_type not in self._measurement_types:
        #     self._logger.error(f"Unknown measurement type: {measurement_type}. Valid: {self._measurement_types}")
        #     return None
        
        try:
            query_command = f":MEASure:{measurement_type}? CHANnel{channel}"
            self._scpi_wrapper.query("*OPC?")
            time.sleep(0.1)
            response = self._scpi_wrapper.query(query_command).strip()
            
            try:
                value = float(response)
                self._logger.debug(f"Ch{channel} {measurement_type}: {value}")
                return value
            except ValueError:
                self._logger.error(f"Failed to parse measurement response: '{response}'")
                return None
        except Exception as e:
            self._logger.error(f"Measurement failed for Ch{channel} ({measurement_type}): {e}")
            return None

    def measure_multiple(self, channel: int, measurement_types: List[str]) -> Optional[Dict[str, float]]:
        """
        Perform multiple measurements on specified channel
        
        Args:
            channel (int): Channel number (1-4)
            measurement_types (List[str]): List of measurement types to perform
        
        Returns:
            Dict[str, float]: Dictionary with measurement names as keys and values
        """
        if not self.is_connected:
            self._logger.error("Cannot measure: oscilloscope not connected")
            return None
        
        results = {}
        for meas_type in measurement_types:
            value = self.measure_single(channel, meas_type)
            if value is not None:
                results[meas_type] = value
        
        if results:
            self._logger.info(f"Ch{channel} measurements: {results}")
            return results
        else:
            self._logger.error(f"No measurements succeeded for Ch{channel}")
            return None

    def get_all_measurements(self, channel: int) -> Optional[Dict[str, float]]:
        """
        Get all available measurements for a channel
        
        Args:
            channel (int): Channel number (1-4)
        
        Returns:
            Dict[str, float]: All available measurements
        """
        essential_measurements = ["FREQ", "PERiod", "VAMP", "VAVG", "VRMS", "VMAX", "VMIN"]
        return self.measure_multiple(channel, essential_measurements)

    def measure_frequency(self, channel: int) -> Optional[float]:
        """Measure signal frequency in Hz"""
        return self.measure_single(channel, "FREQ")

    def measure_period(self, channel: int) -> Optional[float]:
        """Measure signal period in seconds"""
        return self.measure_single(channel, "PERiod")

    def measure_amplitude(self, channel: int) -> Optional[float]:
        """Measure signal amplitude (peak-to-peak) in volts"""
        return self.measure_single(channel, "VAMP")

    def measure_average(self, channel: int) -> Optional[float]:
        """Measure signal average voltage in volts"""
        return self.measure_single(channel, "VAVG")

    def measure_rms(self, channel: int) -> Optional[float]:
        """Measure signal RMS voltage in volts"""
        return self.measure_single(channel, "VRMS")

    def measure_max(self, channel: int) -> Optional[float]:
        """Measure maximum voltage in volts"""
        return self.measure_single(channel, "VMAX")

    def measure_min(self, channel: int) -> Optional[float]:
        """Measure minimum voltage in volts"""
        return self.measure_single(channel, "VMIN")

    def measure_rise_time(self, channel: int) -> Optional[float]:
        """Measure signal rise time in seconds"""
        return self.measure_single(channel, "RISE")

    def measure_fall_time(self, channel: int) -> Optional[float]:
        """Measure signal fall time in seconds"""
        return self.measure_single(channel, "FALL")

    def measure_duty_cycle_positive(self, channel: int) -> Optional[float]:
        """Measure positive duty cycle as percentage"""
        return self.measure_single(channel, "PDUTy")

    def measure_duty_cycle_negative(self, channel: int) -> Optional[float]:
        """Measure negative duty cycle as percentage"""
        return self.measure_single(channel, "NDUTy")

    # ============================================================================
    # REMAINING METHODS (configure_function_generator, capture_screenshot, etc.)
    # ============================================================================

    def capture_screenshot(self, filename: Optional[str] = None, image_format: str = "PNG", include_timestamp: bool = True) -> Optional[str]:
        """Capture oscilloscope display screenshot"""
        if not self.is_connected:
            self._logger.error("Cannot capture screenshot: not connected")
            return None
        try:
            self.setup_output_directories()
            if filename is None:
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                filename = f"scope_screenshot_{timestamp}.{image_format.lower()}"
            if not filename.lower().endswith(f".{image_format.lower()}"):
                filename += f".{image_format.lower()}"
            screenshot_path = self.screenshot_dir / filename
            self._scpi_wrapper.write(":HARDcopy:DESTination FILE")
            time.sleep(0.1)
            self._scpi_wrapper.write(f":HARDcopy:FORMat {image_format}")
            time.sleep(0.1)
            image_data = self._scpi_wrapper.query_binary_values(f":DISPlay:DATA? {image_format}", datatype='B')
            if image_data:
                with open(screenshot_path, 'wb') as f:
                    f.write(bytes(image_data))
                self._logger.info(f"Screenshot saved: {screenshot_path}")
                return str(screenshot_path)
            return None
        except Exception as e:
            self._logger.error(f"Screenshot capture failed: {e}")
            return None

    def setup_output_directories(self) -> None:
        """Create default output directories"""
        base_path = Path.cwd()
        self.screenshot_dir = base_path / "oscilloscope_screenshots"
        self.data_dir = base_path / "oscilloscope_data"
        self.graph_dir = base_path / "oscilloscope_graphs"
        for directory in [self.screenshot_dir, self.data_dir, self.graph_dir]:
            directory.mkdir(exist_ok=True)

    def configure_function_generator(self, generator: int, waveform: str = "SIN",
                                    frequency: float = 1000.0, amplitude: float = 1.0,
                                    offset: float = 0.0, enable: bool = True) -> bool:
        """Configure function generator output"""
        if not self.is_connected:
            self._logger.error("Cannot configure function generator: not connected")
            return False
        if generator not in [1, 2]:
            self._logger.error(f"Invalid generator number: {generator}")
            return False
        try:
            self._scpi_wrapper.write(f":WGEN{generator}:FUNCtion {waveform.upper()}")
            time.sleep(0.05)
            if waveform.upper() != "DC":
                self._scpi_wrapper.write(f":WGEN{generator}:FREQuency {frequency}")
                time.sleep(0.05)
            self._scpi_wrapper.write(f":WGEN{generator}:VOLTage {amplitude}")
            time.sleep(0.05)
            self._scpi_wrapper.write(f":WGEN{generator}:VOLTage:OFFSet {offset}")
            time.sleep(0.05)
            output_state = "ON" if enable else "OFF"
            self._scpi_wrapper.write(f":WGEN{generator}:OUTPut {output_state}")
            time.sleep(0.05)
            self._logger.info(f"WGEN{generator} configured: {waveform}, {frequency}Hz, {amplitude}Vpp")
            return True
        except Exception as e:
            self._logger.error(f"Failed to configure WGEN{generator}: {e}")
            return False

    def get_function_generator_config(self, generator: int) -> Optional[Dict[str, Any]]:
        """Query function generator configuration"""
        if not self.is_connected:
            return None
        if generator not in [1, 2]:
            return None
        try:
            self._scpi_wrapper.query("*OPC?")
            time.sleep(0.05)
            config = {
                'generator': generator,
                'function': self._scpi_wrapper.query(f":WGEN{generator}:FUNCtion?").strip(),
            }
            return config
        except Exception as e:
            self._logger.error(f"Failed to get WGEN{generator} config: {e}")
            return None
