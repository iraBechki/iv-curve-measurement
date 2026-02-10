""" 
High-Precision PV Panel I-V Curve Measurement Application
Optimized for: Arduino Uno + ACS712-30A + Voltage Divider + Manual Variable Resistor
Requires: PyQt5, pyserial, matplotlib, pandas, numpy
Install: pip install PyQt5 pyserial matplotlib pandas numpy openpyxl
"""

import sys
import json
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QGridLayout, QLabel, QLineEdit, QPushButton, QComboBox, 
                             QTextEdit, QTableWidget, QTableWidgetItem, QProgressBar,
                             QGroupBox, QMessageBox, QFileDialog, QHeaderView, QTabWidget)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont, QPalette, QColor
import serial
import serial.tools.list_ports
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import time
from datetime import datetime

class PVIVAnalyzer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PV I-V Curve Analyzer V2.0")
        self.setGeometry(100, 100, 1400, 900)
        
        self.serial_conn = None
        self.measurements = []
        self.voc = None
        self.isc = None
        self.measurement_count = 0
        self.current_voltage = None
        self.current_current = None
        self.current_power = None
        self.baud_rate = 9600
        
        self.plot_settings = {
            'show_iv': True,
            'show_pv': True,
            'show_voc_isc': True,
            'show_mpp': True,
            'iv_style': 'both',
            'pv_style': 'both',
            'iv_color': '#0000FF',
            'pv_color': '#FF0000',
            'line_width': 2,
            'marker_size': 6,
            'separate_graphs': False
        }
        
        self.config_file = "config.json"
        self.calibration = {
            "zero_offset": 0.0,
            "voltage_slope": 1.0,
            "voltage_offset": 0.0,
            "current_slope": 1.0,
            "current_offset": 0.0
        }
        self.load_config()
        
        self.results_text = QTextEdit()
        self.results_text.setReadOnly(True)
        
        self.setup_gui()
        self.update_plot()
        
    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    data = json.load(f)
                    if 'calibration' in data:
                        self.calibration.update(data['calibration'])
                    if 'plot_settings' in data:
                        self.plot_settings.update(data['plot_settings'])
                    if 'baud_rate' in data:
                        self.baud_rate = data['baud_rate']
            except Exception as e:
                print(f"Error loading config: {e}")
    
    def save_config(self):
        try:
            with open(self.config_file, 'w') as f:
                json.dump({
                    'calibration': self.calibration,
                    'plot_settings': self.plot_settings,
                    'baud_rate': self.baud_rate
                }, f, indent=2)
        except Exception as e:
            QMessageBox.warning(self, "Warning", f"Could not save config: {e}")
    
    def setup_gui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(5)
        
        main_layout.addWidget(self.create_connection_panel())
        
        self.tabs = QTabWidget()
        self.tabs.addTab(self.create_main_tab(), "Measurement")
        self.tabs.addTab(self.create_settings_tab(), "Settings")
        main_layout.addWidget(self.tabs)
        
        central_widget.setLayout(main_layout)
    
    def create_connection_panel(self):
        group = QGroupBox("Connection")
        layout = QHBoxLayout()
        
        port_label = QLabel("COM Port:")
        layout.addWidget(port_label)
        
        self.port_combo = QComboBox()
        self.port_combo.setMinimumWidth(150)
        self.port_combo.addItem("Select Port...")
        self.refresh_ports_list()
        self.port_combo.currentTextChanged.connect(self.on_port_selected)
        layout.addWidget(self.port_combo)
        
        self.connection_status = QLabel("Not Connected")
        self.connection_status.setStyleSheet("color: red; font-weight: bold;")
        layout.addWidget(self.connection_status)
        
        baud_label = QLabel("Baud Rate:")
        layout.addWidget(baud_label)
        
        self.baud_combo = QComboBox()
        self.baud_combo.addItems(["9600", "19200", "38400", "57600", "115200"])
        self.baud_combo.setCurrentText(str(self.baud_rate))
        self.baud_combo.currentTextChanged.connect(self.update_baud_rate)
        self.baud_combo.setMinimumWidth(100)
        layout.addWidget(self.baud_combo)
        
        layout.addStretch()
        
        group.setLayout(layout)
        return group
    
    def create_main_tab(self):
        widget = QWidget()
        main_layout = QHBoxLayout()
        
        left_column = QVBoxLayout()
        left_column.addWidget(self.create_measurement_panel())
        left_column.addWidget(self.create_data_table_panel())
        left_column.addWidget(self.create_parameters_panel())
        left_column.addWidget(self.create_save_panel())
        
        right_column = QVBoxLayout()
        right_column.addWidget(self.create_plot_panel())
        
        main_layout.addLayout(left_column, 1)
        main_layout.addLayout(right_column, 3)
        
        widget.setLayout(main_layout)
        return widget
    
    def create_settings_tab(self):
        widget = QWidget()
        main_layout = QHBoxLayout()
        
        cal_group = QGroupBox("Calibration")
        cal_group.setMaximumWidth(500)
        cal_layout = QVBoxLayout()
        
        title1 = QLabel("Step 1: Zero Current Calibration")
        title1.setFont(QFont("Arial", 10, QFont.Bold))
        cal_layout.addWidget(title1)
        
        cal_layout.addWidget(QLabel("Disconnect all loads, ensure no current flows through sensor"))
        
        zero_btn = QPushButton("Calibrate Zero Current")
        zero_btn.clicked.connect(self.calibrate_zero)
        cal_layout.addWidget(zero_btn)
        
        self.zero_status = QLabel(f"Current offset: {self.calibration['zero_offset']:.6f} A")
        self.zero_status.setStyleSheet("color: blue;")
        cal_layout.addWidget(self.zero_status)
        
        cal_layout.addWidget(self.create_separator())
        
        title2 = QLabel("Step 2: Voltage Calibration (Two-Point)")
        title2.setFont(QFont("Arial", 10, QFont.Bold))
        cal_layout.addWidget(title2)
        
        v_layout = QGridLayout()
        
        v_layout.addWidget(QLabel("Point 1 - Low Voltage:"), 0, 0, 1, 3)
        
        measure_v1_btn = QPushButton("Measure Point 1")
        measure_v1_btn.clicked.connect(lambda: self.measure_voltage_for_calibration(1))
        v_layout.addWidget(measure_v1_btn, 1, 0)
        
        self.cal_v1_measured_label = QLabel("Measured: -- V")
        v_layout.addWidget(self.cal_v1_measured_label, 1, 1)
        
        self.cal_v1_reference = QLineEdit()
        self.cal_v1_reference.setPlaceholderText("Actual (V)")
        v_layout.addWidget(self.cal_v1_reference, 1, 2)
        
        v_layout.addWidget(QLabel("Point 2 - High Voltage:"), 2, 0, 1, 3)
        
        measure_v2_btn = QPushButton("Measure Point 2")
        measure_v2_btn.clicked.connect(lambda: self.measure_voltage_for_calibration(2))
        v_layout.addWidget(measure_v2_btn, 3, 0)
        
        self.cal_v2_measured_label = QLabel("Measured: -- V")
        v_layout.addWidget(self.cal_v2_measured_label, 3, 1)
        
        self.cal_v2_reference = QLineEdit()
        self.cal_v2_reference.setPlaceholderText("Actual (V)")
        v_layout.addWidget(self.cal_v2_reference, 3, 2)
        
        cal_v_btn = QPushButton("Apply Voltage Calibration")
        cal_v_btn.clicked.connect(self.calculate_voltage_calibration)
        v_layout.addWidget(cal_v_btn, 4, 0, 1, 3)
        
        cal_layout.addLayout(v_layout)
        
        self.v_cal_status = QLabel(f"Slope: {self.calibration['voltage_slope']:.6f}, Offset: {self.calibration['voltage_offset']:.6f}")
        self.v_cal_status.setStyleSheet("color: blue;")
        cal_layout.addWidget(self.v_cal_status)
        
        cal_layout.addWidget(self.create_separator())
        
        title3 = QLabel("Step 3: Current Calibration (Two-Point)")
        title3.setFont(QFont("Arial", 10, QFont.Bold))
        cal_layout.addWidget(title3)
        
        i_layout = QGridLayout()
        
        i_layout.addWidget(QLabel("Point 1 - Low Current:"), 0, 0, 1, 3)
        
        measure_i1_btn = QPushButton("Measure Point 1")
        measure_i1_btn.clicked.connect(lambda: self.measure_current_for_calibration(1))
        i_layout.addWidget(measure_i1_btn, 1, 0)
        
        self.cal_i1_measured_label = QLabel("Measured: -- A")
        i_layout.addWidget(self.cal_i1_measured_label, 1, 1)
        
        self.cal_i1_reference = QLineEdit()
        self.cal_i1_reference.setPlaceholderText("Actual (A)")
        i_layout.addWidget(self.cal_i1_reference, 1, 2)
        
        i_layout.addWidget(QLabel("Point 2 - High Current:"), 2, 0, 1, 3)
        
        measure_i2_btn = QPushButton("Measure Point 2")
        measure_i2_btn.clicked.connect(lambda: self.measure_current_for_calibration(2))
        i_layout.addWidget(measure_i2_btn, 3, 0)
        
        self.cal_i2_measured_label = QLabel("Measured: -- A")
        i_layout.addWidget(self.cal_i2_measured_label, 3, 1)
        
        self.cal_i2_reference = QLineEdit()
        self.cal_i2_reference.setPlaceholderText("Actual (A)")
        i_layout.addWidget(self.cal_i2_reference, 3, 2)
        
        cal_i_btn = QPushButton("Apply Current Calibration")
        cal_i_btn.clicked.connect(self.calculate_current_calibration)
        i_layout.addWidget(cal_i_btn, 4, 0, 1, 3)
        
        cal_layout.addLayout(i_layout)
        
        self.i_cal_status = QLabel(f"Slope: {self.calibration['current_slope']:.6f}, Offset: {self.calibration['current_offset']:.6f}")
        self.i_cal_status.setStyleSheet("color: blue;")
        cal_layout.addWidget(self.i_cal_status)
        
        cal_layout.addStretch()
        
        cal_group.setLayout(cal_layout)
        main_layout.addWidget(cal_group)
        
        main_layout.addStretch()
        
        widget.setLayout(main_layout)
        return widget
    
    def create_measurement_panel(self):
        group = QGroupBox("Measurements")
        layout = QVBoxLayout()
        layout.setSpacing(6)
        
        self.param_label = QLabel("Voc: -- | Isc: --")
        self.param_label.setFont(QFont("Courier", 9))
        self.param_label.setStyleSheet("padding: 3px;")
        self.param_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.param_label)
        
        btn_row1 = QHBoxLayout()
        voc_btn = QPushButton("Measure Voc")
        voc_btn.setToolTip("Open circuit voltage - disconnect all loads")
        voc_btn.clicked.connect(self.measure_voc)
        btn_row1.addWidget(voc_btn)
        
        isc_btn = QPushButton("Measure Isc")
        isc_btn.setToolTip("Short circuit current - short the terminals with thick wire")
        isc_btn.clicked.connect(self.measure_isc)
        btn_row1.addWidget(isc_btn)
        layout.addLayout(btn_row1)
        
        measure_btn = QPushButton("Measure Point")
        measure_btn.clicked.connect(self.measure_single_point)
        layout.addWidget(measure_btn)
        
        self.current_measurement_label = QLabel("V: -- | I: -- | P: --")
        self.current_measurement_label.setFont(QFont("Courier", 9))
        self.current_measurement_label.setStyleSheet("padding: 3px;")
        self.current_measurement_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.current_measurement_label)
        
        btn_row2 = QHBoxLayout()
        add_btn = QPushButton("Add to Table")
        add_btn.clicked.connect(self.add_measurement_to_table)
        btn_row2.addWidget(add_btn)
        
        clear_btn = QPushButton("Clear All")
        clear_btn.clicked.connect(self.clear_data)
        btn_row2.addWidget(clear_btn)
        layout.addLayout(btn_row2)
        
        group.setLayout(layout)
        return group
    
    def create_plot_panel(self):
        group = QGroupBox("I-V and P-V Curves")
        layout = QVBoxLayout()
        
        settings_btn = QPushButton("⚙ Plot Settings")
        settings_btn.setStyleSheet("font-weight: bold; padding: 5px;")
        settings_btn.clicked.connect(self.show_plot_settings)
        layout.addWidget(settings_btn)
        
        self.figure = Figure(figsize=(12, 7.5), dpi=100)
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)
        
        group.setLayout(layout)
        return group
    
    def create_data_table_panel(self):
        group = QGroupBox("Measurement Data")
        layout = QVBoxLayout()
        layout.setSpacing(3)
        
        self.data_table = QTableWidget()
        self.data_table.setColumnCount(3)
        self.data_table.setHorizontalHeaderLabels(['V (V)', 'I (A)', 'P (W)'])
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.data_table.setMinimumHeight(120)
        self.data_table.setMaximumHeight(200)
        layout.addWidget(self.data_table)
        
        group.setLayout(layout)
        return group
    
    def create_parameters_panel(self):
        group = QGroupBox("Parameters")
        layout = QVBoxLayout()
        layout.setSpacing(6)
        
        self.params_display = QLabel("Vmpp: -- | Impp: -- | Pmpp: -- | FF: --")
        self.params_display.setFont(QFont("Courier", 9))
        self.params_display.setStyleSheet("padding: 3px;")
        self.params_display.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.params_display)
        
        self.params_display2 = QLabel("Rs: -- | Rsh: --")
        self.params_display2.setFont(QFont("Courier", 9))
        self.params_display2.setStyleSheet("padding: 3px;")
        self.params_display2.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.params_display2)
        
        calc_btn = QPushButton("Calculate")
        calc_btn.clicked.connect(self.calculate_parameters)
        layout.addWidget(calc_btn)
        
        group.setLayout(layout)
        return group
    
    def create_save_panel(self):
        group = QGroupBox("Export")
        layout = QHBoxLayout()
        
        export_btn = QPushButton("Export Data/Plot")
        export_btn.clicked.connect(self.export_dialog)
        layout.addWidget(export_btn)
        
        layout.addStretch()
        
        group.setLayout(layout)
        return group
    
    def create_separator(self):
        line = QWidget()
        line.setFixedHeight(2)
        line.setStyleSheet("background-color: #c0c0c0; margin: 10px 0;")
        return line
    
    def refresh_ports_list(self):
        current_text = self.port_combo.currentText()
        self.port_combo.blockSignals(True)
        self.port_combo.clear()
        self.port_combo.addItem("Select Port...")
        
        ports = [port.device for port in serial.tools.list_ports.comports()]
        if ports:
            self.port_combo.addItems(ports)
            if current_text in ports:
                self.port_combo.setCurrentText(current_text)
        
        self.port_combo.blockSignals(False)
    
    def on_port_selected(self, port):
        if port and port != "Select Port...":
            if self.serial_conn and self.serial_conn.is_open:
                self.disconnect_device()
            self.connect_device(port)
    
    def connect_device(self, port):
        try:
            self.serial_conn = serial.Serial(port, self.baud_rate, timeout=3)
            time.sleep(2.5)
            
            self.serial_conn.reset_input_buffer()
            
            self.serial_conn.write(b"PING\n")
            time.sleep(0.2)
            response = self.serial_conn.readline().decode().strip()
            
            if "PONG" in response or "READY" in response or "OK" in response or response:
                self.connection_status.setText(f"Connected: {port} @ {self.baud_rate} baud")
                self.connection_status.setStyleSheet("color: green; font-weight: bold;")
                if not ("PONG" in response or "READY" in response or "OK" in response):
                    QMessageBox.warning(self, "Connected", 
                                      f"Connected to {port}\n"
                                      f"Baud rate: {self.baud_rate}\n\n"
                                      f"Note: Device responded with '{response}'\n"
                                      f"Expected 'PONG', 'READY', or 'OK'\n\n"
                                      f"Connection may work, but verify baud rate matches Arduino.")
            else:
                raise Exception(f"No response from device")
                
        except Exception as e:
            QMessageBox.critical(self, "Connection Error", 
                               f"Failed to connect:\n{str(e)}\n\n"
                               f"Current baud rate: {self.baud_rate}\n\n"
                               f"Try:\n- Change baud rate to match Arduino (likely 115200)\n"
                               f"- Select correct COM port\n- Check USB cable\n- Restart Arduino")
            self.disconnect_device()
    
    def disconnect_device(self):
        if self.serial_conn and self.serial_conn.is_open:
            self.serial_conn.close()
        self.serial_conn = None
        self.connection_status.setText("Not Connected")
        self.connection_status.setStyleSheet("color: red; font-weight: bold;")
        self.port_combo.blockSignals(True)
        self.port_combo.setCurrentIndex(0)
        self.port_combo.blockSignals(False)
    
    def update_baud_rate(self, value):
        self.baud_rate = int(value)
        self.save_config()
        if self.serial_conn and self.serial_conn.is_open:
            QMessageBox.information(self, "Baud Rate Changed", 
                                  f"Baud rate set to {self.baud_rate}\n\n"
                                  "Please reconnect for changes to take effect.")
    
    def send_command(self, command, timeout=5):
        if not self.serial_conn or not self.serial_conn.is_open:
            QMessageBox.critical(self, "Error", "Device not connected")
            return None
        
        try:
            self.serial_conn.reset_input_buffer()
            self.serial_conn.write(f"{command}\n".encode())
            response = self.serial_conn.readline().decode().strip()
            return response
        except Exception as e:
            QMessageBox.critical(self, "Communication Error", f"Failed to communicate:\n{str(e)}")
            return None
    
    def calibrate_zero(self):
        reply = QMessageBox.question(self, "Zero Calibration",
                                    "Make sure:\n"
                                    "1. NO load is connected\n"
                                    "2. Current sensor has NO current flowing\n"
                                    "3. Panel can be disconnected or in dark\n\n"
                                    "Continue?",
                                    QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            response = self.send_command("CALIBRATE_ZERO")
            if response and "ZERO_CALIBRATED" in response:
                try:
                    offset = float(response.split(":")[1])
                    self.calibration['zero_offset'] = offset
                    self.save_config()
                    self.zero_status.setText(f"Current offset: {offset:.6f} A ✓")
                    self.zero_status.setStyleSheet("color: green; font-weight: bold;")
                    QMessageBox.information(self, "Success", f"Zero current calibrated\nOffset: {offset:.6f} A")
                except:
                    QMessageBox.warning(self, "Warning", "Calibration completed but could not parse value")
            else:
                self.zero_status.setText("Calibration failed!")
                self.zero_status.setStyleSheet("color: red;")
    
    def measure_voltage_for_calibration(self, point):
        response = self.send_command("VOC")
        if response and response.startswith("VOC:"):
            voltage = float(response.split(":")[1])
            if point == 1:
                self.cal_v1_measured_label.setText(f"Measured: {voltage:.5f} V")
                self.cal_v1_measured_value = voltage
            else:
                self.cal_v2_measured_label.setText(f"Measured: {voltage:.5f} V")
                self.cal_v2_measured_value = voltage
        else:
            QMessageBox.critical(self, "Error", "Failed to measure voltage")
    
    def calculate_voltage_calibration(self):
        if not hasattr(self, 'cal_v1_measured_value') or not hasattr(self, 'cal_v2_measured_value'):
            QMessageBox.critical(self, "Error", "Please measure both voltage points first")
            return
        
        try:
            ref1 = float(self.cal_v1_reference.text())
            ref2 = float(self.cal_v2_reference.text())
            meas1 = self.cal_v1_measured_value
            meas2 = self.cal_v2_measured_value
            
            if abs(meas2 - meas1) < 0.001:
                QMessageBox.critical(self, "Error", "Measured points are too close. Use different voltage levels.")
                return
            
            slope = (ref2 - ref1) / (meas2 - meas1)
            offset = ref1 - slope * meas1
            
            response = self.send_command(f"CAL_V_LINEAR:{slope:.6f},{offset:.6f}")
            if response and "VLINEAR" in response:
                self.calibration['voltage_slope'] = slope
                self.calibration['voltage_offset'] = offset
                self.save_config()
                self.v_cal_status.setText(f"Slope: {slope:.6f}, Offset: {offset:.6f} ✓")
                self.v_cal_status.setStyleSheet("color: green; font-weight: bold;")
                QMessageBox.information(self, "Success", 
                                      f"Voltage calibration applied\n"
                                      f"y = {slope:.6f}*x + {offset:.6f}\n\n"
                                      f"Point 1: {meas1:.3f}V → {ref1:.3f}V\n"
                                      f"Point 2: {meas2:.3f}V → {ref2:.3f}V")
        except ValueError:
            QMessageBox.critical(self, "Error", "Invalid reference voltage values")
    
    def measure_current_for_calibration(self, point):
        response = self.send_command("MEASURE")
        if response and response.startswith("DATA:"):
            values = response.split(":")[1].split(",")
            current = float(values[1])
            if point == 1:
                self.cal_i1_measured_label.setText(f"Measured: {current:.5f} A")
                self.cal_i1_measured_value = current
            else:
                self.cal_i2_measured_label.setText(f"Measured: {current:.5f} A")
                self.cal_i2_measured_value = current
        else:
            QMessageBox.critical(self, "Error", "Failed to measure current")
    
    def calculate_current_calibration(self):
        if not hasattr(self, 'cal_i1_measured_value') or not hasattr(self, 'cal_i2_measured_value'):
            QMessageBox.critical(self, "Error", "Please measure both current points first")
            return
        
        try:
            ref1 = float(self.cal_i1_reference.text())
            ref2 = float(self.cal_i2_reference.text())
            meas1 = self.cal_i1_measured_value
            meas2 = self.cal_i2_measured_value
            
            if abs(meas2 - meas1) < 0.001:
                QMessageBox.critical(self, "Error", "Measured points are too close. Use different current levels.")
                return
            
            slope = (ref2 - ref1) / (meas2 - meas1)
            offset = ref1 - slope * meas1
            
            response = self.send_command(f"CAL_I_LINEAR:{slope:.6f},{offset:.6f}")
            if response and "ILINEAR" in response:
                self.calibration['current_slope'] = slope
                self.calibration['current_offset'] = offset
                self.save_config()
                self.i_cal_status.setText(f"Slope: {slope:.6f}, Offset: {offset:.6f} ✓")
                self.i_cal_status.setStyleSheet("color: green; font-weight: bold;")
                QMessageBox.information(self, "Success", 
                                      f"Current calibration applied\n"
                                      f"y = {slope:.6f}*x + {offset:.6f}\n\n"
                                      f"Point 1: {meas1:.3f}A → {ref1:.3f}A\n"
                                      f"Point 2: {meas2:.3f}A → {ref2:.3f}A")
        except ValueError:
            QMessageBox.critical(self, "Error", "Invalid reference current values")
    
    def measure_voc(self):
        response = self.send_command("VOC")
        if response and response.startswith("VOC:"):
            self.voc = float(response.split(":")[1])
            isc_text = f"{self.isc:.4f}" if self.isc is not None else "--"
            self.param_label.setText(f"Voc: {self.voc:.4f} | Isc: {isc_text}")
            
            self.measurement_count += 1
            self.measurements.append({
                'V': self.voc, 
                'I': 0.0, 
                'P': 0.0
            })
            
            row = self.data_table.rowCount()
            self.data_table.insertRow(row)
            self.data_table.setItem(row, 0, QTableWidgetItem(f"{self.voc:.4f}"))
            self.data_table.setItem(row, 1, QTableWidgetItem("0.0000"))
            self.data_table.setItem(row, 2, QTableWidgetItem("0.0000"))
            
            self.update_results()
            self.update_plot()
        else:
            QMessageBox.critical(self, "Error", "Voc measurement failed")
    
    def measure_isc(self):
        response = self.send_command("ISC")
        if response and response.startswith("ISC:"):
            self.isc = float(response.split(":")[1])
            voc_text = f"{self.voc:.4f}" if self.voc is not None else "--"
            self.param_label.setText(f"Voc: {voc_text} | Isc: {self.isc:.4f}")
            
            self.measurement_count += 1
            self.measurements.append({
                'V': 0.0, 
                'I': self.isc, 
                'P': 0.0
            })
            
            row = self.data_table.rowCount()
            self.data_table.insertRow(row)
            self.data_table.setItem(row, 0, QTableWidgetItem("0.0000"))
            self.data_table.setItem(row, 1, QTableWidgetItem(f"{self.isc:.4f}"))
            self.data_table.setItem(row, 2, QTableWidgetItem("0.0000"))
            
            self.update_results()
            self.update_plot()
        else:
            QMessageBox.critical(self, "Error", "Isc measurement failed")
    
    def measure_single_point(self):
        if not self.serial_conn or not self.serial_conn.is_open:
            QMessageBox.critical(self, "Error", "Device not connected")
            return
        
        response = self.send_command("MEASURE")
        
        if response and response.startswith("DATA:"):
            try:
                values = response.split(":")[1].split(",")
                self.current_voltage = float(values[0])
                self.current_current = float(values[1])
                self.current_power = self.current_voltage * self.current_current
                
                self.current_measurement_label.setText(
                    f"V: {self.current_voltage:.4f} | I: {self.current_current:.4f} | P: {self.current_power:.4f}"
                )
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to parse measurement:\n{str(e)}")
        else:
            QMessageBox.critical(self, "Error", "Measurement failed - check connection")
    
    def add_measurement_to_table(self):
        if self.current_voltage is None or self.current_current is None:
            QMessageBox.warning(self, "Warning", "Please measure a point first before adding to table")
            return
        
        try:
            if self.current_power is None:
                self.current_power = self.current_voltage * self.current_current
            
            self.measurement_count += 1
            self.measurements.append({
                'V': self.current_voltage, 
                'I': self.current_current, 
                'P': self.current_power
            })
            
            row = self.data_table.rowCount()
            self.data_table.insertRow(row)
            self.data_table.setItem(row, 0, QTableWidgetItem(f"{self.current_voltage:.4f}"))
            self.data_table.setItem(row, 1, QTableWidgetItem(f"{self.current_current:.4f}"))
            self.data_table.setItem(row, 2, QTableWidgetItem(f"{self.current_power:.4f}"))
            
            self.current_measurement_label.setText("V: -- | I: -- | P: --")
            
            self.current_voltage = None
            self.current_current = None
            self.current_power = None
            
            self.update_plot()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to add measurement:\n{str(e)}")
    
    def clear_data(self):
        reply = QMessageBox.question(self, "Clear Data", "Clear all measurement points?",
                                    QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            self.measurements.clear()
            self.measurement_count = 0
            self.data_table.setRowCount(0)
            self.update_plot()
            self.update_results()
            self.current_measurement_label.setText("V: -- | I: -- | P: --")
            self.current_voltage = None
            self.current_current = None
            self.current_power = None
    
    def calculate_parameters(self):
        if not self.measurements:
            QMessageBox.warning(self, "Warning", "No measurement data available")
            self.params_display.setText("Vmpp: -- | Impp: -- | Pmpp: -- | FF: --")
            self.params_display2.setText("Rs: -- | Rsh: --")
            return
        
        df = pd.DataFrame(self.measurements)
        
        mpp_idx = df['P'].idxmax()
        vmpp = df.loc[mpp_idx, 'V']
        impp = df.loc[mpp_idx, 'I']
        pmpp = df.loc[mpp_idx, 'P']
        
        ff = None
        if self.voc and self.isc and self.isc > 0:
            ff = pmpp / (self.voc * self.isc)
        
        rs = None
        if len(df) > 5:
            voc_region = df.nlargest(3, 'V')
            if len(voc_region) >= 2:
                dv = voc_region['V'].diff().abs().mean()
                di = voc_region['I'].diff().abs().mean()
                if di > 0:
                    rs = dv / di
        
        rsh = None
        if len(df) > 5:
            isc_region = df.nsmallest(3, 'V')
            if len(isc_region) >= 2:
                dv = isc_region['V'].diff().abs().mean()
                di = isc_region['I'].diff().abs().mean()
                if dv > 0:
                    rsh = di / dv
        
        ff_text = f"{ff:.4f}" if ff is not None else "--"
        params_text = f"Vmpp: {vmpp:.4f} | Impp: {impp:.4f} | Pmpp: {pmpp:.4f} | FF: {ff_text}"
        self.params_display.setText(params_text)
        
        rs_text = f"{rs:.4f}" if rs is not None else "--"
        rsh_text = f"{rsh:.4f}" if rsh is not None else "--"
        params_text2 = f"Rs: {rs_text} Ω | Rsh: {rsh_text} Ω"
        self.params_display2.setText(params_text2)
        
        self.update_results(vmpp, impp, pmpp, ff, rs, rsh)
    
    def update_results(self, vmpp=None, impp=None, pmpp=None, ff=None, rs=None, rsh=None):
        text = "╔══════════════════════════════════════╗\n"
        text += "║   PV PANEL CHARACTERISTICS           ║\n"
        text += "╚══════════════════════════════════════╝\n\n"
        
        text += "─── Basic Parameters ───────────────────\n\n"
        
        if self.voc is not None:
            text += f"  Open Circuit Voltage (Voc):\n"
            text += f"    {self.voc:.5f} V\n\n"
        
        if self.isc is not None:
            text += f"  Short Circuit Current (Isc):\n"
            text += f"    {self.isc:.5f} A\n\n"
        
        if vmpp is not None:
            text += "─── Maximum Power Point ────────────────\n\n"
            text += f"  Voltage at MPP (Vmpp):\n"
            text += f"    {vmpp:.5f} V\n\n"
            text += f"  Current at MPP (Impp):\n"
            text += f"    {impp:.5f} A\n\n"
            text += f"  Maximum Power (Pmpp):\n"
            text += f"    {pmpp:.5f} W\n\n"
        
        if ff is not None:
            text += "─── Performance Metrics ────────────────\n\n"
            text += f"  Fill Factor (FF):\n"
            text += f"    {ff:.4f} ({ff*100:.2f}%)\n\n"
        
        if rs is not None:
            text += f"  Series Resistance (Rs):\n"
            text += f"    {rs:.4f} Ω\n\n"
        
        if rsh is not None:
            text += f"  Shunt Resistance (Rsh):\n"
            text += f"    {rsh:.4f} Ω\n\n"
        
        if self.measurements:
            text += "─── Measurement Info ───────────────────\n\n"
            text += f"  Total Data Points: {len(self.measurements)}\n\n"
        
        text += f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        
        self.results_text.setText(text)
    
    def show_plot_settings(self):
        from PyQt5.QtWidgets import (QDialog, QCheckBox, QRadioButton, QButtonGroup, 
                                     QSpinBox, QColorDialog, QTabWidget)
        
        dialog = QDialog(self)
        dialog.setWindowTitle("⚙ Plot Customization")
        dialog.setMinimumWidth(500)
        dialog.setMinimumHeight(600)
        main_layout = QVBoxLayout()
        
        tabs = QTabWidget()
        
        general_tab = QWidget()
        general_layout = QVBoxLayout()
        
        layout_group = QGroupBox("📊 Graph Layout")
        layout_layout = QVBoxLayout()
        layout_button_group = QButtonGroup(dialog)
        
        same_graph_radio = QRadioButton("Same Graph (I-V and P-V on dual axis)")
        same_graph_radio.setChecked(not self.plot_settings['separate_graphs'])
        layout_button_group.addButton(same_graph_radio, 1)
        layout_layout.addWidget(same_graph_radio)
        
        separate_graphs_radio = QRadioButton("Separate Graphs (I-V above, P-V below)")
        separate_graphs_radio.setChecked(self.plot_settings['separate_graphs'])
        layout_button_group.addButton(separate_graphs_radio, 2)
        layout_layout.addWidget(separate_graphs_radio)
        
        layout_group.setLayout(layout_layout)
        general_layout.addWidget(layout_group)
        
        curves_group = QGroupBox("📈 Display Elements")
        curves_layout = QVBoxLayout()
        
        iv_check = QCheckBox("✓ I-V Curve")
        iv_check.setChecked(self.plot_settings['show_iv'])
        iv_check.setStyleSheet("font-weight: bold;")
        curves_layout.addWidget(iv_check)
        
        pv_check = QCheckBox("✓ P-V Curve")
        pv_check.setChecked(self.plot_settings['show_pv'])
        pv_check.setStyleSheet("font-weight: bold;")
        curves_layout.addWidget(pv_check)
        
        voc_isc_check = QCheckBox(" Voc and Isc Reference Points")
        voc_isc_check.setChecked(self.plot_settings['show_voc_isc'])
        curves_layout.addWidget(voc_isc_check)
        
        mpp_check = QCheckBox(" Maximum Power Point (MPP)")
        mpp_check.setChecked(self.plot_settings['show_mpp'])
        curves_layout.addWidget(mpp_check)
        
        curves_group.setLayout(curves_layout)
        general_layout.addWidget(curves_group)
        
        general_tab.setLayout(general_layout)
        tabs.addTab(general_tab, "General")
        
        style_tab = QWidget()
        style_layout = QVBoxLayout()
        
        iv_style_group = QGroupBox("🔵 I-V Curve Style")
        iv_style_layout = QGridLayout()
        iv_button_group = QButtonGroup(dialog)
        
        iv_style_layout.addWidget(QLabel("Style:"), 0, 0)
        
        iv_line = QRadioButton("─ Line")
        iv_button_group.addButton(iv_line, 1)
        iv_style_layout.addWidget(iv_line, 0, 1)
        
        iv_points = QRadioButton("● Points")
        iv_button_group.addButton(iv_points, 2)
        iv_style_layout.addWidget(iv_points, 0, 2)
        
        iv_both = QRadioButton("●─ Both")
        iv_both.setChecked(self.plot_settings['iv_style'] == 'both')
        iv_button_group.addButton(iv_both, 3)
        iv_style_layout.addWidget(iv_both, 0, 3)
        
        if self.plot_settings['iv_style'] == 'line':
            iv_line.setChecked(True)
        elif self.plot_settings['iv_style'] == 'points':
            iv_points.setChecked(True)
        
        iv_style_layout.addWidget(QLabel("Color:"), 1, 0)
        iv_color_btn = QPushButton("Choose Color")
        iv_color_btn.setStyleSheet(f"background-color: {self.plot_settings['iv_color']}; color: white; font-weight: bold; padding: 5px;")
        iv_color_btn.clicked.connect(lambda: self.choose_color(iv_color_btn, 'iv_color'))
        iv_style_layout.addWidget(iv_color_btn, 1, 1, 1, 3)
        
        iv_style_group.setLayout(iv_style_layout)
        style_layout.addWidget(iv_style_group)
        
        pv_style_group = QGroupBox("🔴 P-V Curve Style")
        pv_style_layout = QGridLayout()
        pv_button_group = QButtonGroup(dialog)
        
        pv_style_layout.addWidget(QLabel("Style:"), 0, 0)
        
        pv_line = QRadioButton("─ Line")
        pv_button_group.addButton(pv_line, 1)
        pv_style_layout.addWidget(pv_line, 0, 1)
        
        pv_points = QRadioButton("■ Points")
        pv_button_group.addButton(pv_points, 2)
        pv_style_layout.addWidget(pv_points, 0, 2)
        
        pv_both = QRadioButton("■─ Both")
        pv_both.setChecked(self.plot_settings['pv_style'] == 'both')
        pv_button_group.addButton(pv_both, 3)
        pv_style_layout.addWidget(pv_both, 0, 3)
        
        if self.plot_settings['pv_style'] == 'line':
            pv_line.setChecked(True)
        elif self.plot_settings['pv_style'] == 'points':
            pv_points.setChecked(True)
        
        pv_style_layout.addWidget(QLabel("Color:"), 1, 0)
        pv_color_btn = QPushButton("Choose Color")
        pv_color_btn.setStyleSheet(f"background-color: {self.plot_settings['pv_color']}; color: white; font-weight: bold; padding: 5px;")
        pv_color_btn.clicked.connect(lambda: self.choose_color(pv_color_btn, 'pv_color'))
        pv_style_layout.addWidget(pv_color_btn, 1, 1, 1, 3)
        
        pv_style_group.setLayout(pv_style_layout)
        style_layout.addWidget(pv_style_group)
        
        size_group = QGroupBox("📏 Line & Marker Sizes")
        size_layout = QGridLayout()
        
        size_layout.addWidget(QLabel("Line Width:"), 0, 0)
        line_width_spin = QSpinBox()
        line_width_spin.setRange(1, 10)
        line_width_spin.setValue(self.plot_settings['line_width'])
        line_width_spin.setSuffix(" px")
        size_layout.addWidget(line_width_spin, 0, 1)
        
        size_layout.addWidget(QLabel("Marker Size:"), 1, 0)
        marker_size_spin = QSpinBox()
        marker_size_spin.setRange(2, 20)
        marker_size_spin.setValue(self.plot_settings['marker_size'])
        marker_size_spin.setSuffix(" pt")
        size_layout.addWidget(marker_size_spin, 1, 1)
        
        size_group.setLayout(size_layout)
        style_layout.addWidget(size_group)
        
        style_tab.setLayout(style_layout)
        tabs.addTab(style_tab, "Styles & Colors")
        
        axes_tab = QWidget()
        axes_layout = QVBoxLayout()
        
        axes_group = QGroupBox("📐 Axes Scale")
        axes_grid = QGridLayout()
        
        axes_grid.addWidget(QLabel("Voltage Axis:"), 0, 0, 1, 2)
        axes_grid.addWidget(QLabel("Min:"), 1, 0)
        v_min_spin = QLineEdit()
        v_min_spin.setPlaceholderText("Auto")
        if 'v_min' in self.plot_settings and self.plot_settings['v_min'] is not None:
            v_min_spin.setText(str(self.plot_settings['v_min']))
        axes_grid.addWidget(v_min_spin, 1, 1)
        
        axes_grid.addWidget(QLabel("Max:"), 2, 0)
        v_max_spin = QLineEdit()
        v_max_spin.setPlaceholderText("Auto")
        if 'v_max' in self.plot_settings and self.plot_settings['v_max'] is not None:
            v_max_spin.setText(str(self.plot_settings['v_max']))
        axes_grid.addWidget(v_max_spin, 2, 1)
        
        axes_grid.addWidget(QLabel("Current Axis:"), 3, 0, 1, 2)
        axes_grid.addWidget(QLabel("Min:"), 4, 0)
        i_min_spin = QLineEdit()
        i_min_spin.setPlaceholderText("Auto")
        if 'i_min' in self.plot_settings and self.plot_settings['i_min'] is not None:
            i_min_spin.setText(str(self.plot_settings['i_min']))
        axes_grid.addWidget(i_min_spin, 4, 1)
        
        axes_grid.addWidget(QLabel("Max:"), 5, 0)
        i_max_spin = QLineEdit()
        i_max_spin.setPlaceholderText("Auto")
        if 'i_max' in self.plot_settings and self.plot_settings['i_max'] is not None:
            i_max_spin.setText(str(self.plot_settings['i_max']))
        axes_grid.addWidget(i_max_spin, 5, 1)
        
        axes_grid.addWidget(QLabel("Power Axis:"), 6, 0, 1, 2)
        axes_grid.addWidget(QLabel("Min:"), 7, 0)
        p_min_spin = QLineEdit()
        p_min_spin.setPlaceholderText("Auto")
        if 'p_min' in self.plot_settings and self.plot_settings['p_min'] is not None:
            p_min_spin.setText(str(self.plot_settings['p_min']))
        axes_grid.addWidget(p_min_spin, 7, 1)
        
        axes_grid.addWidget(QLabel("Max:"), 8, 0)
        p_max_spin = QLineEdit()
        p_max_spin.setPlaceholderText("Auto")
        if 'p_max' in self.plot_settings and self.plot_settings['p_max'] is not None:
            p_max_spin.setText(str(self.plot_settings['p_max']))
        axes_grid.addWidget(p_max_spin, 8, 1)
        
        reset_axes_btn = QPushButton("Reset to Auto")
        reset_axes_btn.clicked.connect(lambda: (
            v_min_spin.clear(), v_max_spin.clear(),
            i_min_spin.clear(), i_max_spin.clear(),
            p_min_spin.clear(), p_max_spin.clear()
        ))
        axes_grid.addWidget(reset_axes_btn, 9, 0, 1, 2)
        
        axes_group.setLayout(axes_grid)
        axes_layout.addWidget(axes_group)
        axes_layout.addStretch()
        
        axes_tab.setLayout(axes_layout)
        tabs.addTab(axes_tab, "Axes")
        
        main_layout.addWidget(tabs)
        
        btn_layout = QHBoxLayout()
        apply_btn = QPushButton("✓ Apply Changes")
        apply_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 8px;")
        apply_btn.clicked.connect(lambda: self.apply_plot_settings(
            iv_check.isChecked(), pv_check.isChecked(), 
            voc_isc_check.isChecked(), mpp_check.isChecked(),
            layout_button_group.checkedId() == 2,
            ['line', 'points', 'both'][iv_button_group.checkedId() - 1],
            ['line', 'points', 'both'][pv_button_group.checkedId() - 1],
            line_width_spin.value(), marker_size_spin.value(),
            v_min_spin.text(), v_max_spin.text(),
            i_min_spin.text(), i_max_spin.text(),
            p_min_spin.text(), p_max_spin.text(),
            dialog
        ))
        btn_layout.addWidget(apply_btn)
        
        cancel_btn = QPushButton("✗ Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        btn_layout.addWidget(cancel_btn)
        
        main_layout.addLayout(btn_layout)
        dialog.setLayout(main_layout)
        dialog.exec_()
    
    def choose_color(self, button, setting_key):
        from PyQt5.QtWidgets import QColorDialog
        color = QColorDialog.getColor()
        if color.isValid():
            self.plot_settings[setting_key] = color.name()
            button.setStyleSheet(f"background-color: {color.name()}; color: white; font-weight: bold; padding: 5px;")
    
    def apply_plot_settings(self, show_iv, show_pv, show_voc_isc, show_mpp, separate_graphs,
                           iv_style, pv_style, line_width, marker_size, 
                           v_min, v_max, i_min, i_max, p_min, p_max, dialog):
        self.plot_settings['show_iv'] = show_iv
        self.plot_settings['show_pv'] = show_pv
        self.plot_settings['show_voc_isc'] = show_voc_isc
        self.plot_settings['show_mpp'] = show_mpp
        self.plot_settings['separate_graphs'] = separate_graphs
        self.plot_settings['iv_style'] = iv_style
        self.plot_settings['pv_style'] = pv_style
        self.plot_settings['line_width'] = line_width
        self.plot_settings['marker_size'] = marker_size
        
        self.plot_settings['v_min'] = float(v_min) if v_min else None
        self.plot_settings['v_max'] = float(v_max) if v_max else None
        self.plot_settings['i_min'] = float(i_min) if i_min else None
        self.plot_settings['i_max'] = float(i_max) if i_max else None
        self.plot_settings['p_min'] = float(p_min) if p_min else None
        self.plot_settings['p_max'] = float(p_max) if p_max else None
        
        self.save_config()
        self.update_plot()
        dialog.accept()
    
    def update_plot(self):
        self.figure.clear()
        
        if not self.measurements:
            ax = self.figure.add_subplot(111)
            ax.text(0.5, 0.5, 'No Data', transform=ax.transAxes,
                   ha='center', va='center', fontsize=14, color='gray')
            ax.set_xlabel('Voltage (V)', fontsize=12, fontweight='bold')
            ax.set_ylabel('Current (A)', fontsize=12, fontweight='bold')
            self.canvas.draw()
            return
        
        df = pd.DataFrame(self.measurements)
        
        iv_marker = 'o' if self.plot_settings['iv_style'] in ['points', 'both'] else ''
        iv_line = '-' if self.plot_settings['iv_style'] in ['line', 'both'] else ''
        iv_fmt = f"{iv_marker}{iv_line}" if iv_marker or iv_line else '-'
        
        pv_marker = 's' if self.plot_settings['pv_style'] in ['points', 'both'] else ''
        pv_line = '-' if self.plot_settings['pv_style'] in ['line', 'both'] else ''
        pv_fmt = f"{pv_marker}{pv_line}" if pv_marker or pv_line else '-'
        
        if self.plot_settings['separate_graphs']:
            if self.plot_settings['show_iv']:
                ax1 = self.figure.add_subplot(211)
            else:
                ax1 = None
            
            if self.plot_settings['show_pv']:
                ax2 = self.figure.add_subplot(212)
            else:
                ax2 = None
            
            if ax1:
                ax1.plot(df['V'], df['I'], iv_fmt, 
                        color=self.plot_settings['iv_color'],
                        linewidth=self.plot_settings['line_width'], 
                        markersize=self.plot_settings['marker_size'], 
                        label='I-V Curve')
                ax1.set_xlabel('Voltage (V)', fontsize=11, fontweight='bold')
                ax1.set_ylabel('Current (A)', fontsize=11, fontweight='bold')
                ax1.grid(True, alpha=0.3, linestyle='--')
                ax1.set_title('I-V Characteristic', fontsize=12, fontweight='bold')
                
                if self.plot_settings['show_voc_isc']:
                    if self.voc is not None:
                        ax1.plot(self.voc, 0, 'ms', markersize=10, markeredgewidth=2, 
                                markerfacecolor='magenta', markeredgecolor='darkmagenta', label='Voc')
                        ax1.annotate(f'Voc={self.voc:.3f}V', xy=(self.voc, 0), 
                                   xytext=(self.voc*0.9, 0.1*ax1.get_ylim()[1]),
                                   fontsize=8, fontweight='bold', color='magenta',
                                   bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor='magenta', alpha=0.8),
                                   arrowprops=dict(arrowstyle='->', color='magenta', lw=1.5))
                    
                    if self.isc is not None:
                        ax1.plot(0, self.isc, 'cs', markersize=10, markeredgewidth=2,
                                markerfacecolor='cyan', markeredgecolor='darkcyan', label='Isc')
                        ax1.annotate(f'Isc={self.isc:.3f}A', xy=(0, self.isc), 
                                   xytext=(0.1*ax1.get_xlim()[1], self.isc*0.9),
                                   fontsize=8, fontweight='bold', color='darkcyan',
                                   bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor='darkcyan', alpha=0.8),
                                   arrowprops=dict(arrowstyle='->', color='darkcyan', lw=1.5))
                
                if self.plot_settings['show_mpp']:
                    mpp_idx = df['P'].idxmax()
                    mpp_i = df.loc[mpp_idx, 'I']
                    mpp_v = df.loc[mpp_idx, 'V']
                    ax1.plot(mpp_v, mpp_i, 'g*', markersize=15, markeredgewidth=2, 
                            markeredgecolor='darkgreen', label='MPP')
                
                ax1.legend(loc='best', fontsize=9, framealpha=0.9)
            
            if ax2:
                ax2.plot(df['V'], df['P'], pv_fmt, 
                        color=self.plot_settings['pv_color'],
                        linewidth=self.plot_settings['line_width'], 
                        markersize=self.plot_settings['marker_size'], 
                        label='P-V Curve')
                ax2.set_xlabel('Voltage (V)', fontsize=11, fontweight='bold')
                ax2.set_ylabel('Power (W)', fontsize=11, fontweight='bold')
                ax2.grid(True, alpha=0.3, linestyle='--')
                ax2.set_title('P-V Characteristic', fontsize=12, fontweight='bold')
                
                if self.plot_settings['show_mpp']:
                    mpp_idx = df['P'].idxmax()
                    mpp_v = df.loc[mpp_idx, 'V']
                    mpp_p = df.loc[mpp_idx, 'P']
                    mpp_i = df.loc[mpp_idx, 'I']
                    ax2.plot(mpp_v, mpp_p, 'g*', markersize=15, markeredgewidth=2, 
                            markeredgecolor='darkgreen', label='MPP')
                    ax2.annotate(f'MPP\nV={mpp_v:.3f}V\nI={mpp_i:.3f}A\nP={mpp_p:.3f}W', 
                               xy=(mpp_v, mpp_p),
                               xytext=(mpp_v*1.1, mpp_p*0.85),
                               fontsize=8, fontweight='bold', color='darkgreen',
                               bbox=dict(boxstyle='round,pad=0.4', facecolor='lightgreen', edgecolor='darkgreen', alpha=0.9),
                               arrowprops=dict(arrowstyle='->', color='darkgreen', lw=1.5))
                
                ax2.legend(loc='best', fontsize=9, framealpha=0.9)
        
        else:
            ax1 = self.figure.add_subplot(111)
            ax2 = ax1.twinx()
            
            if self.plot_settings['show_iv']:
                ax1.plot(df['V'], df['I'], iv_fmt, 
                        color=self.plot_settings['iv_color'],
                        linewidth=self.plot_settings['line_width'], 
                        markersize=self.plot_settings['marker_size'], 
                        label='I-V Curve')
            
            ax1.set_xlabel('Voltage (V)', fontsize=12, fontweight='bold')
            if self.plot_settings['show_iv']:
                ax1.set_ylabel('Current (A)', fontsize=12, fontweight='bold')
                ax1.tick_params(axis='y', labelcolor=self.plot_settings['iv_color'])
            else:
                ax1.set_yticks([])
            ax1.grid(True, alpha=0.3, linestyle='--')
            
            title_parts = []
            if self.plot_settings['show_iv']:
                title_parts.append('I-V')
            if self.plot_settings['show_pv']:
                title_parts.append('P-V')
            title = f"PV Panel {' and '.join(title_parts)} Characteristics" if title_parts else "PV Panel"
            ax1.set_title(title, fontsize=14, fontweight='bold', pad=15)
            
            if self.plot_settings['show_pv']:
                ax2.plot(df['V'], df['P'], pv_fmt, 
                        color=self.plot_settings['pv_color'],
                        linewidth=self.plot_settings['line_width'], 
                        markersize=self.plot_settings['marker_size'], 
                        label='P-V Curve')
                ax2.set_ylabel('Power (W)', fontsize=12, fontweight='bold')
                ax2.tick_params(axis='y', labelcolor=self.plot_settings['pv_color'])
            else:
                ax2.set_yticks([])
            
            if self.plot_settings['show_voc_isc']:
                if self.voc is not None:
                    ax1.plot(self.voc, 0, 'ms', markersize=12, markeredgewidth=2, 
                            markerfacecolor='magenta', markeredgecolor='darkmagenta', label='Voc')
                    ax1.annotate(f'Voc={self.voc:.3f}V', xy=(self.voc, 0), 
                               xytext=(self.voc*0.93, 0.07*ax1.get_ylim()[1]),
                               fontsize=9, fontweight='bold', color='magenta',
                               bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor='magenta', alpha=0.8),
                               arrowprops=dict(arrowstyle='->', color='magenta', lw=1.5))
                
                if self.isc is not None:
                    ax1.plot(0, self.isc, 'cs', markersize=12, markeredgewidth=2,
                            markerfacecolor='cyan', markeredgecolor='darkcyan', label='Isc')
                    ax1.annotate(f'Isc={self.isc:.3f}A', xy=(0, self.isc), 
                               xytext=(0.07*ax1.get_xlim()[1], self.isc*0.93),
                               fontsize=9, fontweight='bold', color='darkcyan',
                               bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor='darkcyan', alpha=0.8),
                               arrowprops=dict(arrowstyle='->', color='darkcyan', lw=1.5))
            
            if self.plot_settings['show_mpp'] and len(df) > 0:
                mpp_idx = df['P'].idxmax()
                mpp_v = df.loc[mpp_idx, 'V']
                mpp_p = df.loc[mpp_idx, 'P']
                mpp_i = df.loc[mpp_idx, 'I']
                ax2.plot(mpp_v, mpp_p, 'g*', markersize=20, markeredgewidth=2, 
                        markeredgecolor='darkgreen', label=f'MPP')
                ax1.plot(mpp_v, mpp_i, 'g*', markersize=20, markeredgewidth=2, 
                        markeredgecolor='darkgreen')
                ax2.annotate(f'MPP\nV={mpp_v:.3f}V\nI={mpp_i:.3f}A\nP={mpp_p:.3f}W', 
                           xy=(mpp_v, mpp_p),
                           xytext=(mpp_v*1.1, mpp_p*0.9),
                           fontsize=9, fontweight='bold', color='darkgreen',
                           bbox=dict(boxstyle='round,pad=0.5', facecolor='lightgreen', edgecolor='darkgreen', alpha=0.9),
                           arrowprops=dict(arrowstyle='->', color='darkgreen', lw=2))
            
            if self.plot_settings['show_iv']:
                ax1.legend(loc='upper right', fontsize=10, framealpha=0.9)
            if self.plot_settings['show_pv']:
                ax2.legend(loc='lower right', fontsize=10, framealpha=0.9)
            
            if self.plot_settings.get('v_min') is not None or self.plot_settings.get('v_max') is not None:
                ax1.set_xlim(self.plot_settings.get('v_min'), self.plot_settings.get('v_max'))
            if self.plot_settings.get('i_min') is not None or self.plot_settings.get('i_max') is not None:
                if self.plot_settings['show_iv']:
                    ax1.set_ylim(self.plot_settings.get('i_min'), self.plot_settings.get('i_max'))
            if self.plot_settings.get('p_min') is not None or self.plot_settings.get('p_max') is not None:
                if self.plot_settings['show_pv']:
                    ax2.set_ylim(self.plot_settings.get('p_min'), self.plot_settings.get('p_max'))
        
        self.figure.tight_layout()
        self.canvas.draw()
    
    def export_dialog(self):
        from PyQt5.QtWidgets import QDialog, QRadioButton, QButtonGroup
        
        if not self.measurements:
            QMessageBox.critical(self, "Error", "No data to export")
            return
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Export Data/Plot")
        dialog.setMinimumWidth(400)
        dialog_layout = QVBoxLayout()
        
        type_group = QGroupBox("Export Type")
        type_layout = QVBoxLayout()
        type_button_group = QButtonGroup(dialog)
        
        data_radio = QRadioButton("📊 Data Table")
        data_radio.setChecked(True)
        type_button_group.addButton(data_radio, 1)
        type_layout.addWidget(data_radio)
        
        plot_radio = QRadioButton("📈 Plot/Graph")
        type_button_group.addButton(plot_radio, 2)
        type_layout.addWidget(plot_radio)
        
        type_group.setLayout(type_layout)
        dialog_layout.addWidget(type_group)
        
        format_group = QGroupBox("Format")
        format_layout = QVBoxLayout()
        format_button_group = QButtonGroup(dialog)
        
        csv_radio = QRadioButton("CSV (.csv) - Comma Separated")
        csv_radio.setChecked(True)
        format_button_group.addButton(csv_radio, 1)
        format_layout.addWidget(csv_radio)
        
        excel_radio = QRadioButton("Excel (.xlsx) - Spreadsheet")
        format_button_group.addButton(excel_radio, 2)
        format_layout.addWidget(excel_radio)
        
        txt_radio = QRadioButton("Text (.txt) - Plain Text")
        format_button_group.addButton(txt_radio, 3)
        format_layout.addWidget(txt_radio)
        
        png_radio = QRadioButton("PNG (.png) - High Quality Image")
        format_button_group.addButton(png_radio, 4)
        format_layout.addWidget(png_radio)
        
        jpg_radio = QRadioButton("JPG (.jpg) - Compressed Image")
        format_button_group.addButton(jpg_radio, 5)
        format_layout.addWidget(jpg_radio)
        
        pdf_radio = QRadioButton("PDF (.pdf) - Document")
        format_button_group.addButton(pdf_radio, 6)
        format_layout.addWidget(pdf_radio)
        
        format_group.setLayout(format_layout)
        dialog_layout.addWidget(format_group)
        
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("✓ Choose Location & Export")
        ok_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 8px;")
        ok_btn.clicked.connect(dialog.accept)
        btn_layout.addWidget(ok_btn)
        
        cancel_btn = QPushButton("✗ Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        btn_layout.addWidget(cancel_btn)
        
        dialog_layout.addLayout(btn_layout)
        dialog.setLayout(dialog_layout)
        
        if dialog.exec_() == QDialog.Accepted:
            is_plot = type_button_group.checkedId() == 2
            format_id = format_button_group.checkedId()
            
            if is_plot:
                if format_id == 4:
                    self.export_plot_with_dialog('png')
                elif format_id == 5:
                    self.export_plot_with_dialog('jpg')
                elif format_id == 6:
                    self.export_plot_with_dialog('pdf')
                else:
                    QMessageBox.warning(self, "Warning", "Please select a valid plot format (PNG, JPG, or PDF)")
            else:
                if format_id == 1:
                    self.export_data_with_dialog('csv')
                elif format_id == 2:
                    self.export_data_with_dialog('excel')
                elif format_id == 3:
                    self.export_data_with_dialog('txt')
                else:
                    QMessageBox.warning(self, "Warning", "Please select a valid data format (CSV, Excel, or TXT)")
    
    def export_data_with_dialog(self, format_type):
        if not self.measurements:
            QMessageBox.critical(self, "Error", "No data to export")
            return
        
        try:
            df = pd.DataFrame(self.measurements)
            df.insert(0, 'Point', range(1, len(df) + 1))
            
            if format_type == 'csv':
                filename, _ = QFileDialog.getSaveFileName(self, "Choose Save Location - CSV Export", 
                                                         f"IV_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                                         "CSV files (*.csv)")
                if filename:
                    df.to_csv(filename, index=False)
                    QMessageBox.information(self, "Export Successful", f"Data exported to:\n{filename}")
            
            elif format_type == 'excel':
                filename, _ = QFileDialog.getSaveFileName(self, "Choose Save Location - Excel Export", 
                                                         f"IV_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                                         "Excel files (*.xlsx)")
                if filename:
                    df.to_excel(filename, index=False)
                    QMessageBox.information(self, "Export Successful", f"Data exported to:\n{filename}")
            
            elif format_type == 'txt':
                filename, _ = QFileDialog.getSaveFileName(self, "Choose Save Location - Text Export", 
                                                         f"IV_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                                                         "Text files (*.txt)")
                if filename:
                    with open(filename, 'w') as f:
                        f.write(self.results_text.toPlainText())
                        f.write("\n\n" + "="*50 + "\n")
                        f.write("RAW MEASUREMENT DATA\n")
                        f.write("="*50 + "\n\n")
                        f.write(df.to_string(index=False))
                    QMessageBox.information(self, "Export Successful", f"Data exported to:\n{filename}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export data:\n{str(e)}")
    
    def export_plot_with_dialog(self, format_type):
        if not self.measurements:
            QMessageBox.critical(self, "Error", "No plot to export")
            return
        
        try:
            ext = format_type
            filter_str = f"{ext.upper()} files (*.{ext})"
            filename, _ = QFileDialog.getSaveFileName(self, f"Choose Save Location - {ext.upper()} Export", 
                                                     f"IV_Plot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}",
                                                     filter_str)
            
            if filename:
                self.figure.savefig(filename, dpi=300, bbox_inches='tight')
                QMessageBox.information(self, "Export Successful", f"Plot exported to:\n{filename}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export plot:\n{str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PVIVAnalyzer()
    window.show()
    sys.exit(app.exec_())
