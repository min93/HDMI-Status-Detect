import sys
import os
import json
import datetime
import wmi
import time
import psutil
import GPUtil
import pythoncom
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QPushButton, QTextEdit, QTableWidget, 
                            QTableWidgetItem, QMessageBox, QHeaderView, QDateEdit, QSplitter,
                            QSystemTrayIcon, QMenu, QAction, QProgressBar, QFrame)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QDate, QTimer
from PyQt5.QtGui import QFont, QColor, QIcon
from win10toast import ToastNotifier
# แก้ไขการนำเข้า pycec
# from pycec import pycec

# กำหนดค่าคงที่
APP_TITLE = "ตรวจสอบสถานะ HDMI บน Windows"
APP_ICON = "monitor.ico"
LOG_FILE = "hdmi_monitor_log.json"
HDMI_CHECK_INTERVAL = 5  # ตรวจสอบทุก 5 วินาที
SYSTEM_CHECK_INTERVAL = 1  # ตรวจสอบทุก 1 วินาที
CONNECTION_TYPE_HDMI = 5  # รหัสสำหรับการเชื่อมต่อ HDMI
NOTIFICATION_DURATION = 5  # ระยะเวลาแสดงการแจ้งเตือน (วินาที)

# ประเภทการเชื่อมต่อ
CONNECTION_TYPES = {
    0: "VGA (HD15)",
    1: "TV Composite",
    2: "TV S-Video",
    3: "TV Component (RGB)",
    4: "DVI",
    5: "HDMI",
    6: "LVDS",
    8: "D-Jpn",
    9: "SDI",
    10: "DisplayPort",
    11: "HDMI External",
    12: "Virtual",
}

# รหัสข้อผิดพลาดของอุปกรณ์
ERROR_CODES = {
    0: "อุปกรณ์ทำงานปกติ",
    1: "อุปกรณ์ยังตั้งค่าไม่ถูกต้อง",
    2: "Windows ไม่สามารถโหลดไดรเวอร์ของอุปกรณ์นี้",
    3: "ไดรเวอร์อาจมีปัญหาหรือระบบเมมโมรีไม่พอ",
    4: "อุปกรณ์นี้ทำงานไม่ถูกต้อง",
    22: "อุปกรณ์นี้ถูกปิดการใช้งาน",
    28: "ไม่มีไดรเวอร์ติดตั้งสำหรับอุปกรณ์นี้",
    43: "Windows หยุดอุปกรณ์นี้เนื่องจากพบปัญหา"
}

# สถานะไอคอน
STATUS_ICONS = {
    "normal": "✅",
    "warning": "⚠️",
    "error": "❌",
    "unknown": "❓"
}

class HDMIMonitor(QThread):
    """
    เธรดสำหรับการตรวจสอบสถานะ HDMI และไดรเวอร์การ์ดจอ
    """
    update_signal = pyqtSignal(dict)
    log_signal = pyqtSignal(dict)
    notification_signal = pyqtSignal(str, str)  # สัญญาณสำหรับการแจ้งเตือน (หัวข้อ, ข้อความ)
    
    def __init__(self, parent=None):
        super(HDMIMonitor, self).__init__(parent)
        self.running = True
        self.interval = HDMI_CHECK_INTERVAL
        
        # สถานะครั้งล่าสุด เพื่อตรวจสอบการเปลี่ยนแปลง
        self.last_status = {
            'hdmi_connected': False,
            'hdmi_active': False,
            'gpu_status': {}
        }
        
    def run(self):
        pythoncom.CoInitialize()
        while self.running:
            try:
                status = self.check_hdmi_status()
                self.update_signal.emit(status)
                
                # ตรวจสอบการเปลี่ยนแปลงสถานะและบันทึกล็อกหากมีการเปลี่ยนแปลง
                if self._has_status_changed(status):
                    log_entry = {
                        'timestamp': datetime.datetime.now().isoformat(),
                        'status': status,
                        'event_type': self._determine_event_type(status)
                    }
                    self.log_signal.emit(log_entry)
                    
                    # ส่งการแจ้งเตือนเมื่อมีการเปลี่ยนแปลงสถานะ
                    event_type = self._determine_event_type(status)
                    notification_title = "การแจ้งเตือนสถานะ HDMI"
                    notification_message = self._get_notification_message(event_type, status)
                    self.notification_signal.emit(notification_title, notification_message)
                    
                    self.last_status = status.copy()
                    
                time.sleep(self.interval)
            except Exception as e:
                error_msg = f"Error in monitoring thread: {str(e)}"
                print(error_msg)
                self.notification_signal.emit("เกิดข้อผิดพลาด", error_msg)
                time.sleep(self.interval)
    
    def _has_status_changed(self, current_status):
        """ตรวจสอบว่าสถานะมีการเปลี่ยนแปลงหรือไม่"""
        if self.last_status['hdmi_connected'] != current_status['hdmi_connected']:
            return True
        if self.last_status['hdmi_active'] != current_status['hdmi_active']:
            return True
            
        # ตรวจสอบการเปลี่ยนแปลงสถานะ GPU
        for gpu_id, gpu_info in current_status['gpu_status'].items():
            if gpu_id in self.last_status['gpu_status']:
                if self.last_status['gpu_status'][gpu_id]['error_code'] != gpu_info['error_code']:
                    return True
            else:
                return True
                
        return False
    
    def _determine_event_type(self, status):
        """ระบุประเภทของเหตุการณ์ที่เกิดขึ้น"""
        if not status['hdmi_connected']:
            return "HDMI_DISCONNECTED"
        elif not status['hdmi_active']:
            return "HDMI_INACTIVE"
        
        # ตรวจสอบว่ามี GPU ที่มีปัญหาหรือไม่
        for gpu_info in status['gpu_status'].values():
            if gpu_info['error_code'] != 0:
                return f"GPU_ERROR_CODE_{gpu_info['error_code']}"
        
        return "NORMAL"
    
    def _get_notification_message(self, event_type, status):
        """สร้างข้อความแจ้งเตือนตามประเภทเหตุการณ์"""
        if event_type == "HDMI_DISCONNECTED":
            return "HDMI ถูกถอดออกจากเครื่อง"
        elif event_type == "HDMI_INACTIVE":
            return "HDMI เชื่อมต่ออยู่แต่ไม่ทำงาน"
        elif event_type.startswith("GPU_ERROR_CODE_"):
            error_code = event_type.split("_")[-1]
            for gpu_id, gpu_info in status['gpu_status'].items():
                if str(gpu_info['error_code']) == error_code:
                    return f"การ์ดจอ {gpu_info['name']} มีปัญหา: {gpu_info['error_description']}"
            return f"การ์ดจอมีปัญหา (รหัส {error_code})"
        elif event_type == "NORMAL":
            return "การเชื่อมต่อ HDMI กลับมาทำงานปกติแล้ว"
        else:
            return f"เหตุการณ์: {event_type}"
    
    def check_hdmi_status(self):
        """ตรวจสอบสถานะการเชื่อมต่อ HDMI และไดรเวอร์การ์ดจอ"""
        pythoncom.CoInitialize()
        result = {
            'hdmi_connected': False,
            'hdmi_active': False,
            'hdmi_devices': [],
            'gpu_status': {},
            'driver_check_time': datetime.datetime.now().isoformat()
        }
        
        try:
            # ตรวจสอบการเชื่อมต่อ HDMI
            try:
                wmi_obj = wmi.WMI(namespace="root\\wmi")
                monitors_found = False
                for monitor in wmi_obj.WmiMonitorConnectionParams():
                    monitors_found = True
                    try:
                        conn_type = monitor.VideoOutputTechnology
                        is_active = monitor.Active
                        
                        monitor_info = {
                            'instance_name': monitor.InstanceName,
                            'connection_type': conn_type,
                            'active': is_active,
                            'status': 'normal'
                        }
                        
                        result['hdmi_devices'].append(monitor_info)
                        
                        if conn_type == CONNECTION_TYPE_HDMI:  # 5 = HDMI
                            result['hdmi_connected'] = True
                            if is_active:
                                result['hdmi_active'] = True
                    except Exception as monitor_error:
                        monitor_info = {
                            'instance_name': getattr(monitor, 'InstanceName', 'Unknown'),
                            'connection_type': -1,
                            'active': False,
                            'status': f'error: {str(monitor_error)}'
                        }
                        result['hdmi_devices'].append(monitor_info)
                
                if not monitors_found:
                    result['wmi_error'] = "ไม่พบจอภาพที่เชื่อมต่อ"
                    
            except Exception as wmi_error:
                result['wmi_error'] = str(wmi_error)
            
            # ตรวจสอบสถานะไดรเวอร์การ์ดจอ
            try:
                wmi_obj = wmi.WMI()
                gpu_found = False
                
                # ลองใช้ GPUtil ก่อน
                try:
                    gpus = GPUtil.getGPUs()
                    if gpus:
                        for idx, gpu in enumerate(gpus):
                            gpu_found = True
                            gpu_id = f"gpu_{idx}"
                            result['gpu_status'][gpu_id] = {
                                'name': gpu.name,
                                'status': "OK",
                                'error_code': 0,
                                'driver_version': "Unknown",  # GPUtil ไม่มีข้อมูลนี้
                                'driver_status': "normal",
                                'load': f"{gpu.load * 100:.1f}%",
                                'temperature': f"{gpu.temperature}°C",
                                'memory_used': f"{gpu.memoryUsed}MB / {gpu.memoryTotal}MB",
                                'error_description': "ทำงานปกติ"
                            }
                except Exception:
                    pass  # ถ้าใช้ GPUtil ไม่ได้ จะใช้ WMI แทน
                
                # ถ้ายังไม่พบ GPU ให้ใช้ WMI
                if not gpu_found:
                    for idx, gpu in enumerate(wmi_obj.Win32_VideoController()):
                        gpu_found = True
                        gpu_id = f"gpu_{idx}"
                        
                        try:
                            if gpu.ConfigManagerErrorCode == 0:
                                if gpu.Status == "OK":
                                    driver_status = "normal"
                                else:
                                    driver_status = "warning"
                            else:
                                driver_status = "error"
                        except:
                            driver_status = "unknown"
                        
                        result['gpu_status'][gpu_id] = {
                            'name': getattr(gpu, 'Name', 'Unknown GPU'),
                            'status': getattr(gpu, 'Status', 'Unknown'),
                            'error_code': getattr(gpu, 'ConfigManagerErrorCode', -1),
                            'driver_version': getattr(gpu, 'DriverVersion', 'Unknown'),
                            'driver_status': driver_status,
                            'driver_date': getattr(gpu, 'DriverDate', 'Unknown'),
                            'error_description': self._get_error_description(
                                getattr(gpu, 'ConfigManagerErrorCode', -1)
                            )
                        }
                
                if not gpu_found:
                    result['gpu_status']['gpu_0'] = {
                        'name': "ไม่พบการ์ดจอ",
                        'status': "Not found",
                        'error_code': -1,
                        'driver_version': "N/A",
                        'driver_status': "error",
                        'error_description': "ไม่พบการ์ดจอในระบบ หรือไม่สามารถเข้าถึงข้อมูลได้"
                    }
                    
            except Exception as gpu_error:
                result['gpu_error'] = str(gpu_error)
                result['gpu_status']['gpu_error'] = {
                    'name': "Error checking GPU",
                    'status': "Error",
                    'error_code': -1,
                    'driver_version': "Unknown",
                    'driver_status': "error",
                    'error_description': str(gpu_error)
                }
                
        except Exception as e:
            error_msg = f"Error checking HDMI status: {str(e)}"
            print(error_msg)
            result['error'] = error_msg
            
        return result
    
    def _get_error_description(self, error_code):
        """แปลงรหัสข้อผิดพลาดเป็นคำอธิบาย"""
        return ERROR_CODES.get(error_code, f"ข้อผิดพลาดที่ไม่รู้จัก (รหัส {error_code})")
    
    def stop(self):
        self.running = False


class LogManager:
    """จัดการการบันทึกและเรียกดูล็อก"""
    
    def __init__(self, log_file=LOG_FILE):
        self.log_file = log_file
        self._ensure_log_file()
    
    def _ensure_log_file(self):
        """สร้างไฟล์ล็อกหากยังไม่มี"""
        if not os.path.exists(self.log_file):
            with open(self.log_file, 'w', encoding='utf-8') as f:
                json.dump([], f, ensure_ascii=False)
    
    def add_log_entry(self, log_entry):
        """เพิ่มรายการล็อกใหม่"""
        try:
            logs = self.get_logs()
            logs.append(log_entry)
            
            with open(self.log_file, 'w', encoding='utf-8') as f:
                json.dump(logs, f, ensure_ascii=False, indent=2)
                
            return True
        except Exception as e:
            print(f"Error adding log entry: {str(e)}")
            return False
    
    def get_logs(self):
        """อ่านล็อกทั้งหมด"""
        try:
            with open(self.log_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error reading logs: {str(e)}")
            return []
    
    def get_logs_by_date(self, date):
        """อ่านล็อกตามวันที่"""
        logs = self.get_logs()
        date_str = date.toString("yyyy-MM-dd")
        
        filtered_logs = []
        for log in logs:
            log_date = log['timestamp'].split('T')[0]
            if log_date == date_str:
                filtered_logs.append(log)
                
        return filtered_logs


class NotificationManager:
    """จัดการการแจ้งเตือนบน Windows"""
    
    def __init__(self):
        self.toaster = ToastNotifier()
    
    def show_notification(self, title, message, duration=NOTIFICATION_DURATION):
        """แสดงการแจ้งเตือนบน Windows"""
        try:
            self.toaster.show_toast(
                title,
                message,
                icon_path=APP_ICON,  # ใช้ไอคอนแอพ
                duration=duration,
                threaded=True  # ทำงานในเธรดแยก ไม่บล็อกโปรแกรมหลัก
            )
            return True
        except Exception as e:
            print(f"Error showing notification: {str(e)}")
            return False


class SystemMonitor(QThread):
    """เธรดสำหรับการตรวจสอบสถานะระบบ (CPU, RAM, GPU)"""
    
    update_signal = pyqtSignal(dict)
    
    def __init__(self, parent=None):
        super(SystemMonitor, self).__init__(parent)
        self.running = True
        self.interval = SYSTEM_CHECK_INTERVAL
        
    def run(self):
        while self.running:
            try:
                status = self.check_system_status()
                self.update_signal.emit(status)
                time.sleep(self.interval)
            except Exception as e:
                print(f"Error in system monitoring thread: {str(e)}")
                time.sleep(self.interval)
    
    def check_system_status(self):
        """ตรวจสอบสถานะ CPU, RAM และ GPU"""
        result = {
            'cpu': {},
            'memory': {},
            'gpu': [],
            'temperatures': {}
        }
        
        try:
            # CPU
            cpu_percent = psutil.cpu_percent(interval=None)
            cpu_count = psutil.cpu_count(logical=True)
            cpu_freq = psutil.cpu_freq()
            
            result['cpu'] = {
                'percent': cpu_percent,
                'cores': cpu_count,
                'frequency': cpu_freq.current if cpu_freq else 0
            }
            
            # Memory
            memory = psutil.virtual_memory()
            result['memory'] = {
                'total': memory.total,
                'available': memory.available,
                'percent': memory.percent,
                'used': memory.used
            }
            
            # อุณหภูมิ - ใช้ psutil แทนการใช้ WMI
            try:
                # ตรวจสอบว่า psutil มีข้อมูลอุณหภูมิหรือไม่
                if hasattr(psutil, 'sensors_temperatures'):
                    temps = psutil.sensors_temperatures()
                    if temps:
                        # CPU Temperature
                        for name, entries in temps.items():
                            if 'cpu' in name.lower() or 'coretemp' in name.lower():
                                if entries:
                                    result['temperatures']['cpu'] = entries[0].current
                            elif 'acpi' in name.lower() or 'system' in name.lower() or 'motherboard' in name.lower():
                                if entries:
                                    result['temperatures']['mainboard'] = entries[0].current
                else:
                    # ถ้า psutil ไม่มีข้อมูล temperature
                    result['temperatures']['error'] = "ไม่สามารถอ่านอุณหภูมิได้ (psutil ไม่สนับสนุนบนระบบนี้)"
            except Exception as e:
                result['temperatures']['error'] = f"ไม่สามารถอ่านอุณหภูมิได้: {str(e)}"
            
            # GPU
            try:
                gpus = GPUtil.getGPUs()
                if not gpus:
                    # ถ้าไม่พบ GPU ให้ใช้ wmi เพื่อดึงข้อมูลการ์ดจอแทน
                    wmi_obj = wmi.WMI()
                    for idx, gpu in enumerate(wmi_obj.Win32_VideoController()):
                        result['gpu'].append({
                            'name': gpu.Name,
                            'driver_version': gpu.DriverVersion,
                            'driver_status': "ปกติ" if gpu.ConfigManagerErrorCode == 0 else "มีปัญหา",
                            'status_message': self._get_error_description(gpu.ConfigManagerErrorCode)
                        })
                    if not result['gpu']:
                        result['gpu'].append({'error': "ไม่พบข้อมูล GPU"})
                else:
                    for gpu in gpus:
                        result['gpu'].append({
                            'name': gpu.name,
                            'load': gpu.load * 100,  # Convert to percentage
                            'memory_total': gpu.memoryTotal,
                            'memory_used': gpu.memoryUsed,
                            'memory_percent': (gpu.memoryUsed / gpu.memoryTotal) * 100 if gpu.memoryTotal > 0 else 0,
                            'temperature': gpu.temperature
                        })
            except Exception as e:
                print(f"Error getting GPU info: {str(e)}")
                result['gpu'].append({'error': str(e)})
                
        except Exception as e:
            print(f"Error checking system status: {str(e)}")
            
        if not result['gpu']:
            result['gpu'].append({'error': "ไม่พบข้อมูล GPU"})
            
        return result
    
    def _get_error_description(self, error_code):
        """แปลงรหัสข้อผิดพลาดเป็นคำอธิบาย"""
        return ERROR_CODES.get(error_code, f"ข้อผิดพลาดที่ไม่รู้จัก (รหัส {error_code})")
    
    def stop(self):
        self.running = False


class MainWindow(QMainWindow):
    """หน้าต่างหลักของแอปพลิเคชัน"""
    
    def __init__(self):
        super(MainWindow, self).__init__()
        
        self.setWindowTitle(APP_TITLE)
        self.setMinimumSize(800, 600)
        
        # ตั้งค่าไอคอน
        self.setWindowIcon(QIcon(APP_ICON))
        
        self.log_manager = LogManager()
        self.notification_manager = NotificationManager()
        self.init_ui()
        
        # ตั้งค่า System Tray
        self.setup_system_tray()
        
        # เริ่มการตรวจสอบ HDMI
        self.hdmi_monitor = HDMIMonitor()
        self.hdmi_monitor.update_signal.connect(self.update_status)
        self.hdmi_monitor.log_signal.connect(self.add_log_entry)
        self.hdmi_monitor.notification_signal.connect(self.show_notification)
        self.hdmi_monitor.start()
        
        # เริ่มการตรวจสอบระบบ (CPU, RAM, GPU)
        self.system_monitor = SystemMonitor()
        self.system_monitor.update_signal.connect(self.update_system_status)
        self.system_monitor.start()
        
        # ตั้งเวลาอัปเดตสถานะทุก 1 วินาที (สำหรับแสดงเวลาปัจจุบัน)
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_current_time)
        self.timer.start(1000)
        
        # อัปเดตรายการล็อกเมื่อเริ่มต้น
        self.update_log_view(QDate.currentDate())
    
    def setup_system_tray(self):
        """ตั้งค่า System Tray Icon"""
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon(APP_ICON))
        
        # สร้างเมนูสำหรับ System Tray
        tray_menu = QMenu()
        
        # เพิ่มรายการเมนู
        show_action = QAction("แสดงหน้าต่าง", self)
        show_action.triggered.connect(self.show)
        tray_menu.addAction(show_action)
        
        refresh_action = QAction("รีเฟรชสถานะ", self)
        refresh_action.triggered.connect(self.refresh_status)
        tray_menu.addAction(refresh_action)
        
        tray_menu.addSeparator()
        
        exit_action = QAction("ออกจากโปรแกรม", self)
        exit_action.triggered.connect(self.close_application)
        tray_menu.addAction(exit_action)
        
        # ตั้งค่าเมนูให้กับ System Tray
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
        
        # เชื่อมต่อสัญญาณเมื่อคลิกที่ไอคอน
        self.tray_icon.activated.connect(self.tray_icon_activated)
    
    def tray_icon_activated(self, reason):
        """จัดการเมื่อมีการคลิกที่ไอคอนใน System Tray"""
        if reason == QSystemTrayIcon.DoubleClick or reason == QSystemTrayIcon.Trigger:
            if self.isVisible():
                self.hide()
            else:
                self.show()
                self.activateWindow()
    
    def show_notification(self, title, message):
        """แสดงการแจ้งเตือน"""
        self.notification_manager.show_notification(title, message)
        
        # อัปเดตข้อความใน System Tray
        self.tray_icon.setToolTip(f"{title}: {message}")
        
        # แสดงการแจ้งเตือนใน System Tray
        if self.isHidden():
            self.tray_icon.showMessage(title, message, QSystemTrayIcon.Information, 5000)
    
    def keyPressEvent(self, event):
        """จัดการเหตุการณ์การกดปุ่มคีย์บอร์ด"""
        # ตรวจสอบการกด Ctrl+C
        if event.key() == Qt.Key_C and event.modifiers() == Qt.ControlModifier:
            QMessageBox.information(self, "หยุดการทำงาน", "กำลังหยุดการทำงานของโปรแกรม...")
            self.close_application()
        else:
            super(MainWindow, self).keyPressEvent(event)
    
    def init_ui(self):
        """สร้างอินเตอร์เฟซผู้ใช้"""
        # สร้าง main widget และ layout หลัก
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        self.setCentralWidget(main_widget)
        
        # แท็บหลัก
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # แท็บสถานะปัจจุบัน
        self.status_tab = QWidget()
        self.status_layout = QVBoxLayout()
        self.status_tab.setLayout(self.status_layout)
        
        # แท็บประวัติล็อก
        self.log_tab = QWidget()
        self.log_layout = QVBoxLayout()
        self.log_tab.setLayout(self.log_layout)
        
        # แท็บสถานะระบบ
        self.system_tab = QWidget()
        self.system_layout = QVBoxLayout()
        self.system_tab.setLayout(self.system_layout)
        
        self.tabs.addTab(self.status_tab, "Status")
        self.tabs.addTab(self.system_tab, "Hardware Status")
        self.tabs.addTab(self.log_tab, "Log")
        
        # เพิ่ม footer
        footer = QLabel("Developed by rattapon@seanetasia.com | © Seanet Asia")
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet("color: #666; padding: 5px;")
        main_layout.addWidget(footer)
        
        # --- หน้าสถานะปัจจุบัน ---
        # เวลาปัจจุบัน
        self.time_label = QLabel("เวลาปัจจุบัน: ")
        self.time_label.setFont(QFont("Arial", 12))
        self.status_layout.addWidget(self.time_label)
        
        # --- หน้าสถานะระบบ ---
        # CPU
        cpu_group = QWidget()
        cpu_layout = QVBoxLayout()
        cpu_group.setLayout(cpu_layout)
        
        cpu_header = QLabel("สถานะ CPU:")
        cpu_header.setFont(QFont("Arial", 12, QFont.Bold))
        cpu_layout.addWidget(cpu_header)
        
        # CPU Usage
        cpu_usage_layout = QHBoxLayout()
        cpu_usage_layout.addWidget(QLabel("การใช้งาน CPU:"))
        self.cpu_percent_label = QLabel("0%")
        cpu_usage_layout.addWidget(self.cpu_percent_label)
        cpu_layout.addLayout(cpu_usage_layout)
        
        self.cpu_progress = QProgressBar()
        self.cpu_progress.setRange(0, 100)
        self.cpu_progress.setValue(0)
        cpu_layout.addWidget(self.cpu_progress)
        
        # CPU Info
        self.cpu_info_label = QLabel("จำนวน Core: 0 | ความเร็ว: 0 MHz")
        cpu_layout.addWidget(self.cpu_info_label)
        
        self.system_layout.addWidget(cpu_group)
        
        # Separator
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        self.system_layout.addWidget(line)
        
        # RAM
        ram_group = QWidget()
        ram_layout = QVBoxLayout()
        ram_group.setLayout(ram_layout)
        
        ram_header = QLabel("สถานะหน่วยความจำ (RAM):")
        ram_header.setFont(QFont("Arial", 12, QFont.Bold))
        ram_layout.addWidget(ram_header)
        
        # RAM Usage
        ram_usage_layout = QHBoxLayout()
        ram_usage_layout.addWidget(QLabel("การใช้งาน RAM:"))
        self.ram_percent_label = QLabel("0%")
        ram_usage_layout.addWidget(self.ram_percent_label)
        ram_layout.addLayout(ram_usage_layout)
        
        self.ram_progress = QProgressBar()
        self.ram_progress.setRange(0, 100)
        self.ram_progress.setValue(0)
        ram_layout.addWidget(self.ram_progress)
        
        # RAM Info
        self.ram_info_label = QLabel("ใช้งาน: 0 GB / ทั้งหมด: 0 GB")
        ram_layout.addWidget(self.ram_info_label)
        
        self.system_layout.addWidget(ram_group)
        
        # Separator
        line2 = QFrame()
        line2.setFrameShape(QFrame.HLine)
        line2.setFrameShadow(QFrame.Sunken)
        self.system_layout.addWidget(line2)
        
        # GPU
        gpu_monitor_group = QWidget()
        gpu_monitor_layout = QVBoxLayout()
        gpu_monitor_group.setLayout(gpu_monitor_layout)
        
        gpu_header = QLabel("สถานะการ์ดจอ (GPU):")
        gpu_header.setFont(QFont("Arial", 12, QFont.Bold))
        gpu_monitor_layout.addWidget(gpu_header)
        
        self.gpu_info_text = QTextEdit()
        self.gpu_info_text.setReadOnly(True)
        gpu_monitor_layout.addWidget(self.gpu_info_text)
        
        self.system_layout.addWidget(gpu_monitor_group)
        
        # สถานะ HDMI
        hdmi_group = QWidget()
        hdmi_layout = QVBoxLayout()
        hdmi_group.setLayout(hdmi_layout)
        
        self.hdmi_status_label = QLabel("สถานะการเชื่อมต่อ HDMI: กำลังตรวจสอบ...")
        self.hdmi_status_label.setFont(QFont("Arial", 12, QFont.Bold))
        hdmi_layout.addWidget(self.hdmi_status_label)
        
        self.hdmi_active_label = QLabel("สถานะการทำงาน HDMI: กำลังตรวจสอบ...")
        hdmi_layout.addWidget(self.hdmi_active_label)
        
        self.status_layout.addWidget(hdmi_group)
        
        # สถานะไดรเวอร์การ์ดจอ
        gpu_group = QWidget()
        gpu_layout = QVBoxLayout()
        gpu_group.setLayout(gpu_layout)
        
        gpu_label = QLabel("สถานะไดรเวอร์การ์ดจอ:")
        gpu_label.setFont(QFont("Arial", 12, QFont.Bold))
        gpu_layout.addWidget(gpu_label)
        
        self.gpu_status_text = QTextEdit()
        self.gpu_status_text.setReadOnly(True)
        gpu_layout.addWidget(self.gpu_status_text)
        
        self.status_layout.addWidget(gpu_group)
        
        # พื้นที่สำหรับแสดงข้อมูลจอภาพทั้งหมด
        monitor_group = QWidget()
        monitor_layout = QVBoxLayout()
        monitor_group.setLayout(monitor_layout)
        
        monitor_label = QLabel("จอภาพที่ตรวจพบทั้งหมด:")
        monitor_label.setFont(QFont("Arial", 12, QFont.Bold))
        monitor_layout.addWidget(monitor_label)
        
        self.monitor_info_text = QTextEdit()
        self.monitor_info_text.setReadOnly(True)
        monitor_layout.addWidget(self.monitor_info_text)
        
        self.status_layout.addWidget(monitor_group)
        
        # ปุ่มกระทำ
        action_group = QWidget()
        action_layout = QHBoxLayout()
        action_group.setLayout(action_layout)
        
        refresh_btn = QPushButton("รีเฟรชสถานะทันที")
        refresh_btn.clicked.connect(self.refresh_status)
        action_layout.addWidget(refresh_btn)
        
        minimize_btn = QPushButton("ซ่อนไปที่ System Tray")
        minimize_btn.clicked.connect(self.hide)
        action_layout.addWidget(minimize_btn)
        
        self.status_layout.addWidget(action_group)
        
        # --- หน้าประวัติล็อก ---
        date_group = QWidget()
        date_layout = QHBoxLayout()
        date_group.setLayout(date_layout)
        
        date_layout.addWidget(QLabel("เลือกวันที่:"))
        self.date_selector = QDateEdit()
        self.date_selector.setDate(QDate.currentDate())
        self.date_selector.setCalendarPopup(True)
        self.date_selector.dateChanged.connect(self.date_changed)
        date_layout.addWidget(self.date_selector)
        
        date_layout.addStretch()
        
        self.log_layout.addWidget(date_group)
        
        # ตารางแสดงล็อก
        self.log_table = QTableWidget()
        self.log_table.setColumnCount(4)
        self.log_table.setHorizontalHeaderLabels(["เวลา", "ประเภทเหตุการณ์", "สถานะ HDMI", "รายละเอียด"])
        self.log_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.log_layout.addWidget(self.log_table)
    
    def update_current_time(self):
        """อัปเดตเวลาปัจจุบันในอินเตอร์เฟซ"""
        try:
            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.time_label.setText(f"เวลาปัจจุบัน: {current_time}")
        except KeyboardInterrupt:
            # จัดการกับ KeyboardInterrupt ที่อาจเกิดขึ้น
            print("ได้รับ KeyboardInterrupt ในการอัปเดตเวลา")
        except Exception as e:
            print(f"เกิดข้อผิดพลาดในการอัปเดตเวลา: {str(e)}")
    
    def update_status(self, status):
        """อัปเดตการแสดงผลสถานะ"""
        try:
            # อัปเดตสถานะ HDMI
            if status['hdmi_connected']:
                self.hdmi_status_label.setText(f"สถานะการเชื่อมต่อ HDMI: {STATUS_ICONS['normal']} เชื่อมต่อแล้ว")
                self.hdmi_status_label.setStyleSheet("color: green;")
            else:
                self.hdmi_status_label.setText(f"สถานะการเชื่อมต่อ HDMI: {STATUS_ICONS['error']} ไม่ได้เชื่อมต่อ")
                self.hdmi_status_label.setStyleSheet("color: red;")
            
            if status['hdmi_active']:
                self.hdmi_active_label.setText(f"สถานะการทำงาน HDMI: {STATUS_ICONS['normal']} กำลังทำงาน (Active)")
                self.hdmi_active_label.setStyleSheet("color: green;")
            else:
                self.hdmi_active_label.setText(f"สถานะการทำงาน HDMI: {STATUS_ICONS['warning']} ไม่ได้ทำงาน (Inactive)")
                self.hdmi_active_label.setStyleSheet("color: orange;")
            
            # อัปเดตข้อมูลการ์ดจอ
            gpu_info = ""
            if 'gpu_error' in status:
                gpu_info = f"{STATUS_ICONS['warning']} พบข้อผิดพลาดในการตรวจสอบการ์ดจอ: {status['gpu_error']}\n\n"
            
            for gpu_id, gpu_data in status['gpu_status'].items():
                gpu_info += f"การ์ดจอ: {gpu_data['name']}\n"
                gpu_info += f"สถานะ: {gpu_data['status']}\n"
                
                # แสดงสถานะไดรเวอร์
                driver_status = gpu_data.get('driver_status', 'unknown')
                if driver_status == 'normal':
                    gpu_info += f"ไดรเวอร์: {STATUS_ICONS['normal']} ทำงานปกติ\n"
                elif driver_status == 'warning':
                    gpu_info += f"ไดรเวอร์: {STATUS_ICONS['warning']} ทำงานได้แต่อาจมีปัญหา\n"
                elif driver_status == 'error':
                    gpu_info += f"ไดรเวอร์: {STATUS_ICONS['error']} มีปัญหา (รหัส {gpu_data['error_code']})\n"
                    gpu_info += f"รายละเอียด: {gpu_data['error_description']}\n"
                else:
                    gpu_info += f"ไดรเวอร์: {STATUS_ICONS['unknown']} ไม่ทราบสถานะ\n"
                
                # แสดงข้อมูลเวอร์ชันและวันที่
                gpu_info += f"เวอร์ชันไดรเวอร์: {gpu_data['driver_version']}\n"
                if 'driver_date' in gpu_data and gpu_data['driver_date'] != "Unknown":
                    gpu_info += f"วันที่ไดรเวอร์: {gpu_data['driver_date']}\n"
                gpu_info += "\n"
            
            self.gpu_status_text.setText(gpu_info)
            
            # อัปเดตข้อมูลจอภาพทั้งหมด
            monitor_info = ""
            if 'wmi_error' in status:
                monitor_info = f"{STATUS_ICONS['warning']} พบข้อผิดพลาดในการตรวจสอบจอภาพ: {status['wmi_error']}\n\n"
            
            for i, monitor in enumerate(status['hdmi_devices']):
                conn_type_name = self._get_connection_type_name(monitor['connection_type'])
                active_status = "กำลังทำงาน" if monitor['active'] else "ไม่ได้ทำงาน"
                
                monitor_info += f"จอภาพ #{i+1}:\n"
                monitor_info += f"ชนิดการเชื่อมต่อ: {conn_type_name}"
                if monitor['connection_type'] != -1:
                    monitor_info += f" (รหัส {monitor['connection_type']})"
                monitor_info += "\n"
                monitor_info += f"สถานะ: {active_status}\n"
                
                if monitor.get('status') != 'normal':
                    monitor_info += f"ข้อผิดพลาด: {monitor.get('status')}\n"
                
                monitor_info += f"InstanceName: {monitor['instance_name']}\n\n"
            
            if not monitor_info:
                monitor_info = "ไม่พบจอภาพที่เชื่อมต่อ"
            
            self.monitor_info_text.setText(monitor_info)
            
            # อัปเดตข้อความใน System Tray
            tray_tooltip = f"{APP_TITLE}\n"
            if status['hdmi_connected']:
                tray_tooltip += "HDMI: เชื่อมต่อแล้ว"
                if status['hdmi_active']:
                    tray_tooltip += ", กำลังทำงาน"
                else:
                    tray_tooltip += ", ไม่ได้ทำงาน"
            else:
                tray_tooltip += "HDMI: ไม่ได้เชื่อมต่อ"
            
            if 'driver_check_time' in status:
                tray_tooltip += f"\nตรวจสอบไดรเวอร์ล่าสุด: {status['driver_check_time'].split('T')[1].split('.')[0]}"
            
            self.tray_icon.setToolTip(tray_tooltip)
            
        except Exception as e:
            error_msg = f"เกิดข้อผิดพลาดในการอัปเดตสถานะ: {str(e)}"
            print(error_msg)
            self.gpu_status_text.setText(error_msg)
            self.monitor_info_text.setText(error_msg)
    
    def _get_connection_type_name(self, type_code):
        """แปลงรหัสประเภทการเชื่อมต่อเป็นชื่อ"""
        return CONNECTION_TYPES.get(type_code, f"ไม่รู้จัก ({type_code})")
    
    def add_log_entry(self, log_entry):
        """เพิ่มรายการล็อกใหม่และอัปเดตการแสดงผลหากจำเป็น"""
        self.log_manager.add_log_entry(log_entry)
        
        # อัปเดตหน้าแสดงล็อกหากเป็นวันนี้
        log_date = QDate.fromString(log_entry['timestamp'].split('T')[0], "yyyy-MM-dd")
        if log_date == self.date_selector.date():
            self.update_log_view(log_date)
    
    def date_changed(self, date):
        """จัดการเมื่อมีการเลือกวันที่ใหม่"""
        self.update_log_view(date)
    
    def update_log_view(self, date):
        """อัปเดตการแสดงรายการล็อกตามวันที่"""
        logs = self.log_manager.get_logs_by_date(date)
        
        self.log_table.setRowCount(0)  # ล้างข้อมูลเดิม
        
        for log in logs:
            row_position = self.log_table.rowCount()
            self.log_table.insertRow(row_position)
            
            # เวลา (ช่อง 0)
            timestamp = log['timestamp'].replace('T', ' ').split('.')[0]  # ตัดหน่วยไมโครวินาที
            self.log_table.setItem(row_position, 0, QTableWidgetItem(timestamp))
            
            # ประเภทเหตุการณ์ (ช่อง 1)
            event_type = log['event_type']
            event_cell = QTableWidgetItem(self._format_event_type(event_type))
            if "ERROR" in event_type:
                event_cell.setForeground(QColor("red"))
            elif event_type == "HDMI_DISCONNECTED":
                event_cell.setForeground(QColor("red"))
            elif event_type == "HDMI_INACTIVE":
                event_cell.setForeground(QColor("orange"))
            self.log_table.setItem(row_position, 1, event_cell)
            
            # สถานะ HDMI (ช่อง 2)
            hdmi_status = "เชื่อมต่อแล้ว" if log['status']['hdmi_connected'] else "ไม่ได้เชื่อมต่อ"
            hdmi_active = "กำลังทำงาน" if log['status']['hdmi_active'] else "ไม่ได้ทำงาน"
            hdmi_cell = QTableWidgetItem(f"{hdmi_status}, {hdmi_active}")
            self.log_table.setItem(row_position, 2, hdmi_cell)
            
            # รายละเอียด (ช่อง 3)
            details = self._format_log_details(log)
            self.log_table.setItem(row_position, 3, QTableWidgetItem(details))
    
    def _format_event_type(self, event_type):
        """แปลงรหัสเหตุการณ์เป็นข้อความที่อ่านง่าย"""
        event_types = {
            "NORMAL": "ปกติ",
            "HDMI_DISCONNECTED": "HDMI ถูกถอด",
            "HDMI_INACTIVE": "HDMI ไม่ทำงาน"
        }
        
        if event_type.startswith("GPU_ERROR_CODE_"):
            error_code = event_type.split("_")[-1]
            return f"การ์ดจอมีปัญหา (รหัส {error_code})"
            
        return event_types.get(event_type, event_type)
    
    def _format_log_details(self, log):
        """สร้างข้อความรายละเอียดจากล็อก"""
        details = ""
        
        # รายละเอียดการ์ดจอ
        for gpu_id, gpu_data in log['status']['gpu_status'].items():
            if gpu_data['error_code'] != 0:
                details += f"การ์ดจอ {gpu_data['name']} มีปัญหา: {gpu_data['error_description']}\n"
        
        # รายละเอียดจอภาพ
        hdmi_devices = log['status']['hdmi_devices']
        if hdmi_devices:
            hdmi_count = sum(1 for dev in hdmi_devices if dev['connection_type'] == CONNECTION_TYPE_HDMI)
            details += f"พบจอ HDMI {hdmi_count} จอ, "
            active_count = sum(1 for dev in hdmi_devices if dev['connection_type'] == CONNECTION_TYPE_HDMI and dev['active'])
            details += f"ทำงาน {active_count} จอ"
        else:
            details += "ไม่พบจอภาพที่เชื่อมต่อ"
            
        return details
    
    def refresh_status(self):
        """รีเฟรชสถานะทันที"""
        try:
            status = self.hdmi_monitor.check_hdmi_status()
            self.update_status(status)
            QMessageBox.information(self, "รีเฟรชสถานะ", "รีเฟรชสถานะเรียบร้อยแล้ว")
        except Exception as e:
            QMessageBox.warning(self, "เกิดข้อผิดพลาด", f"ไม่สามารถรีเฟรชสถานะได้: {str(e)}")
    
    def close_application(self):
        """ปิดแอปพลิเคชัน"""
        print("กำลังปิดแอปพลิเคชัน...")
        # หยุดการทำงานของเธรด
        self.hdmi_monitor.stop()
        self.hdmi_monitor.wait()
        self.system_monitor.stop()
        self.system_monitor.wait()
        
        # ปิด QApplication และออกจากโปรแกรม
        QApplication.quit()
        sys.exit(0)
    
    def update_system_status(self, status):
        """อัปเดตการแสดงผลสถานะระบบ"""
        # CPU
        cpu_percent = status['cpu'].get('percent', 0)
        self.cpu_percent_label.setText(f"{cpu_percent:.1f}%")
        self.cpu_progress.setValue(int(cpu_percent))
        
        # Set color based on CPU usage
        if cpu_percent < 60:
            self.cpu_progress.setStyleSheet("QProgressBar::chunk { background-color: green; }")
        elif cpu_percent < 85:
            self.cpu_progress.setStyleSheet("QProgressBar::chunk { background-color: orange; }")
        else:
            self.cpu_progress.setStyleSheet("QProgressBar::chunk { background-color: red; }")
        
        cpu_cores = status['cpu'].get('cores', 0)
        cpu_freq = status['cpu'].get('frequency', 0)
        self.cpu_info_label.setText(f"จำนวน Core: {cpu_cores} | ความเร็ว: {cpu_freq:.0f} MHz")
        
        # RAM
        ram_percent = status['memory'].get('percent', 0)
        self.ram_percent_label.setText(f"{ram_percent:.1f}%")
        self.ram_progress.setValue(int(ram_percent))
        
        # Set color based on RAM usage
        if ram_percent < 60:
            self.ram_progress.setStyleSheet("QProgressBar::chunk { background-color: green; }")
        elif ram_percent < 85:
            self.ram_progress.setStyleSheet("QProgressBar::chunk { background-color: orange; }")
        else:
            self.ram_progress.setStyleSheet("QProgressBar::chunk { background-color: red; }")
        
        ram_total = status['memory'].get('total', 0) / (1024 ** 3)  # Convert to GB
        ram_used = status['memory'].get('used', 0) / (1024 ** 3)  # Convert to GB
        self.ram_info_label.setText(f"ใช้งาน: {ram_used:.2f} GB / ทั้งหมด: {ram_total:.2f} GB")
        
        # อุณหภูมิ
        temp_info = "อุณหภูมิระบบ:\n"
        if 'temperatures' in status:
            if 'error' in status['temperatures']:
                temp_info += f"ไม่สามารถอ่านอุณหภูมิได้: {status['temperatures']['error']}\n"
            else:
                if 'cpu' in status['temperatures']:
                    temp_info += f"CPU: {status['temperatures']['cpu']:.1f}°C\n"
                if 'mainboard' in status['temperatures']:
                    temp_info += f"Mainboard: {status['temperatures']['mainboard']:.1f}°C\n"
        else:
            temp_info += "ไม่พบข้อมูลอุณหภูมิ\n"
        
        # GPU
        gpu_info = ""
        for i, gpu in enumerate(status['gpu']):
            if 'error' in gpu:
                gpu_info = f"ข้อผิดพลาด: {gpu['error']}"
                continue
                
            if 'load' in gpu:
                gpu_info += f"GPU #{i+1}: {gpu.get('name', 'ไม่ทราบชื่อ')}\n"
                gpu_info += f"การใช้งาน: {gpu.get('load', 0):.1f}%\n"
                gpu_info += f"หน่วยความจำ: {gpu.get('memory_used', 0):.0f} MB / {gpu.get('memory_total', 0):.0f} MB ({gpu.get('memory_percent', 0):.1f}%)\n"
                gpu_info += f"อุณหภูมิ: {gpu.get('temperature', 0):.1f}°C\n\n"
            else:
                gpu_info += f"GPU #{i+1}: {gpu.get('name', 'ไม่ทราบชื่อ')}\n"
                gpu_info += f"ไดรเวอร์: {gpu.get('driver_version', 'ไม่ทราบ')}\n"
                gpu_info += f"สถานะ: {gpu.get('driver_status', 'ไม่ทราบ')}\n"
                if 'status_message' in gpu:
                    gpu_info += f"รายละเอียด: {gpu.get('status_message')}\n\n"
                else:
                    gpu_info += "\n"
        
        if not gpu_info:
            gpu_info = "ไม่พบข้อมูล GPU หรือไม่สามารถอ่านข้อมูลได้"
            
        # รวมข้อมูลอุณหภูมิและ GPU
        self.gpu_info_text.setText(temp_info + "\n" + gpu_info)
    
    def closeEvent(self, event):
        """จัดการเมื่อปิดแอปพลิเคชัน (เมื่อคลิกปุ่ม X ที่มุมขวาบน)"""
        self.close_application()
        # ไม่จำเป็นต้องเรียก event.accept() เพราะโปรแกรมจะถูกปิดโดย sys.exit() ใน close_application


def check_cec_info():
    """ตรวจสอบข้อมูล CEC"""
    try:
        # แก้ไขการใช้งาน pycec
        print("ฟังก์ชันการตรวจสอบข้อมูล CEC ถูกปิดการใช้งานชั่วคราว")
        # ตัวอย่างข้อมูลที่จะแสดง (แทนการใช้ pycec)
        sample_devices = [
            {"name": "Samsung TV", "vendor": "Samsung", "model": "UE55NU7400", "serial": "XYZ123456789"},
            {"name": "LG Display", "vendor": "LG", "model": "27GL850", "serial": "ABC987654321"}
        ]
        
        for device in sample_devices:
            print(f"ชื่อผลิตภัณฑ์: {device['name']}")
            print(f"แบรนด์: {device['vendor']}")
            print(f"หมายเลขรุ่น: {device['model']}")
            print(f"หมายเลขซีเรียล: {device['serial']}")
        
        print("\nหมายเหตุ: ข้อมูลนี้เป็นเพียงตัวอย่าง ไม่ใช่ข้อมูลจริงจากอุปกรณ์")
    except Exception as e:
        print(f"Error checking CEC info: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # ตรวจสอบว่า System Tray สามารถใช้งานได้หรือไม่
    if not QSystemTrayIcon.isSystemTrayAvailable():
        QMessageBox.critical(None, "System Tray ไม่พร้อมใช้งาน",
                           "ระบบของคุณไม่รองรับ System Tray ซึ่งจำเป็นสำหรับแอปพลิเคชันนี้")
        sys.exit(1)
    
    # ตั้งค่าให้แอปพลิเคชันไม่ปิดเมื่อปิดหน้าต่างสุดท้าย (เพื่อให้ทำงานใน System Tray ได้)
    QApplication.setQuitOnLastWindowClosed(False)
    
    # จัดการสัญญาณเพื่อให้โปรแกรมปิดอย่างสวยงาม
    import signal
    
    # จัดการสัญญาณ Ctrl+C
    def signal_handler(sig, frame):
        print("ได้รับสัญญาณให้ปิดโปรแกรม (Ctrl+C)")
        if hasattr(window, 'close_application'):
            window.close_application()
        else:
            sys.exit(0)
    
    # ลงทะเบียนตัวจัดการสัญญาณ
    signal.signal(signal.SIGINT, signal_handler)
    
    window = MainWindow()
    window.show()
    
    try:
        sys.exit(app.exec_())
    except KeyboardInterrupt:
        print("ได้รับ KeyboardInterrupt ในเมธอดหลัก")
        window.close_application()
    except Exception as e:
        print(f"เกิดข้อผิดพลาดที่ไม่คาดคิด: {str(e)}")
        sys.exit(1)