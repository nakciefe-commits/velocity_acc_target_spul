import sys
import os
import shutil
import subprocess
import importlib


def _ensure_dependencies():
    required_packages = {
        "PyQt5": "PyQt5",
        "pandas": "pandas",
        "numpy": "numpy",
        "matplotlib": "matplotlib",
        "docxtpl": "docxtpl",
        "openpyxl": "openpyxl",
        "docx": "python-docx",
        "xlrd": "xlrd"
    }

    for module_name, pip_name in required_packages.items():
        try:
            importlib.import_module(module_name)
        except ImportError:
            print(f"Eksik paket bulundu: {pip_name}. Kurulum başlatılıyor...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])
            except Exception as exc:
                print(f"{pip_name} kurulamadı: {exc}")
                print("Lütfen şu komutu çalıştırın:")
                print(f"  {sys.executable} -m pip install {pip_name}")
                raise


_ensure_dependencies()

from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel, QMessageBox, QDialog, QFormLayout, QLineEdit, QDialogButtonBox, QComboBox, QScrollArea, QFrame
from PyQt5.QtCore import Qt

from spul.spul_app import SledAnalyzerApp
import shared.global_data as global_data
import kapak.kapak_app as kapak_app
from photos.photo_report_app import PhotoReportApp

class ReportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Rapor Bilgileri")
        self.resize(800, 900)  # Kaydırma çubuğunu kaldırmak için boyut epey büyütüldü
        
        main_layout = QVBoxLayout(self)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        self.form_layout = QFormLayout(scroll_content)
        
        # General Info
        self.inputs = {}
        fields = ["TEST_NAME", "REPORT_NO", "TEST_ID", "WO_NO", "TEST_NO", "TEST_DATE", "OEM", "PROGRAM", "PURPOSE"]
        
        display_names = {
            "TEST_NAME": "Test Name",
            "REPORT_NO": "Report No",
            "TEST_ID": "Test Id",
            "WO_NO": "Wo No",
            "TEST_NO": "Test No",
            "TEST_DATE": "Test Date",
            "OEM": "Oem",
            "PROGRAM": "Program",
            "PURPOSE": "PURPOSE"
        }
        
        for field in fields:
            val = global_data.config.get(field)
            if val is None: val = ""
            if field == "TEST_NO" and not val: val = "2026/096"
            if field == "TEST_DATE" and not val: val = "08.03.2026"
            
            le = QLineEdit(str(val))
            self.inputs[field] = le
            self.form_layout.addRow(f"{display_names[field]}:", le)
            
        # Seat Count
        self.cb_seat_count = QComboBox()
        self.cb_seat_count.addItems(["1", "2", "3", "4", "5"])
        current_seat = str(global_data.config.get("SEAT_COUNT", 1))
        self.cb_seat_count.setCurrentText(current_seat)
        self.cb_seat_count.currentIndexChanged.connect(self.update_dynamic_fields)
        self.form_layout.addRow("Koltuk Sayısı:", self.cb_seat_count)
        
        # Dynamic Fields Container
        self.dynamic_frame = QFrame()
        self.dynamic_layout = QFormLayout(self.dynamic_frame)
        self.form_layout.addRow(self.dynamic_frame)
        
        self.dynamic_inputs = {"SMP_ID": [], "TEST_SAMPLE": []}
        self.update_dynamic_fields()
        
        scroll.setWidget(scroll_content)
        main_layout.addWidget(scroll)
        
        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        main_layout.addWidget(self.buttons)

    def update_dynamic_fields(self):
        # Clear existing dynamic fields
        for i in reversed(range(self.dynamic_layout.count())): 
            widget = self.dynamic_layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()
                
        self.dynamic_inputs = {"SMP_ID": [], "TEST_SAMPLE": []}
        seat_count = int(self.cb_seat_count.currentText())
        
        for i in range(seat_count):
            # Load existing if available, else empty
            smp_val = global_data.config["SMP_ID"][i] if global_data.config["SMP_ID"][i] else ""
            ts_val = global_data.config["TEST_SAMPLE"][i] if global_data.config["TEST_SAMPLE"][i] else ""
            
            le_smp = QLineEdit(str(smp_val))
            le_ts = QLineEdit(str(ts_val))
            
            self.dynamic_inputs["SMP_ID"].append(le_smp)
            self.dynamic_inputs["TEST_SAMPLE"].append(le_ts)
            
            self.dynamic_layout.addRow(QLabel(f"--- Koltuk {i+1} ---"))
            self.dynamic_layout.addRow(f"Smp Id {i+1}:", le_smp)
            self.dynamic_layout.addRow(f"Test Sample {i+1}:", le_ts)

    def get_data(self):
        data = {field: le.text().strip() for field, le in self.inputs.items()}
        data["SEAT_COUNT"] = int(self.cb_seat_count.currentText())
        
        smp_ids = ["" for _ in range(5)]
        test_samples = ["" for _ in range(5)]
        
        for i in range(data["SEAT_COUNT"]):
            smp_ids[i] = self.dynamic_inputs["SMP_ID"][i].text().strip()
            test_samples[i] = self.dynamic_inputs["TEST_SAMPLE"][i].text().strip()
            
        data["SMP_ID"] = smp_ids
        data["TEST_SAMPLE"] = test_samples
        
        return data

class MainApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ana Uygulama")
        self.resize(400, 300)
        
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        lbl_title = QLabel("Adient Sled Test Merkezi")
        lbl_title.setAlignment(Qt.AlignCenter)
        lbl_title.setStyleSheet("font-size: 20px; font-weight: bold; margin-bottom: 20px;")
        layout.addWidget(lbl_title)
        
        btn_global_info = QPushButton("Genel Bilgileri Gir")
        btn_global_info.setStyleSheet("font-size: 16px; padding: 15px; background-color: #2196F3; color: white; font-weight: bold;")
        btn_global_info.clicked.connect(self.open_global_info)
        layout.addWidget(btn_global_info)
        
        btn_kapak = QPushButton("Kapak Oluştur")
        btn_kapak.setStyleSheet("font-size: 16px; padding: 15px; background-color: #FF9800; color: white; font-weight: bold;")
        btn_kapak.clicked.connect(self.create_kapak)
        layout.addWidget(btn_kapak)
        
        btn_photo_report = QPushButton("Photo Report Modülünü Aç")
        btn_photo_report.setStyleSheet("font-size: 16px; padding: 15px; background-color: #673AB7; color: white; font-weight: bold;")
        btn_photo_report.clicked.connect(self.open_photo_report_app)
        layout.addWidget(btn_photo_report)

        btn_spul = QPushButton("Spul Uygulamasını Aç")
        btn_spul.setStyleSheet("font-size: 16px; padding: 15px; background-color: #4CAF50; color: white; font-weight: bold;")
        btn_spul.clicked.connect(self.open_spul_app)
        layout.addWidget(btn_spul)
        
        layout.addStretch()
        
        # Tempfiles klasörünü kontrol et
        self.check_tempfiles()

    def check_tempfiles(self):
        tempfiles_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tempfiles")
        if os.path.exists(tempfiles_dir):
            files = os.listdir(tempfiles_dir)
            if files: # Klasör boş değilse
                reply = QMessageBox.question(self, 'Taslakları Sil', 
                                             'Tempfiles klasöründe kayıtlı taslaklar var, silinsin mi?',
                                             QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply == QMessageBox.Yes:
                    for filename in files:
                        file_path = os.path.join(tempfiles_dir, filename)
                        try:
                            if os.path.isfile(file_path):
                                os.unlink(file_path)
                            elif os.path.isdir(file_path):
                                shutil.rmtree(file_path)
                        except Exception as e:
                            print(f"Hata: {file_path} silinemedi. {e}")

    def open_global_info(self):
        dialog = ReportDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            for key, val in data.items():
                global_data.config[key] = val
            # Backward compatibility
            global_data.config["PROJECT"] = global_data.config["PROGRAM"]
            QMessageBox.information(self, "Başarılı", "Genel bilgiler kaydedildi.")

    def create_kapak(self):
        if not global_data.config["TEST_NO"] or not global_data.config["TEST_DATE"]:
            QMessageBox.warning(self, "Uyarı", "Lütfen önce genel bilgileri eksiksiz girin!")
            return
        
        kapak_app.generate_cover_report(self)

    def open_photo_report_app(self):
        self.hide()
        self.photo_report_window = PhotoReportApp(main_window=self)
        self.photo_report_window.show()

    def open_spul_app(self):
        if not global_data.config["TEST_NO"] or not global_data.config["TEST_DATE"] or not global_data.config["WO_NO"]:
            QMessageBox.warning(self, "Uyarı", "Lütfen önce genel bilgileri eksiksiz girin!")
            return
            
        self.hide()
        self.spul_window = SledAnalyzerApp(main_window=self)
        self.spul_window.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec_())
