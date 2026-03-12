import sys
import os
import shutil
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel, QMessageBox, QDialog, QFormLayout, QLineEdit, QDialogButtonBox
from PyQt5.QtCore import Qt

from spul.spul_app import SledAnalyzerApp
import shared.global_data as global_data

class ReportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Rapor Bilgileri")
        self.resize(300, 150)
        
        self.layout = QFormLayout(self)
        
        # Eğer daha önceden veriler girildiyse, onları varsayılan olarak göster
        current_test_no = global_data.config["TEST_NO"] if global_data.config["TEST_NO"] else "2026/096"
        current_test_date = global_data.config["TEST_DATE"] if global_data.config["TEST_DATE"] else "08.03.2026"
        current_project = global_data.config["PROJECT"] if global_data.config["PROJECT"] else "V227"
        
        self.txt_test_no = QLineEdit(current_test_no)
        self.txt_test_date = QLineEdit(current_test_date)
        self.txt_project = QLineEdit(current_project)
        
        self.layout.addRow("Test No:", self.txt_test_no)
        self.layout.addRow("Test Date:", self.txt_test_date)
        self.layout.addRow("Project:", self.txt_project)
        
        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        
        self.layout.addWidget(self.buttons)

    def get_data(self):
        return {
            "TEST_NO": self.txt_test_no.text().strip(),
            "TEST_DATE": self.txt_test_date.text().strip(),
            "PROJECT": self.txt_project.text().strip()
        }

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
            global_data.config["TEST_NO"] = data["TEST_NO"]
            global_data.config["TEST_DATE"] = data["TEST_DATE"]
            global_data.config["PROJECT"] = data["PROJECT"]
            QMessageBox.information(self, "Başarılı", "Genel bilgiler kaydedildi.")

    def open_spul_app(self):
        if not global_data.config["TEST_NO"] or not global_data.config["TEST_DATE"] or not global_data.config["PROJECT"]:
            QMessageBox.warning(self, "Uyarı", "Lütfen önce genel bilgileri girin!")
            return
            
        self.hide()
        self.spul_window = SledAnalyzerApp(main_window=self)
        self.spul_window.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec_())
