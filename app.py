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
        "xlrd": "xlrd",
        "PIL": "Pillow",
        "docxcompose": "docxcompose",
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

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QMessageBox, QDialog, QFormLayout, QLineEdit,
    QDialogButtonBox, QComboBox, QScrollArea, QFrame, QGroupBox,
    QFileDialog, QSizePolicy,
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

from spul.spul_app import SledAnalyzerApp
import shared.global_data as global_data
import kapak.kapak_app as kapak_app
from photos.photo_report_app import PhotoReportApp
from eva.eva_app import EvaApp


# ---------------------------------------------------------------------------
#  Stylesheet
# ---------------------------------------------------------------------------
DIALOG_STYLE = """
    QDialog {
        background-color: #FAFAFA;
    }
    QGroupBox {
        font-size: 13px;
        font-weight: bold;
        color: #1565C0;
        border: 1px solid #BBDEFB;
        border-radius: 6px;
        margin-top: 14px;
        padding: 18px 12px 10px 12px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 14px;
        padding: 0 6px;
        background-color: #FAFAFA;
    }
    QLineEdit {
        padding: 6px 8px;
        border: 1px solid #BDBDBD;
        border-radius: 4px;
        font-size: 13px;
        background: white;
    }
    QLineEdit:focus {
        border: 1.5px solid #1976D2;
    }
    QComboBox {
        padding: 6px 8px;
        border: 1px solid #BDBDBD;
        border-radius: 4px;
        font-size: 13px;
        background: white;
    }
    QLabel {
        font-size: 13px;
        color: #424242;
    }
    QPushButton#btnOk {
        background-color: #1976D2;
        color: white;
        font-weight: bold;
        font-size: 14px;
        padding: 10px 32px;
        border-radius: 5px;
        border: none;
    }
    QPushButton#btnOk:hover {
        background-color: #1565C0;
    }
    QPushButton#btnCancel {
        background-color: #9E9E9E;
        color: white;
        font-size: 14px;
        padding: 10px 32px;
        border-radius: 5px;
        border: none;
    }
    QPushButton#btnCancel:hover {
        background-color: #757575;
    }
"""

MAIN_STYLE = """
    QMainWindow {
        background-color: #FAFAFA;
    }
"""

MAIN_BTN_STYLE = "font-size: 15px; padding: 14px; border-radius: 6px; border: none; font-weight: bold; color: white; background-color: {color};"


# ---------------------------------------------------------------------------
#  ReportDialog  –  Genel Bilgiler formu
# ---------------------------------------------------------------------------
class ReportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Genel Bilgiler")
        self.setStyleSheet(DIALOG_STYLE)
        self.resize(620, 740)

        root = QVBoxLayout(self)
        root.setSpacing(6)

        # Başlık
        header = QLabel("Rapor Genel Bilgileri")
        header.setAlignment(Qt.AlignCenter)
        header.setStyleSheet("font-size: 18px; font-weight: bold; color: #0D47A1; margin: 4px 0 8px 0;")
        root.addWidget(header)

        # Scroll
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; }")
        scroll_content = QWidget()
        content_layout = QVBoxLayout(scroll_content)
        content_layout.setSpacing(4)

        # ---- Test Bilgileri grubu ----
        grp_test = QGroupBox("Test Bilgileri")
        form_test = QFormLayout()
        form_test.setLabelAlignment(Qt.AlignRight)
        form_test.setHorizontalSpacing(16)
        form_test.setVerticalSpacing(8)

        self.inputs = {}
        test_fields = [
            ("TEST_NAME",  "Test Adı"),
            ("REPORT_NO",  "Rapor No"),
            ("TEST_ID",    "Test ID"),
            ("WO_NO",      "WO No"),
            ("TEST_NO",    "Test No"),
            ("TEST_DATE",  "Test Tarihi"),
        ]
        for key, label in test_fields:
            le = QLineEdit()
            le.setMinimumWidth(280)
            val = global_data.config.get(key)
            if val:
                le.setPlaceholderText(str(val))
            self.inputs[key] = le
            form_test.addRow(f"{label}:", le)

        grp_test.setLayout(form_test)
        content_layout.addWidget(grp_test)

        # ---- Müşteri / Proje grubu ----
        grp_project = QGroupBox("Müşteri ve Proje")
        form_project = QFormLayout()
        form_project.setLabelAlignment(Qt.AlignRight)
        form_project.setHorizontalSpacing(16)
        form_project.setVerticalSpacing(8)

        project_fields = [
            ("OEM",      "OEM"),
            ("PROGRAM",  "Program"),
            ("PURPOSE",  "Amaç"),
        ]
        for key, label in project_fields:
            le = QLineEdit()
            le.setMinimumWidth(280)
            val = global_data.config.get(key)
            if val:
                le.setPlaceholderText(str(val))
            self.inputs[key] = le
            form_project.addRow(f"{label}:", le)

        grp_project.setLayout(form_project)
        content_layout.addWidget(grp_project)

        # ---- EVA / Dummy grubu ----
        grp_eva = QGroupBox("EVA Bilgileri")
        form_eva = QFormLayout()
        form_eva.setLabelAlignment(Qt.AlignRight)
        form_eva.setHorizontalSpacing(16)
        form_eva.setVerticalSpacing(8)

        eva_fields = [
            ("DUMMY_PCT",  "Dummy %"),
            ("SENSOR",   "Sensor"),
        ]
        for key, label in eva_fields:
            le = QLineEdit()
            le.setMinimumWidth(280)
            val = global_data.config.get(key)
            if val:
                le.setPlaceholderText(str(val))
            self.inputs[key] = le
            form_eva.addRow(f"{label}:", le)

        grp_eva.setLayout(form_eva)
        content_layout.addWidget(grp_eva)

        # ---- Koltuk Bilgileri grubu ----
        grp_seats = QGroupBox("Koltuk Bilgileri")
        seats_layout = QVBoxLayout()
        seats_layout.setSpacing(8)

        seat_top = QHBoxLayout()
        seat_top.addWidget(QLabel("Koltuk Sayısı:"))
        self.cb_seat_count = QComboBox()
        self.cb_seat_count.addItems(["1", "2", "3", "4", "5"])
        current_seat = str(global_data.config.get("SEAT_COUNT", 1))
        self.cb_seat_count.setCurrentText(current_seat)
        self.cb_seat_count.currentIndexChanged.connect(self._rebuild_seat_fields)
        seat_top.addWidget(self.cb_seat_count)
        seat_top.addStretch()
        seats_layout.addLayout(seat_top)

        self.seat_frame = QFrame()
        self.seat_form = QFormLayout(self.seat_frame)
        self.seat_form.setLabelAlignment(Qt.AlignRight)
        self.seat_form.setHorizontalSpacing(16)
        self.seat_form.setVerticalSpacing(8)
        seats_layout.addWidget(self.seat_frame)

        grp_seats.setLayout(seats_layout)
        content_layout.addWidget(grp_seats)

        content_layout.addStretch()
        scroll.setWidget(scroll_content)
        root.addWidget(scroll)

        # ---- Butonlar ----
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        btn_ok = QPushButton("Kaydet")
        btn_ok.setObjectName("btnOk")
        btn_ok.clicked.connect(self.accept)
        btn_row.addWidget(btn_ok)

        btn_cancel = QPushButton("İptal")
        btn_cancel.setObjectName("btnCancel")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)

        btn_row.addStretch()
        root.addLayout(btn_row)

        # Koltuk alanlarını oluştur
        self.dynamic_inputs = {"SMP_ID": [], "TEST_SAMPLE": []}
        self._rebuild_seat_fields()

    # ---- dynamic seat fields ----
    def _rebuild_seat_fields(self):
        # Temizle
        while self.seat_form.count():
            item = self.seat_form.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        self.dynamic_inputs = {"SMP_ID": [], "TEST_SAMPLE": []}
        seat_count = int(self.cb_seat_count.currentText())

        for i in range(seat_count):
            smp_val = global_data.config["SMP_ID"][i] if global_data.config["SMP_ID"][i] else ""
            ts_val = global_data.config["TEST_SAMPLE"][i] if global_data.config["TEST_SAMPLE"][i] else ""

            lbl = QLabel(f"  Koltuk {i + 1}")
            lbl.setStyleSheet("font-weight: bold; color: #1565C0; font-size: 12px; margin-top: 4px;")
            self.seat_form.addRow(lbl)

            le_smp = QLineEdit()
            le_ts = QLineEdit()
            if smp_val:
                le_smp.setPlaceholderText(str(smp_val))
            if ts_val:
                le_ts.setPlaceholderText(str(ts_val))
            self.dynamic_inputs["SMP_ID"].append(le_smp)
            self.dynamic_inputs["TEST_SAMPLE"].append(le_ts)

            self.seat_form.addRow("SMP ID:", le_smp)
            self.seat_form.addRow("Test Sample:", le_ts)

    def _field_value(self, le):
        """Kullanıcı yazdıysa onu al, yoksa placeholder'ı kullan."""
        text = le.text().strip()
        if text:
            return text
        return le.placeholderText().strip()

    def get_data(self):
        data = {field: self._field_value(le) for field, le in self.inputs.items()}
        data["SEAT_COUNT"] = int(self.cb_seat_count.currentText())

        smp_ids = ["" for _ in range(5)]
        test_samples = ["" for _ in range(5)]

        for i in range(data["SEAT_COUNT"]):
            smp_ids[i] = self._field_value(self.dynamic_inputs["SMP_ID"][i])
            test_samples[i] = self._field_value(self.dynamic_inputs["TEST_SAMPLE"][i])

        data["SMP_ID"] = smp_ids
        data["TEST_SAMPLE"] = test_samples
        return data


# ---------------------------------------------------------------------------
#  MainApp
# ---------------------------------------------------------------------------
class MainApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Adient Sled Test Merkezi")
        self.setStyleSheet(MAIN_STYLE)
        self.resize(460, 420)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setSpacing(10)
        layout.setContentsMargins(24, 18, 24, 18)

        # Başlık
        lbl_title = QLabel("Adient Sled Test Merkezi")
        lbl_title.setAlignment(Qt.AlignCenter)
        lbl_title.setStyleSheet("font-size: 22px; font-weight: bold; color: #0D47A1; margin-bottom: 12px;")
        layout.addWidget(lbl_title)

        # Durum etiketi (yüklü config bilgisi)
        self.lbl_status = QLabel()
        self.lbl_status.setAlignment(Qt.AlignCenter)
        self.lbl_status.setStyleSheet("font-size: 12px; color: #757575; margin-bottom: 4px;")
        layout.addWidget(self.lbl_status)

        # Butonlar
        buttons = [
            ("Genel Bilgileri Gir",         "#2196F3", self.open_global_info),
            ("Kapak Oluştur",               "#FF9800", self.create_kapak),
            ("Photo Report Modülünü Aç",    "#673AB7", self.open_photo_report_app),
            ("Spul Uygulamasını Aç",        "#4CAF50", self.open_spul_app),
            ("EVA Modülünü Aç",             "#E91E63", self.open_eva_app),
        ]

        for text, color, handler in buttons:
            btn = QPushButton(text)
            btn.setStyleSheet(MAIN_BTN_STYLE.format(color=color))
            btn.clicked.connect(handler)
            layout.addWidget(btn)

        # Alt satır: Tempfiles Yükle | Test Klasörü Yükle
        bottom_row = QHBoxLayout()

        btn_load_tempfiles = QPushButton("Tempfiles Yükle")
        btn_load_tempfiles.setStyleSheet(MAIN_BTN_STYLE.format(color="#607D8B"))
        btn_load_tempfiles.clicked.connect(self.load_tempfiles)
        bottom_row.addWidget(btn_load_tempfiles)

        btn_load_test = QPushButton("Test Klasörü Yükle")
        btn_load_test.setStyleSheet(MAIN_BTN_STYLE.format(color="#607D8B"))
        btn_load_test.clicked.connect(self.load_test_folder)
        bottom_row.addWidget(btn_load_test)

        layout.addLayout(bottom_row)

        layout.addStretch()

        # Uygulama açılışında tempfiles kontrol et, sonra config yükle
        self._check_tempfiles()
        self._auto_load_config()

    def _check_tempfiles(self):
        """Başlangıçta tempfiles klasörü doluysa silme seçeneği sunar."""
        tempfiles_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tempfiles")
        if not os.path.exists(tempfiles_dir):
            return
        entries = os.listdir(tempfiles_dir)
        if not entries:
            return
        reply = QMessageBox.question(
            self, "Taslakları Sil",
            "Tempfiles klasöründe kayıtlı taslaklar var, silinsin mi?\n(Genel bilgiler dahil her şey silinir)",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            for name in entries:
                path = os.path.join(tempfiles_dir, name)
                try:
                    if os.path.isfile(path):
                        os.unlink(path)
                    elif os.path.isdir(path):
                        shutil.rmtree(path)
                except Exception as e:
                    print(f"Hata: {path} silinemedi. {e}")

    def _auto_load_config(self):
        """Tempfiles'ta kayıtlı config varsa otomatik yükler."""
        if global_data.load_config():
            test_no = global_data.config.get("TEST_NO") or "-"
            test_date = global_data.config.get("TEST_DATE") or "-"
            self.lbl_status.setText(f"Kayıtlı bilgiler yüklendi  |  Test No: {test_no}  |  Tarih: {test_date}")
        else:
            self.lbl_status.setText("Henüz kayıtlı bilgi yok. 'Genel Bilgileri Gir' ile başlayın.")

    def _update_status(self):
        test_no = global_data.config.get("TEST_NO") or "-"
        test_date = global_data.config.get("TEST_DATE") or "-"
        program = global_data.config.get("PROGRAM") or "-"
        self.lbl_status.setText(f"Test No: {test_no}  |  Tarih: {test_date}  |  Program: {program}")

    def open_global_info(self):
        dialog = ReportDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            for key, val in data.items():
                global_data.config[key] = val
            global_data.config["PROJECT"] = global_data.config["PROGRAM"]

            # Otomatik kaydet
            saved_path = global_data.save_config()
            self._update_status()
            QMessageBox.information(self, "Başarılı", f"Genel bilgiler kaydedildi.\n{saved_path}")

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

    def open_eva_app(self):
        self.hide()
        self.eva_window = EvaApp(main_window=self)
        self.eva_window.show()

    # ------------------------------------------------------------------
    #  Test Klasörünü Seç
    # ------------------------------------------------------------------
    FOLDER_TO_CATEGORY = {
        "PRE":               "PRE",
        "POST":              "POST",
        "TEARDOWN":          "TEARDOWN",
        "HANDLE-SIDE COVER": "HANDLE_SIDE_COVER",
    }
    VALID_PHOTO_EXT = {".jpg", ".jpeg", ".png"}

    def select_test_folder(self):
        directory = QFileDialog.getExistingDirectory(self, "Test Klasörünü Seç", "")
        if not directory:
            return

        folder_name = os.path.basename(directory)

        # REPORT_NO olarak kaydet
        global_data.config["REPORT_NO"] = folder_name
        global_data.save_config()
        self._update_status()

        # PHOTOS alt klasörünü bul
        photos_dir = os.path.join(directory, "PHOTOS")
        if not os.path.isdir(photos_dir):
            QMessageBox.warning(
                self, "Uyarı",
                f"Seçilen klasörde PHOTOS alt klasörü bulunamadı:\n{photos_dir}",
            )
            return

        # Fotoğrafları tara
        photo_map = {}
        summary_lines = []
        for folder_name_key, category in self.FOLDER_TO_CATEGORY.items():
            cat_dir = os.path.join(photos_dir, folder_name_key)
            if not os.path.isdir(cat_dir):
                continue
            files = sorted(
                os.path.join(cat_dir, f)
                for f in os.listdir(cat_dir)
                if os.path.splitext(f)[1].lower() in self.VALID_PHOTO_EXT
            )
            if files:
                photo_map[category] = files
                summary_lines.append(f"  {folder_name_key}: {len(files)} fotoğraf")

        if not photo_map:
            QMessageBox.warning(
                self, "Uyarı",
                "PHOTOS klasöründe hiç fotoğraf bulunamadı.\n"
                "Beklenen alt klasörler: PRE, POST, TEARDOWN, HANDLE-SIDE COVER",
            )
            return

        summary = "\n".join(summary_lines)
        reply = QMessageBox.question(
            self, "Fotoğraflar Bulundu",
            f"Klasör: {folder_name}\n\n{summary}\n\n"
            f"Raporlar oluşturulsun mu?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes,
        )
        if reply != QMessageBox.Yes:
            return

        # Çıktı klasörü: tempfiles/<folder_name>-Photos
        root_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(root_dir, "tempfiles", f"{folder_name}-Photos")

        try:
            generator = PhotoReportApp()
            created = generator.batch_generate(photo_map, output_dir, folder_name)

            result_text = "\n".join(os.path.basename(p) for p in created)
            QMessageBox.information(
                self, "Başarılı",
                f"{len(created)} rapor oluşturuldu.\n\n"
                f"Klasör: {output_dir}\n\n{result_text}",
            )
        except Exception as exc:
            QMessageBox.critical(
                self, "Hata", f"Rapor oluşturulurken hata:\n{exc}"
            )

    def load_tempfiles(self):
        """Tempfiles klasöründen config yükler."""
        tempfiles_default = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tempfiles")
        directory = QFileDialog.getExistingDirectory(
            self, "Tempfiles Klasörünü Seç", tempfiles_default,
        )
        if not directory:
            return

        if global_data.load_config(directory):
            self._update_status()
            QMessageBox.information(self, "Başarılı", f"Bilgiler yüklendi:\n{directory}")
        else:
            QMessageBox.warning(self, "Bulunamadı", f"Seçilen klasörde config.json bulunamadı:\n{directory}")

    def load_test_folder(self):
        """Daha önce oluşturulmuş test klasörünü (tempfiles içinden) açar."""
        tempfiles_default = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tempfiles")
        directory = QFileDialog.getExistingDirectory(
            self, "Test Klasörünü Yükle", tempfiles_default,
        )
        if not directory:
            return

        # Klasördeki dosyaları listele
        files = [f for f in os.listdir(directory) if f.endswith(".docx")]
        if not files:
            QMessageBox.warning(self, "Boş", f"Seçilen klasörde docx dosyası bulunamadı:\n{directory}")
            return

        try:
            os.startfile(directory)
        except Exception:
            pass

        file_list = "\n".join(files)
        QMessageBox.information(
            self, "Test Klasörü",
            f"Klasör: {os.path.basename(directory)}\n\n"
            f"Dosyalar ({len(files)}):\n{file_list}",
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec_())
