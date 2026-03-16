import os
from datetime import datetime

from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QFileDialog,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QGroupBox,
    QProgressBar,
)
from PyQt5.QtCore import Qt

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Mm

import shared.global_data as global_data


class PhotoReportApp(QMainWindow):
    IMAGE_FILTER = "Fotoğraflar (*.jpg *.jpeg *.png)"
    CATEGORIES = {
        "PRE": {
            "title": "Pre",
            "template": "PRE.docx",
            "color": "#03A9F4",
        },
        "POST": {
            "title": "Post",
            "template": "POST.docx",
            "color": "#4CAF50",
        },
        "TEARDOWN": {
            "title": "Teardown",
            "template": "TEARDOWN.docx",
            "color": "#FF9800",
        },
        "HANDLE_SIDE_COVER": {
            "title": "Handle Side Cover",
            "template": "HANDLE_SIDE_COVER.docx",
            "color": "#9C27B0",
        },
    }

    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window
        self.setWindowTitle("Photo Report")
        self.resize(900, 760)

        self.selected_files = {key: [] for key in self.CATEGORIES}
        self.list_widgets = {}

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        title = QLabel("Photo Report Modülü")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 20px; font-weight: bold; margin-bottom: 10px;")
        main_layout.addWidget(title)

        subtitle = QLabel("Her kategori için çoklu fotoğraf seçebilir ve ayrı Word raporu oluşturabilirsiniz.")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #616161; margin-bottom: 8px;")
        main_layout.addWidget(subtitle)

        for category, cfg in self.CATEGORIES.items():
            main_layout.addWidget(self._build_category_section(category, cfg))

        action_row = QHBoxLayout()

        self.btn_generate = QPushButton("Seçilen Kategoriler İçin Raporları Oluştur")
        self.btn_generate.setStyleSheet(
            "font-size: 15px; padding: 12px; background-color: #1976D2; color: white; font-weight: bold;"
        )
        self.btn_generate.clicked.connect(self.generate_reports)

        self.btn_back = QPushButton("Ana Menüye Dön")
        self.btn_back.setStyleSheet("font-size: 15px; padding: 12px; background-color: #9E9E9E; color: white;")
        self.btn_back.clicked.connect(self.close_and_return)

        action_row.addWidget(self.btn_generate)
        action_row.addWidget(self.btn_back)

        main_layout.addLayout(action_row)

        self.progress = QProgressBar()
        self.progress.setMinimum(0)
        self.progress.setValue(0)
        self.progress.setFormat("Hazır")
        main_layout.addWidget(self.progress)

        self.lbl_output = QLabel("Çıktı klasörü: tempfiles/photo_reports")
        self.lbl_output.setStyleSheet("color: #616161; margin-top: 4px;")
        main_layout.addWidget(self.lbl_output)

    def _build_category_section(self, category, cfg):
        group = QGroupBox(cfg["title"])
        group.setStyleSheet(f"QGroupBox {{ font-weight: bold; color: {cfg['color']}; }}")

        layout = QVBoxLayout(group)

        button_row = QHBoxLayout()

        btn_select = QPushButton(f"{cfg['title']} fotoğraf seç")
        btn_select.clicked.connect(lambda _, c=category: self.select_photos(c))
        button_row.addWidget(btn_select)

        btn_reset = QPushButton("Listeyi temizle")
        btn_reset.clicked.connect(lambda _, c=category: self.clear_category(c))
        button_row.addWidget(btn_reset)

        layout.addLayout(button_row)

        list_widget = QListWidget()
        list_widget.setSelectionMode(QListWidget.ExtendedSelection)
        list_widget.setMinimumHeight(80)
        layout.addWidget(list_widget)

        self.list_widgets[category] = list_widget
        self._refresh_list(category)

        return group

    def select_photos(self, category):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            f"{self.CATEGORIES[category]['title']} fotoğraflarını seç",
            "",
            self.IMAGE_FILTER,
        )

        if not files:
            return

        valid_ext = {".jpg", ".jpeg", ".png"}
        dedup = set(self.selected_files[category])

        for path in files:
            ext = os.path.splitext(path)[1].lower()
            if ext in valid_ext and path not in dedup:
                self.selected_files[category].append(path)
                dedup.add(path)

        self._refresh_list(category)

    def clear_category(self, category):
        self.selected_files[category] = []
        self._refresh_list(category)

    def _refresh_list(self, category):
        list_widget = self.list_widgets[category]
        list_widget.clear()

        files = self.selected_files[category]
        if not files:
            list_widget.addItem(QListWidgetItem("Henüz fotoğraf seçilmedi"))
            list_widget.item(0).setFlags(Qt.NoItemFlags)
            return

        for photo in files:
            list_widget.addItem(QListWidgetItem(os.path.basename(photo)))

    def _ensure_output_dir(self):
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        test_no = str(global_data.config.get("TEST_NO") or "UNSPECIFIED").replace("/", "_").replace("\\", "_").strip()
        date_part = datetime.now().strftime("%Y%m%d")
        out_dir = os.path.join(root_dir, "tempfiles", "photo_reports", test_no, date_part)
        os.makedirs(out_dir, exist_ok=True)
        return out_dir

    def _resolve_template_path(self, category):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        template_name = self.CATEGORIES[category]["template"]
        return os.path.join(base_dir, "templates", template_name)

    def _safe_output_name(self, category, output_dir):
        test_no = global_data.config.get("TEST_NO") or "UNSPECIFIED"
        normalized_test = str(test_no).replace("/", "_").replace("\\", "_").strip()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = f"photo_report_{category.lower()}_{normalized_test}_{timestamp}"

        candidate = os.path.join(output_dir, f"{base_name}.docx")
        idx = 1
        while os.path.exists(candidate):
            candidate = os.path.join(output_dir, f"{base_name}_{idx}.docx")
            idx += 1

        return candidate

    def _chunk_photos(self, photos, size=6):
        return [photos[idx : idx + size] for idx in range(0, len(photos), size)]

    def _remove_template_tables(self, doc):
        for table in list(doc.tables):
            table._element.getparent().remove(table._element)

    def _insert_photo_into_cell(self, cell, photo_path):
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        run = paragraph.add_run()
        run.add_picture(photo_path, width=Mm(78))

    def _add_photo_page(self, doc, photo_chunk):
        # 2x3 düzen: 3 satır x 2 sütun = 6 slot
        table = doc.add_table(rows=3, cols=2)
        table.style = "Table Grid"

        for idx in range(6):
            row_idx = idx // 2
            col_idx = idx % 2
            cell = table.cell(row_idx, col_idx)
            cell.text = ""
            if idx < len(photo_chunk):
                self._insert_photo_into_cell(cell, photo_chunk[idx])

    def _build_document_from_template(self, template_path, photos):
        doc = Document(template_path)
        self._remove_template_tables(doc)

        photo_chunks = self._chunk_photos(photos, size=6)
        print(f"[PhotoReport] selected photo count: {len(photos)}")
        print(f"[PhotoReport] page count (6'lı): {len(photo_chunks)}")

        for page_no, chunk in enumerate(photo_chunks, start=1):
            if page_no > 1:
                doc.add_page_break()
            print(f"[PhotoReport] page {page_no}: {len(chunk)} foto")
            self._add_photo_page(doc, chunk)

        return doc

    def _render_category_report(self, category, output_path):
        template_path = self._resolve_template_path(category)
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template bulunamadı: {template_path}")

        doc = self._build_document_from_template(template_path, self.selected_files[category])
        doc.save(output_path)

    def generate_reports(self):
        selected_categories = [key for key, files in self.selected_files.items() if files]
        if not selected_categories:
            QMessageBox.warning(self, "Uyarı", "Rapor üretmek için en az bir kategoriye fotoğraf ekleyin.")
            return

        output_dir = self._ensure_output_dir()
        self.progress.setMinimum(0)
        self.progress.setMaximum(len(selected_categories))
        self.progress.setValue(0)

        created_files = []

        try:
            for step, category in enumerate(selected_categories, start=1):
                output_path = self._safe_output_name(category, output_dir)
                self._render_category_report(category, output_path)
                created_files.append(output_path)

                self.progress.setValue(step)
                self.progress.setFormat(f"%p - {self.CATEGORIES[category]['title']} oluşturuldu")
                QApplication.processEvents()

            result_text = "\n".join(created_files)
            QMessageBox.information(
                self,
                "Başarılı",
                f"{len(created_files)} rapor başarıyla oluşturuldu.\n\nÇıktılar:\n{result_text}",
            )
        except Exception as exc:
            QMessageBox.critical(self, "Hata", f"Rapor oluşturulurken hata oluştu:\n{exc}")

    def close_and_return(self):
        self.close()
        if self.main_window is not None:
            self.main_window.show()
