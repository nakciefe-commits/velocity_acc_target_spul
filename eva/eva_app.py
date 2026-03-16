import os
import re
import tempfile
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

from docxtpl import DocxTemplate, InlineImage
from docx import Document as DocxDocument
from docx.shared import Mm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docxcompose.composer import Composer

import shared.global_data as global_data


# ---------------------------------------------------------------------------
#  Dosya adından sınıflandırma
# ---------------------------------------------------------------------------
#
#  Dosya isimleri ornekleri:
#    005_10SEBE0000G4FO0P_3_S_.png   -> Belt (SEBE), shoulder/lap fark etmez
#    006_10SEBE0000B3FO0P_4_S_.png   -> Belt
#    007_11HEAD0000H3ACXP_5_R_.png   -> Head, Resultant (_R_)
#    008_11HEAD0000H3ACXP_5_S_.png   -> Head, X axis (ACXP + _S_)
#    009_11HEAD0000H3ACYP_6_S_.png   -> Head, Y axis (ACYP + _S_)
#    010_11HEAD0000H3ACZP_7_S_.png   -> Head, Z axis (ACZP + _S_)
#    011_11CHST0000H3ACXP_8_R_.png   -> Chest, Resultant
#    013_11CHST0000H3ACZP_10_S_.png  -> Chest, Z axis
#    015_11PELV0000H3ACXP_11_R_.png  -> Pelvis, Resultant
#    017_11PELV0000H3ACYP_12_S_.png  -> Pelvis, Y axis
#
#  Siniflandirma:
#    SEBE  -> Belt template   (shoulder + lap ayni sayfada, sira farketmez)
#    HEAD  -> Head_r_x veya Head_y_z
#    CHST  -> Chest_r_x veya Chest_y_z
#    PELV  -> Pelvis_r_x veya Pelvis_y_z
#
#  _R_ -> resultant  -> _r_x template'inin IMG_TOP'u
#  _S_ + ACXP -> x   -> _r_x template'inin IMG_BOTTOM'u
#  _S_ + ACYP -> y   -> _y_z template'inin IMG_TOP'u
#  _S_ + ACZP -> z   -> _y_z template'inin IMG_BOTTOM'u
# ---------------------------------------------------------------------------

# Body region keyword -> template group name
REGION_MAP = {
    "SEBE": "Belt",
    "HEAD": "Head",
    "CHST": "Chest",
    "PELV": "Pelvis",
}

# Belt alt tipleri (shoulder / lap) - her ikisi de ayni sayfaya gider
BELT_SUBTYPES = {
    "SHBE": "shoulder",
    "LABE": "lap",
}


def classify_eva_file(filename):
    """
    Dosya adindan siniflandirma yapar.

    Dondurur: (template_file, slot)
      template_file: "Belt", "Head_r_x", "Head_y_z", "Chest_r_x", ...
      slot:          "shoulder"/"lap" (Belt icin) veya "IMG_TOP"/"IMG_BOTTOM" (digerleri)
    Eslenemezse: (None, None)
    """
    name_upper = filename.upper()

    # 1) Hangi bolge?
    region = None
    for keyword, region_name in REGION_MAP.items():
        if keyword in name_upper:
            region = region_name
            break

    if region is None:
        return None, None

    # 2) Belt ozel durum
    if region == "Belt":
        for sub_kw, sub_name in BELT_SUBTYPES.items():
            if sub_kw in name_upper:
                return "Belt", sub_name
        # SEBE var ama SHBE/LABE yok - "belt_auto" olarak isle,
        # group_eva_files icerisinde sirayla shoulder/lap olarak atanacak
        return "Belt", "belt_auto"

    # 3) HEAD / CHST / PELV -> axis belirleme
    #    _R_ -> resultant
    #    _S_ -> specific axis (ACXP, ACYP, ACZP)
    is_resultant = "_R_" in name_upper

    if is_resultant:
        # Resultant -> _r_x template, IMG_TOP slot
        return "%s_r_x" % region, "IMG_TOP"

    # _S_ durumu: axis bul
    if "ACXP" in name_upper:
        return "%s_r_x" % region, "IMG_BOTTOM"
    elif "ACYP" in name_upper:
        return "%s_y_z" % region, "IMG_TOP"
    elif "ACZP" in name_upper:
        return "%s_y_z" % region, "IMG_BOTTOM"

    # Axis bulunamadi - resultant olarak varsay
    return "%s_r_x" % region, "IMG_TOP"


def group_eva_files(file_paths):
    """
    Dosyalari template_file -> slot -> [paths] seklinde gruplar.
    Ornek:
      {
        "Belt":       {"shoulder": [...], "lap": [...]},
        "Head_r_x":   {"IMG_TOP": [...], "IMG_BOTTOM": [...]},
        "Head_y_z":   {"IMG_TOP": [...], "IMG_BOTTOM": [...]},
        ...
      }
    """
    groups = {}
    unmatched = []
    for path in file_paths:
        fname = os.path.basename(path)
        template_file, slot = classify_eva_file(fname)
        if template_file is None:
            unmatched.append(path)
            continue
        if template_file not in groups:
            groups[template_file] = {}
        if slot not in groups[template_file]:
            groups[template_file][slot] = []
        groups[template_file][slot].append(path)

    # Belt "belt_auto" slotlarini sirayla shoulder/lap olarak dagit
    if "Belt" in groups and "belt_auto" in groups["Belt"]:
        auto_files = groups["Belt"].pop("belt_auto")
        if "shoulder" not in groups["Belt"]:
            groups["Belt"]["shoulder"] = []
        if "lap" not in groups["Belt"]:
            groups["Belt"]["lap"] = []
        for i, f in enumerate(auto_files):
            if i % 2 == 0:
                groups["Belt"]["shoulder"].append(f)
            else:
                groups["Belt"]["lap"].append(f)

    return groups, unmatched


# ---------------------------------------------------------------------------
#  Template bilgileri
# ---------------------------------------------------------------------------
TEMPLATES = {
    "Belt": {
        "file": "Belt.docx",
        "title": "Seat Belt Force",
        "color": "#E91E63",
        "slots": ["shoulder", "lap"],
        "img_keys": {"shoulder": "SHOULDER_IMG", "lap": "LAP_IMG"},
    },
    "Head_r_x": {
        "file": "Head_r_x.docx",
        "title": "Head Accel. Resultant & X",
        "color": "#2196F3",
        "slots": ["IMG_TOP", "IMG_BOTTOM"],
        "img_keys": {"IMG_TOP": "IMG_TOP", "IMG_BOTTOM": "IMG_BOTTOM"},
    },
    "Head_y_z": {
        "file": "Head_y_z.docx",
        "title": "Head Accel. Y & Z",
        "color": "#2196F3",
        "slots": ["IMG_TOP", "IMG_BOTTOM"],
        "img_keys": {"IMG_TOP": "IMG_TOP", "IMG_BOTTOM": "IMG_BOTTOM"},
    },
    "Chest_r_x": {
        "file": "Chest_r_x.docx",
        "title": "Chest Accel. Resultant & X",
        "color": "#4CAF50",
        "slots": ["IMG_TOP", "IMG_BOTTOM"],
        "img_keys": {"IMG_TOP": "IMG_TOP", "IMG_BOTTOM": "IMG_BOTTOM"},
    },
    "Chest_y_z": {
        "file": "Chest_y_z.docx",
        "title": "Chest Accel. Y & Z",
        "color": "#4CAF50",
        "slots": ["IMG_TOP", "IMG_BOTTOM"],
        "img_keys": {"IMG_TOP": "IMG_TOP", "IMG_BOTTOM": "IMG_BOTTOM"},
    },
    "Pelvis_r_x": {
        "file": "Pelvis_r_x.docx",
        "title": "Pelvis Accel. Resultant & X",
        "color": "#FF9800",
        "slots": ["IMG_TOP", "IMG_BOTTOM"],
        "img_keys": {"IMG_TOP": "IMG_TOP", "IMG_BOTTOM": "IMG_BOTTOM"},
    },
    "Pelvis_y_z": {
        "file": "Pelvis_y_z.docx",
        "title": "Pelvis Accel. Y & Z",
        "color": "#FF9800",
        "slots": ["IMG_TOP", "IMG_BOTTOM"],
        "img_keys": {"IMG_TOP": "IMG_TOP", "IMG_BOTTOM": "IMG_BOTTOM"},
    },
}


class EvaApp(QMainWindow):
    IMAGE_FILTER = "EVA Grafikleri (*.png *.jpg *.jpeg)"

    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window
        self.setWindowTitle("EVA Modulu")
        self.resize(800, 640)

        self.selected_files = []

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        title = QLabel("EVA Modulu")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 20px; font-weight: bold; margin-bottom: 10px;")
        main_layout.addWidget(title)

        subtitle = QLabel(
            "EVA grafiklerini (PNG) toplu yukleyin.\n"
            "Dosya adlarindan otomatik kategorilere ayrilir:\n"
            "SEBE->Belt | HEAD->Head | CHST->Chest | PELV->Pelvis\n"
            "_R_->Resultant | ACXP->X | ACYP->Y | ACZP->Z"
        )
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setWordWrap(True)
        subtitle.setStyleSheet("color: #616161; margin-bottom: 8px;")
        main_layout.addWidget(subtitle)

        # Dosya secimi
        file_group = QGroupBox("EVA Dosyalari")
        file_layout = QVBoxLayout(file_group)

        btn_row = QHBoxLayout()
        btn_select = QPushButton("EVA Dosyalarini Sec")
        btn_select.clicked.connect(self.select_files)
        btn_row.addWidget(btn_select)

        btn_clear = QPushButton("Tumunu Temizle")
        btn_clear.clicked.connect(self.clear_files)
        btn_row.addWidget(btn_clear)

        file_layout.addLayout(btn_row)

        self.list_widget = QListWidget()
        self.list_widget.setMinimumHeight(150)
        file_layout.addWidget(self.list_widget)

        main_layout.addWidget(file_group)

        # Onizleme
        preview_group = QGroupBox("Siniflandirma Onizlemesi")
        preview_layout = QVBoxLayout(preview_group)
        self.lbl_preview = QLabel("Henuz dosya secilmedi.")
        self.lbl_preview.setWordWrap(True)
        self.lbl_preview.setStyleSheet("color: #424242; font-size: 12px;")
        preview_layout.addWidget(self.lbl_preview)
        main_layout.addWidget(preview_group)

        # Butonlar
        action_row = QHBoxLayout()

        self.btn_generate = QPushButton("Raporlari Olustur")
        self.btn_generate.setStyleSheet(
            "font-size: 15px; padding: 12px; background-color: #E91E63; color: white; font-weight: bold;"
        )
        self.btn_generate.clicked.connect(self.generate_reports)

        self.btn_back = QPushButton("Ana Menuye Don")
        self.btn_back.setStyleSheet("font-size: 15px; padding: 12px; background-color: #9E9E9E; color: white;")
        self.btn_back.clicked.connect(self.close)

        action_row.addWidget(self.btn_generate)
        action_row.addWidget(self.btn_back)
        main_layout.addLayout(action_row)

        self.progress = QProgressBar()
        self.progress.setMinimum(0)
        self.progress.setValue(0)
        self.progress.setFormat("Hazir")
        main_layout.addWidget(self.progress)

    # ------------------------------------------------------------------
    #  Dosya secimi
    # ------------------------------------------------------------------
    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "EVA Dosyalarini Sec", "", self.IMAGE_FILTER,
        )
        if not files:
            return
        valid_ext = {".png", ".jpg", ".jpeg"}
        dedup = set(self.selected_files)
        for f in files:
            if os.path.splitext(f)[1].lower() in valid_ext and f not in dedup:
                self.selected_files.append(f)
                dedup.add(f)
        self._refresh_list()

    def clear_files(self):
        self.selected_files = []
        self._refresh_list()

    def _refresh_list(self):
        self.list_widget.clear()
        if not self.selected_files:
            self.list_widget.addItem(QListWidgetItem("Henuz dosya secilmedi"))
            self.list_widget.item(0).setFlags(Qt.NoItemFlags)
            self.lbl_preview.setText("Henuz dosya secilmedi.")
            return

        for i, f in enumerate(self.selected_files):
            fname = os.path.basename(f)
            tpl, slot = classify_eva_file(fname)
            tag = "[%s/%s]" % (tpl, slot) if tpl else "[???]"
            self.list_widget.addItem(QListWidgetItem("%d. %s %s" % (i + 1, tag, fname)))

        groups, unmatched = group_eva_files(self.selected_files)
        lines = []
        for tpl_key in sorted(groups.keys()):
            tpl_info = TEMPLATES.get(tpl_key, {})
            tpl_title = tpl_info.get("title", tpl_key)
            slots = groups[tpl_key]
            for slot_name, files in sorted(slots.items()):
                lines.append("  %s / %s: %d dosya" % (tpl_title, slot_name, len(files)))
        if unmatched:
            lines.append("  Taninmayan: %d dosya" % len(unmatched))
        self.lbl_preview.setText("\n".join(lines) if lines else "Siniflandirma sonucu yok.")

    # ------------------------------------------------------------------
    #  Template ve render islemleri
    # ------------------------------------------------------------------
    def _resolve_template(self, template_key):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        tpl_info = TEMPLATES[template_key]
        return os.path.join(base_dir, "template", tpl_info["file"])

    def _ensure_output_dir(self):
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        test_no = str(global_data.config.get("TEST_NO") or "UNSPECIFIED").replace("/", "_").replace("\\", "_").strip()
        out_dir = os.path.join(root_dir, "tempfiles", "eva_reports", test_no)
        os.makedirs(out_dir, exist_ok=True)
        return out_dir

    def _build_base_context(self):
        return {
            "DUMMY_PCT": global_data.config.get("DUMMY_PCT") or "",
            "SENSOR": global_data.config.get("SENSOR") or "",
            "TEST_NO": global_data.config.get("TEST_NO") or "",
            "TEST_DATE": global_data.config.get("TEST_DATE") or "",
            "PROJECT": global_data.config.get("PROJECT") or global_data.config.get("PROGRAM") or "",
        }

    def _render_page(self, template_key, img_top_path, img_bottom_path):
        """
        Herhangi bir template'i render eder.
        template_key: "Belt", "Head_r_x", "Head_y_z", ...
        img_top_path: ust gorselin yolu (None olabilir)
        img_bottom_path: alt gorselin yolu (None olabilir)
        Dondurur: DocxTemplate (rendered)
        """
        template_path = self._resolve_template(template_key)
        doc = DocxTemplate(template_path)
        context = self._build_base_context()

        tpl_info = TEMPLATES[template_key]
        img_keys = tpl_info["img_keys"]
        slots = tpl_info["slots"]

        # slots[0] = top image slot name, slots[1] = bottom image slot name
        top_key = img_keys[slots[0]]
        bottom_key = img_keys[slots[1]]

        if img_top_path:
            context[top_key] = InlineImage(doc, img_top_path, height=Mm(100))
        else:
            context[top_key] = ""

        if img_bottom_path:
            context[bottom_key] = InlineImage(doc, img_bottom_path, height=Mm(100))
        else:
            context[bottom_key] = ""

        doc.render(context)
        return doc

    def _merge_docs(self, doc_list, out_path):
        """
        Birden fazla DocxTemplate'i tek dosyada birlestirir.
        docxcompose kullanarak gorsel referanslarini da dogru kopyalar.
        """
        # Her DocxTemplate'i gecici dosyaya kaydet
        temp_paths = []
        for doc in doc_list:
            tmp = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
            tmp.close()
            doc.save(tmp.name)
            temp_paths.append(tmp.name)

        try:
            # Composer ile birlestir
            master = DocxDocument(temp_paths[0])
            composer = Composer(master)
            for tmp_path in temp_paths[1:]:
                extra = DocxDocument(tmp_path)
                composer.append(extra)
            composer.save(out_path)
        finally:
            # Gecici dosyalari temizle
            for tmp_path in temp_paths:
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass

    # ------------------------------------------------------------------
    #  Rapor uretimi
    # ------------------------------------------------------------------
    def generate_reports(self):
        if not self.selected_files:
            QMessageBox.warning(self, "Uyari", "Once EVA dosyalarini secin.")
            return

        groups, unmatched = group_eva_files(self.selected_files)

        # Debug: siniflandirma sonuclarini yazdir
        print("\n=== EVA SINIFLANDIRMA ===")
        for tpl_key in sorted(groups.keys()):
            print("  Template: %s" % tpl_key)
            for slot_name, files in sorted(groups[tpl_key].items()):
                for f in files:
                    print("    %s: %s" % (slot_name, os.path.basename(f)))
        if unmatched:
            print("  Taninmayan:")
            for f in unmatched:
                print("    %s" % os.path.basename(f))
        print("=========================\n")

        if not groups:
            QMessageBox.warning(
                self, "Uyari",
                "Dosya adlarinda taninabilir anahtar kelime bulunamadi.\n"
                "Beklenen: SEBE, HEAD, CHST, PELV",
            )
            return

        output_dir = self._ensure_output_dir()
        all_docs = []

        try:
            total_templates = len(groups)
            self.progress.setMaximum(total_templates)
            step = 0

            # Her template grubu icin sayfa(lar) olustur
            # Siralamayi sabit tutalim: Belt, Head_r_x, Head_y_z, Chest_r_x, ...
            template_order = [
                "Belt", "Head_r_x", "Head_y_z",
                "Chest_r_x", "Chest_y_z",
                "Pelvis_r_x", "Pelvis_y_z",
            ]

            for tpl_key in template_order:
                if tpl_key not in groups:
                    continue

                tpl_info = TEMPLATES[tpl_key]
                self.progress.setFormat("%s olusturuluyor..." % tpl_info["title"])
                QApplication.processEvents()

                slot_data = groups[tpl_key]
                slots = tpl_info["slots"]
                top_slot = slots[0]
                bottom_slot = slots[1]

                top_files = slot_data.get(top_slot, [])
                bottom_files = slot_data.get(bottom_slot, [])

                # Her cift (top, bottom) icin bir sayfa
                max_count = max(len(top_files), len(bottom_files), 1)

                # Eger hicbir slot'ta dosya yoksa atla
                if not top_files and not bottom_files:
                    continue

                for i in range(max_count):
                    img_top = top_files[i] if i < len(top_files) else None
                    img_bottom = bottom_files[i] if i < len(bottom_files) else None
                    if img_top or img_bottom:
                        print("  Sayfa render: %s  TOP=%s  BOTTOM=%s" % (
                            tpl_key,
                            os.path.basename(img_top) if img_top else "(bos)",
                            os.path.basename(img_bottom) if img_bottom else "(bos)",
                        ))
                        doc = self._render_page(tpl_key, img_top, img_bottom)
                        all_docs.append(doc)

                step += 1
                self.progress.setValue(step)

            if not all_docs:
                QMessageBox.warning(self, "Uyari", "Olusturulacak sayfa bulunamadi.")
                return

            # Tek dosyada birlestir
            test_no = str(global_data.config.get("TEST_NO") or "EVA").replace("/", "_")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_name = "EVA_%s_%s.docx" % (test_no, timestamp)
            out_path = os.path.join(output_dir, out_name)

            if len(all_docs) == 1:
                all_docs[0].save(out_path)
            else:
                self._merge_docs(all_docs, out_path)

            # Uyari: eslenemeyenler
            if unmatched:
                unmatched_names = "\n".join(os.path.basename(f) for f in unmatched[:10])
                extra = "\n...ve %d daha" % (len(unmatched) - 10) if len(unmatched) > 10 else ""
                QMessageBox.warning(
                    self, "Taninmayan Dosyalar",
                    "Su dosyalar eslenmedi:\n%s%s" % (unmatched_names, extra),
                )

            QMessageBox.information(
                self, "Basarili",
                "EVA raporu olusturuldu.\n\n%s" % out_path,
            )
            self.progress.setFormat("Tamamlandi")

        except Exception as exc:
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Hata", "Rapor olusturulurken hata:\n%s" % exc)
            self.progress.setFormat("Hata!")

    def closeEvent(self, event):
        if self.main_window is not None:
            self.main_window.show()
        event.accept()
