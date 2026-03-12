import sys
import os
import subprocess

def _check_and_install_dependencies():
    required_packages = ['pandas', 'numpy', 'PyQt5', 'matplotlib', 'docxtpl', 'openpyxl', 'xlrd']
    for pkg in required_packages:
        try:
            __import__(pkg)
        except ImportError:
            print(f"Eksik kütüphane tespit edildi, yükleniyor: {pkg}...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])
                print(f"{pkg} başarıyla yüklendi.")
            except Exception as e:
                print(f"{pkg} yüklenirken hata oluştu: {e}")

_check_and_install_dependencies()

import pandas as pd
import numpy as np
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                             QMessageBox, QDoubleSpinBox, QGroupBox, QLineEdit,
                             QCheckBox, QGridLayout, QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
                             QDialog, QFormLayout, QDialogButtonBox)
from PyQt5.QtCore import Qt

import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import shared.global_data as global_data
from PyQt5.QtCore import Qt

import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

class SledAnalyzerApp(QMainWindow):
    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window
        self.setWindowTitle("Sled Test Analyzer (Multi-Graph)")
        self.resize(1100, 900)
        
        self.actual_path = None
        self.target_path = None
        self.df_actual = None
        self.df_target = None
        
        # State
        self.current_graph_idx = 0
        self.graphs = ["Spul", "Acceleration vs Velocity", "Actual vs Target Acceleration"]
        self.local_offsets = [0.0, 0.0, 0.0]
        
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        # --- Top Area Layout ---
        top_layout = QHBoxLayout()
        
        # --- Control Panel (Left) ---
        control_group = QGroupBox("Veri Yükleme ve Ayarlar")
        control_layout = QVBoxLayout()
        control_group.setLayout(control_layout)
        
        # File Selection
        self.btn_actual = QPushButton("Actual Data Yükle")
        self.btn_actual.clicked.connect(self.load_actual)
        self.lbl_actual = QLabel("Seçilmedi")
        control_layout.addWidget(self.btn_actual)
        control_layout.addWidget(self.lbl_actual)
        
        control_layout.addSpacing(10)
        
        self.btn_target = QPushButton("Target Data Yükle")
        self.btn_target.clicked.connect(self.load_target)
        self.lbl_target = QLabel("Seçilmedi")
        control_layout.addWidget(self.btn_target)
        control_layout.addWidget(self.lbl_target)
        
        control_layout.addStretch() # Push items up
        
        # Action Buttons
        self.btn_generate = QPushButton("Oluştur / Güncelle")
        self.btn_generate.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        self.btn_generate.clicked.connect(self.generate_plots)
        control_layout.addWidget(self.btn_generate)
        
        top_layout.addWidget(control_group, stretch=1)
        
        # --- Offset Table Panel (Right) ---
        offset_group = QGroupBox("Grafik Offset Ayarları (ms)")
        offset_layout = QVBoxLayout()
        offset_group.setLayout(offset_layout)
        
        self.table_offset = QTableWidget()
        self.table_offset.setColumnCount(3)
        self.table_offset.setHorizontalHeaderLabels(["Değişken / Grafik", "Mevcut Değer", "Kullanım Yeri"])
        self.table_offset.setRowCount(3)
        self.table_offset.verticalHeader().setVisible(False)
        self.table_offset.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table_offset.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table_offset.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        # Tabloyu dikey olarak sıkıştır
        self.table_offset.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table_offset.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table_offset.setMaximumHeight(150) # 3 satırın tam sığacağı ideal yükseklik
        
        self.table_offset.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_offset.setStyleSheet("QTableWidget { background-color: white; gridline-color: #d3d3d3; } "
                                        "QHeaderView::section { background-color: #f0f0f0; font-weight: bold; }")
        
        # Populate Table
        labels = ["Spul", "Acceleration vs Velocity", "Actual vs Target Acceleration"]
        used_by = ["Spul Hesaplaması", "Hız ve İvme Grafikleri", "İvme Karşılaştırması"]
        
        self.spin_offsets = []
        for i in range(3):
            # Column 0: Variable
            item_var = QTableWidgetItem(labels[i])
            item_var.setFlags(item_var.flags() ^ Qt.ItemIsEditable)
            self.table_offset.setItem(i, 0, item_var)
            
            # Column 1: Current Value (SpinBox inside table)
            spin = QDoubleSpinBox()
            spin.setRange(-10000.0, 10000.0)
            spin.setValue(0.0)
            spin.setSingleStep(4.0) # 0.004 saniye = 4 ms artırım
            spin.setDecimals(1)
            spin.setStyleSheet("border: none; background: transparent;")
            spin.valueChanged.connect(lambda val, idx=i: self.set_local_offset(idx, val))
            self.table_offset.setCellWidget(i, 1, spin)
            self.spin_offsets.append(spin)
            
            # Column 2: Used By (Blue Text)
            item_used = QTableWidgetItem(used_by[i])
            item_used.setFlags(item_used.flags() ^ Qt.ItemIsEditable)
            item_used.setForeground(Qt.blue)
            self.table_offset.setItem(i, 2, item_used)
            
        offset_layout.addWidget(self.table_offset)
        
        # Universal Offset input at bottom of table
        univ_layout = QHBoxLayout()
        univ_layout.addWidget(QLabel("Tüm grafiklere aynı anda evrensel offset uygula:"))
        self.spin_universal = QDoubleSpinBox()
        self.spin_universal.setRange(-10000.0, 10000.0)
        self.spin_universal.setValue(0.0)
        self.spin_universal.setSingleStep(4.0) # 0.004 saniye = 4 ms artırım
        self.spin_universal.setDecimals(1)
        self.spin_universal.valueChanged.connect(self.apply_universal_offset)
        univ_layout.addWidget(self.spin_universal)
        
        # 14 ms tick box
        self.check_14ms = QCheckBox("Tüm Grafikler İçin 14 ms Sabit Offset (Dümdüz 14ms Ayarla)")
        self.check_14ms.stateChanged.connect(self.apply_14ms_offset)
        univ_layout.addWidget(self.check_14ms)
        
        offset_layout.addLayout(univ_layout)
        top_layout.addWidget(offset_group, stretch=2)
        
        main_layout.addLayout(top_layout)
        
        # --- Graph Navigation ---
        nav_layout = QHBoxLayout()
        self.btn_prev = QPushButton("⬅")
        self.btn_prev.setStyleSheet("font-size: 24px; font-weight: bold; width: 60px; height: 40px;")
        self.btn_prev.clicked.connect(self.prev_graph)
        
        self.lbl_graph_name = QLabel(f"{self.graphs[self.current_graph_idx]}")
        self.lbl_graph_name.setAlignment(Qt.AlignCenter)
        self.lbl_graph_name.setStyleSheet("font-size: 16px; font-weight: bold;")
        
        self.btn_next = QPushButton("➡")
        self.btn_next.setStyleSheet("font-size: 24px; font-weight: bold; width: 60px; height: 40px;")
        self.btn_next.clicked.connect(self.next_graph)
        
        nav_layout.addWidget(self.btn_prev)
        nav_layout.addWidget(self.lbl_graph_name)
        nav_layout.addWidget(self.btn_next)
        
        main_layout.addLayout(nav_layout)
        
        # --- Plot Area (Matplotlib) ---
        plot_group = QGroupBox("Grafik Ekranı")
        plot_layout = QVBoxLayout()
        plot_group.setLayout(plot_layout)
        
        self.figure = Figure(figsize=(10, 8))
        self.canvas = FigureCanvas(self.figure)
        plot_layout.addWidget(self.canvas)
        
        # Tablo ayarı
        import matplotlib.gridspec as gridspec
        self.gs = gridspec.GridSpec(2, 1, height_ratios=[4, 1.2]) # Alt tablonun yüksekliğini biraz daha açtım
        self.ax = self.figure.add_subplot(self.gs[0])
        self.ax_table = self.figure.add_subplot(self.gs[1])
        self.ax_table.axis('off')
        
        self.ax2 = None # Sağ eksen için
        
        main_layout.addWidget(plot_group)
        
        # --- Export Area ---
        export_layout = QHBoxLayout()
        export_layout.addWidget(QLabel("Kayıt Dizini:"))
        self.txt_export = QLineEdit(r"c:\Users\pc1\Desktop\adient_data\velocity_acc_target_spul")
        export_layout.addWidget(self.txt_export)
        
        self.btn_browse = QPushButton("Gözat...")
        self.btn_browse.clicked.connect(self.browse_export_dir)
        export_layout.addWidget(self.btn_browse)
        
        self.btn_export = QPushButton("Tüm Grafikleri Kaydet (.png)")
        self.btn_export.clicked.connect(self.export_plots)

        self.btn_report = QPushButton("Rapor Oluştur (Word)")
        self.btn_report.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold;")
        self.btn_report.clicked.connect(self.generate_word_report)
        
        export_layout.addWidget(self.btn_export)
        export_layout.addWidget(self.btn_report)
        
        main_layout.addLayout(export_layout)
        
        # --- Author Info ---
        lbl_author = QLabel("Created by Efe Nakcı")
        lbl_author.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        lbl_author.setStyleSheet("color: gray; font-style: italic; font-size: 11px; padding-top: 5px;")
        main_layout.addWidget(lbl_author)

    def apply_universal_offset(self, val):
        for spin in self.spin_offsets:
            spin.setValue(val)

    def apply_14ms_offset(self, state):
        if state == Qt.Checked:
            # Set all to 14.0 and disable manual edit
            for spin in self.spin_offsets:
                spin.setValue(14.0)
                spin.setEnabled(False)
            self.spin_universal.setValue(14.0)
            self.spin_universal.setEnabled(False)
        else:
            # Re-enable manual edit
            for spin in self.spin_offsets:
                spin.setEnabled(True)
            self.spin_universal.setEnabled(True)

    def browse_export_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Kayıt Klasörü Seç", self.txt_export.text())
        if directory:
            self.txt_export.setText(directory)

    def set_local_offset(self, idx, val):
        self.local_offsets[idx] = val
        if self.current_graph_idx == idx and self.df_actual is not None:
            self.draw_current_graph()

    def prev_graph(self):
        self.current_graph_idx = (self.current_graph_idx - 1) % len(self.graphs)
        self.update_graph_view()

    def next_graph(self):
        self.current_graph_idx = (self.current_graph_idx + 1) % len(self.graphs)
        self.update_graph_view()

    def update_graph_view(self):
        self.lbl_graph_name.setText(f"{self.graphs[self.current_graph_idx]}")
        if self.df_actual is not None:
            self.draw_current_graph()

    def load_actual(self):
        path, _ = QFileDialog.getOpenFileName(self, "Actual Data Seç", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.actual_path = path
            self.lbl_actual.setText(os.path.basename(path))

    def load_target(self):
        path, _ = QFileDialog.getOpenFileName(self, "Target Data Seç", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.target_path = path
            self.lbl_target.setText(os.path.basename(path))

    def process_data(self, df):
        df_proc = df.copy()
        
        # Trim space in column names
        df_proc.columns = df_proc.columns.str.strip()
        
        # Convert necessary columns to num
        for col in df_proc.columns:
            if col in ['Time', 'Velocity', 'Target Velocity', 'Acceleration', 'Target Acceleration']:
                df_proc[col] = pd.to_numeric(df_proc[col], errors='coerce')
        
        return df_proc

    def get_current_offset_sec(self):
        return self.local_offsets[self.current_graph_idx] / 1000.0

    def generate_plots(self):
        if not self.actual_path:
            QMessageBox.warning(self, "Uyarı", "Lütfen Actual Data yükleyin.")
            return

        try:
            # Sadece actual excel'i için ilk 8 satırı atla (metadata) ve 0, 5, 6 indeksli sütunları oku
            df_actual_raw = pd.read_excel(
                self.actual_path,
                skiprows=8,
                usecols=[0, 5, 6]
            )
            # Okunan sütunların isimlerini iç işleyişte beklenen standart listeye çeviriyoruz
            df_actual_raw.columns = ['Time', 'Acceleration', 'Velocity']
            self.df_actual = self.process_data(df_actual_raw)
            
            # Formül gereksinimi kontrol et (Spul = V^2 / t)
            if 'Velocity' in self.df_actual.columns and 'Time' in self.df_actual.columns:
                self.df_actual['Spul'] = np.where(
                    (self.df_actual['Time'] != 0) & (self.df_actual['Time'].notna()), 
                    (self.df_actual['Velocity']**2) / self.df_actual['Time'], 
                    0
                )
            
            if self.target_path:
                df_target_raw = pd.read_excel(self.target_path)
                self.df_target = self.process_data(df_target_raw)
                
                if 'Target Velocity' in self.df_target.columns and 'Time' in self.df_target.columns:
                    self.df_target['Spul'] = np.where(
                        (self.df_target['Time'] != 0) & (self.df_target['Time'].notna()), 
                        (self.df_target['Target Velocity']**2) / self.df_target['Time'], 
                        0
                    )
            else:
                self.df_target = None

            self.draw_current_graph()

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Veri işlenirken bir hata oluştu:\n{str(e)}")

    def apply_offset_to_actual(self, offset_sec):
        # Zamanı offset_sec kadar sola kaydır, negatifleri kes
        df_plot = self.df_actual.copy()
        df_plot['Offset_Time'] = df_plot['Time'] - offset_sec
        return df_plot[df_plot['Offset_Time'] >= 0]

    def _cleanup_axes(self):
        self.ax.clear()
        if self.ax2 is not None:
            self.ax2.remove()
            self.ax2 = None
        self.ax_table.clear()
        self.ax_table.axis('off')

    def draw_current_graph(self):
        if self.df_actual is None:
            return
            
        offset_sec = self.get_current_offset_sec()
        df_plot = self.apply_offset_to_actual(offset_sec)
        
        self._cleanup_axes()
        
        idx = self.current_graph_idx
        
        if idx == 0:
            self._draw_spul(df_plot)
        elif idx == 1:
            self._draw_acc_vel(df_plot)
        elif idx == 2:
            self._draw_acc_target_acc(df_plot)
            
        self.figure.tight_layout()
        self.canvas.draw()

    def _draw_spul(self, df_plot):
        if 'Spul' not in df_plot.columns:
            return
            
        actual_color = '#FFD700'
        target_color = '#2a52be'
        
        self.ax.plot(df_plot['Offset_Time'].values, df_plot['Spul'].values, color=actual_color, linewidth=2, label="SPUL")
        max_actual_spul = df_plot['Spul'].max()
        idx_max = df_plot['Spul'].idxmax()
        if pd.isna(idx_max): max_actual_time_sec = 0
        else: max_actual_time_sec = df_plot.loc[idx_max, 'Offset_Time']
        
        self.ax.vlines(x=max_actual_time_sec, ymin=0, ymax=max_actual_spul, colors=actual_color, linestyles='--', linewidth=1, alpha=0.7)
        
        max_target_spul = "-"
        max_target_time_ms = "-"
        if self.df_target is not None and 'Spul' in self.df_target.columns:
            self.ax.plot(self.df_target['Time'].values, self.df_target['Spul'].values, color=target_color, linewidth=2, linestyle='-', label="Target Spul")
            max_target_spul = self.df_target['Spul'].max()
            t_idx = self.df_target['Spul'].idxmax()
            if not pd.isna(t_idx):
                max_target_time_sec = self.df_target.loc[t_idx, 'Time']
                max_target_time_ms = max_target_time_sec * 1000.0
                self.ax.vlines(x=max_target_time_sec, ymin=0, ymax=max_target_spul, colors=target_color, linestyles='--', linewidth=1, alpha=0.7)

        self.ax.set_xlabel("time [s]", labelpad=10)
        self.ax.set_ylabel("Spul [(m/s)²/s]")
        self.ax.legend(
            loc='upper center', 
            bbox_to_anchor=(0.5, -0.15), 
            ncol=2, 
            frameon=False, 
            fontsize=14, 
            handlelength=2.0
        )
        self.ax.grid(True)
        self.ax.set_xlim(left=0)
        self.ax.set_ylim(bottom=0)
        
        # Tablo
        actual_val_str = f"{max_actual_spul:.1f}  $m^2/s^3$   ({max_actual_time_sec*1000.0:.1f} ms)" if not pd.isna(max_actual_spul) else "-"
        target_val_str = f"{max_target_spul:.1f}  $m^2/s^3$   ({max_target_time_ms:.1f} ms)" if max_target_spul != "-" else "-"
        
        cell_text = [
            ["SPUL", actual_val_str, ""],
            ["Target Spul", target_val_str, ""]
        ]
        self._build_table(cell_text, "SPUL\nSpecific Accident Capability\n$f(t) = v^2 / t$")

    def _draw_acc_vel(self, df_plot):
        if 'Acceleration' not in df_plot.columns or 'Velocity' not in df_plot.columns:
            self.ax.text(0.5, 0.5, "Acceleration veya Velocity Sütunu Bulunamadı", ha='center', va='center')
            return
            
        acc_color = '#1f77b4' # Mavi
        vel_color = '#2ca02c' # Yeşil
        
        self.ax2 = self.ax.twinx()
        
        l1 = self.ax.plot(df_plot['Offset_Time'].values, df_plot['Acceleration'].values, color=acc_color, linewidth=2, label="Acceleration")
        l2 = self.ax2.plot(df_plot['Offset_Time'].values, df_plot['Velocity'].values, color=vel_color, linewidth=2, label="Velocity")
        
        max_acc = df_plot['Acceleration'].max()
        a_idx = df_plot['Acceleration'].idxmax()
        max_acc_t = df_plot.loc[a_idx, 'Offset_Time'] if not pd.isna(a_idx) else 0
        self.ax.vlines(x=max_acc_t, ymin=0, ymax=max_acc, colors=acc_color, linestyles='--', linewidth=1, alpha=0.7)
        
        max_vel = df_plot['Velocity'].max()
        v_idx = df_plot['Velocity'].idxmax()
        max_vel_t = df_plot.loc[v_idx, 'Offset_Time'] if not pd.isna(v_idx) else 0
        self.ax2.vlines(x=max_vel_t, ymin=0, ymax=max_vel, colors=vel_color, linestyles='--', linewidth=1, alpha=0.7)

        self.ax.set_xlabel("Time, (s)", labelpad=10)
        self.ax.set_ylabel("Acceleration, (m/s²)")
        self.ax2.set_ylabel("Velocity, (m/s)")
        
        # Legend (Aşağıda ortalanmış bir şekilde iki kutu)
        lines = l1 + l2
        labels = [l.get_label() for l in lines]
        self.ax.legend(
            lines, labels, 
            loc='upper center', 
            bbox_to_anchor=(0.5, -0.15), 
            ncol=2, 
            frameon=False, 
            fontsize=14, 
            handlelength=2.0
        )
        self.ax.grid(True, alpha=0.5)
        self.ax.set_xlim(left=0)

        # Tablo
        v_str = f"{max_vel:.2f} $m/s$ ({max_vel_t*1000.0:.1f} ms)"
        a_str = f"{max_acc:.2f} $m/s^2$     ({max_acc_t*1000.0:.1f} ms)"
        cell_text = [
            ["Sled Velocity", v_str, ""],
            ["Sled Acceleration", a_str, ""]
        ]
        self._build_table(cell_text, "Sled Acceleration and Velocity")

    def _draw_acc_target_acc(self, df_plot):
        if 'Acceleration' not in df_plot.columns:
            self.ax.text(0.5, 0.5, "Actual'da Acceleration Sütunu Bulunamadı", ha='center', va='center')
            return
            
        acc_color = '#1f77b4'
        target_pulse_color = '#c20078' # Magenta (Morumsı)
        
        l1 = self.ax.plot(df_plot['Offset_Time'].values, df_plot['Acceleration'].values, color=acc_color, linewidth=2, label="Acceleration")
        
        max_acc = df_plot['Acceleration'].max()
        a_idx = df_plot['Acceleration'].idxmax()
        max_acc_t = df_plot.loc[a_idx, 'Offset_Time'] if not pd.isna(a_idx) else 0
        self.ax.vlines(x=max_acc_t, ymin=0, ymax=max_acc, colors=acc_color, linestyles='--', linewidth=1, alpha=0.7)
        
        max_t_acc = "-"
        max_t_acc_t = "-"
        l2 = []
        if self.df_target is not None and 'Target Acceleration' in self.df_target.columns:
            l2 = self.ax.plot(self.df_target['Time'].values, self.df_target['Target Acceleration'].values, color=target_pulse_color, linewidth=2, label="Target Pulse")
            max_t_acc = self.df_target['Target Acceleration'].max()
            ta_idx = self.df_target['Target Acceleration'].idxmax()
            if not pd.isna(ta_idx):
                max_t_acc_t_sec = self.df_target.loc[ta_idx, 'Time']
                max_t_acc_t = max_t_acc_t_sec * 1000.0
                self.ax.vlines(x=max_t_acc_t_sec, ymin=0, ymax=max_t_acc, colors=target_pulse_color, linestyles='--', linewidth=1, alpha=0.7)
                
        self.ax.set_xlabel("Time, (s)", labelpad=10)
        self.ax.set_ylabel("Acceleration, (m/s²)")
        
        lines = l1 + l2
        labels = [l.get_label() for l in lines]
        self.ax.legend(
            lines, labels, 
            loc='upper center', 
            bbox_to_anchor=(0.5, -0.15), 
            ncol=2, 
            frameon=False, 
            fontsize=14, 
            handlelength=2.0
        )
        self.ax.grid(True, alpha=0.5)
        self.ax.set_xlim(left=0)
        
        # Tablo
        a_str = f"{max_acc:.2f} $m/s^2$     ({max_acc_t*1000.0:.1f} ms)"
        t_str = f"{max_t_acc:.2f} $m/s^2$     ({max_t_acc_t:.1f} ms)" if max_t_acc != "-" else "-"
        cell_text = [
            ["Sled Acceleration", a_str, ""],
            ["Target Acceleration", t_str, ""]
        ]
        self._build_table(cell_text, "Sled vs. Target Acceleration")

    def _build_table(self, cell_text, graph_name_text):
        col_labels = ["", "Max. Value", "Graph Name"]
        table = self.ax_table.table(cellText=cell_text, colLabels=col_labels, loc='center', cellLoc='center', bbox=[0, 0, 1, 1])
        table.auto_set_font_size(False)
        table.set_fontsize(10)
        
        for (row, col), cell in table.get_celld().items():
            cell.set_text_props(ha='center', va='center')
            if row == 0:
                cell.set_text_props(weight='bold', ha='center', va='center')
            
            if col == 2 and row == 2:
                cell.visible_edges = 'BRL'
            if col == 2 and row == 1:
                cell.visible_edges = 'TRL'
        
        self.ax_table.text(0.833, 0.333, graph_name_text, ha='center', va='center', fontsize=10, transform=self.ax_table.transAxes)

    def export_plots(self):
        save_dir = self.txt_export.text()
        if not os.path.exists(save_dir) or not os.path.isdir(save_dir):
            QMessageBox.warning(self, "Hata", "Geçersiz kayıt dizini.")
            return
            
        if self.df_actual is None:
            QMessageBox.warning(self, "Hata", "İşlenecek Actual Data yok!")
            return

        try:
            # Current duruma dokunmadan arkada 3 grafiği çizip kaydedeceğiz
            saved_idx = self.current_graph_idx
            
            names = ["Spul.png", "Acc_vs_Vel.png", "Acc_vs_Targetacc.png"]
            
            for i in range(3):
                self.current_graph_idx = i
                self.draw_current_graph()
                path = os.path.join(save_dir, names[i])
                self.figure.savefig(path, dpi=300, bbox_inches='tight')
                
            # Restore
            self.current_graph_idx = saved_idx
            self.update_graph_view()
            
            QMessageBox.information(self, "Başarılı", f"Tüm 3 grafik seçilen klasöre kaydedildi:\n{names[0]}, {names[1]}, {names[2]}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dışa aktarma hatası:\n{str(e)}")

    def generate_word_report(self):
        save_dir = self.txt_export.text()
        if not os.path.exists(save_dir) or not os.path.isdir(save_dir):
            QMessageBox.warning(self, "Hata", "Geçersiz dizin.")
            return
            
        if self.df_actual is None:
            QMessageBox.warning(self, "Hata", "İşlenecek Actual Data yok!")
            return

        template_path = os.path.join(save_dir, "Template.docx")
        if not os.path.exists(template_path):
            QMessageBox.warning(self, "Hata", f"Template.docx dosyası bulunamadı, aynı dizinde olmalı:\n{template_path}")
            return
            
        test_no_input = global_data.config.get("TEST_NO", "Belirtilmedi")
        
        # Sonek çıkart (Örneğin 2026/096 -> 096)
        suffix = test_no_input.split('/')[-1] if '/' in test_no_input else test_no_input
        out_filename = f"graphs_{suffix}.docx"
        
        try:
            import tempfile
            temp_dir = tempfile.mkdtemp()
            
            saved_idx = self.current_graph_idx
            paths = {}
            labels = ["Spul", "Acc_vs_Vel", "Acc_vs_Targetacc"]
            
            # 3 Grafiği temp olarak çizdirip hafızaya alalım
            for i in range(3):
                self.current_graph_idx = i
                self.draw_current_graph()
                path = os.path.join(temp_dir, f"{labels[i]}.png")
                self.figure.savefig(path, dpi=300, bbox_inches='tight')
                paths[labels[i]] = path
                
            self.current_graph_idx = saved_idx
            self.update_graph_view()
            
            # Render the DocxTemplate using jinja2 syntaxes
            doc = DocxTemplate(template_path)
            
            context = {
                "TEST_NO": global_data.config.get("TEST_NO", ""),
                "TEST_DATE": global_data.config.get("TEST_DATE", ""),
                "PROJECT": global_data.config.get("PROJECT", ""),
                "SPUL": InlineImage(doc, paths["Spul"], width=Mm(160)),
                "ACC_VEL": InlineImage(doc, paths["Acc_vs_Vel"], width=Mm(160)),
                "ACC_TARGET": InlineImage(doc, paths["Acc_vs_Targetacc"], width=Mm(160))
            }
            
            doc.render(context)
            
            # Change out_path to point to tempfiles directory
            root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            tempfiles_dir = os.path.join(root_dir, "tempfiles")
            if not os.path.exists(tempfiles_dir):
                os.makedirs(tempfiles_dir)
                
            final_out_path = os.path.join(tempfiles_dir, out_filename)
            doc.save(final_out_path)
            
            QMessageBox.information(self, "Başarılı", f"Word Raporu başarıyla oluşturuldu!\n\nDosya Yolu: {final_out_path}")
            
            # Kapat ve ana uygulamaya dön
            if self.main_window:
                self.close()
                self.main_window.show()
            
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Rapor oluşturulurken hata oluştu:\n{str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.callbacks) if hasattr(sys, 'callbacks') else QApplication(sys.argv)
    window = SledAnalyzerApp()
    window.show()
    sys.exit(app.exec_())
