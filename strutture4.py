import sys
import os
import json
import pandas as pd
import requests
import random
import time
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QTableWidget, QTableWidgetItem, QMessageBox,
    QPushButton, QComboBox, QLineEdit, QFrame, QSplashScreen, QProgressBar
)
from PyQt6.QtCore import Qt, QUrl, QTimer, QPropertyAnimation, QEasingCurve, QParallelAnimationGroup
from PyQt6.QtGui import QColor, QDesktopServices, QFont, QPalette, QPixmap, QPainter, QLinearGradient, QBrush, QIcon

# Configurazione
CONFIG = {
    "update_urls": [
        "https://raw.githubusercontent.com/mattygoi/strutture/main/latest.json",
        "https://file.garden/Z-hU1H4Shk27aYus/latest.json"
    ],
    "fallback_url": "https://file.garden/Z-hU1H4Shk27aYus/STRUTTURE.xlsx",
    "config_file": "app_config.json"
}

class CinematicLoadingScreen(QSplashScreen):
    def __init__(self):
        super().__init__(QPixmap(800, 500))
        self.setWindowTitle("Caricamento...")
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setFixedSize(800, 500)
        
        self.main_layout = QVBoxLayout()
        self.setLayout(self.main_layout)
        
        self.setStyleSheet("""
            background-color: #0f0f1a;
            color: #ffffff;
            font-family: 'Segoe UI';
        """)
        
        self.logo = QLabel("")
        self.logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.logo.setStyleSheet("font-size: 100px; margin-bottom: 30px;")
        self.main_layout.addWidget(self.logo)
        
        self.title = QLabel("M.G STRUTTURE")
        self.title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title.setStyleSheet("""
            font-size: 28px; 
            font-weight: bold; 
            margin-bottom: 40px;
            color: #3498db;
        """)
        self.main_layout.addWidget(self.title)
        
        self.progress_container = QWidget()
        self.progress_layout = QVBoxLayout()
        self.progress_container.setLayout(self.progress_layout)
        
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        self.progress.setStyleSheet("""
            QProgressBar {
                border: 2px solid #2c3e50;
                border-radius: 5px;
                height: 20px;
                background: rgba(44, 62, 80, 0.5);
            }
            QProgressBar::chunk {
                background: qlineargradient(
                    spread:pad, x1:0, y1:0.5, x2:1, y2:0.5, 
                    stop:0 #3498db, stop:0.5 #9b59b6, stop:1 #e74c3c
                );
                border-radius: 3px;
            }
        """)
        self.progress_layout.addWidget(self.progress)
        
        self.percent = QLabel("0%")
        self.percent.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.percent.setStyleSheet("""
            font-size: 16px; 
            font-weight: bold;
            color: #ecf0f1;
            margin-top: 5px;
        """)
        self.progress_layout.addWidget(self.percent)
        
        self.message = QLabel("Inizializzazione del sistema...")
        self.message.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.message.setStyleSheet("""
            font-size: 14px; 
            color: #bdc3c7;
            margin-top: 20px;
            min-height: 40px;
        """)
        self.progress_layout.addWidget(self.message)
        
        self.details = QLabel("")
        self.details.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.details.setStyleSheet("""
            font-size: 12px; 
            color: #7f8c8d;
            margin-top: 10px;
            font-style: italic;
        """)
        self.progress_layout.addWidget(self.details)
        
        self.main_layout.addWidget(self.progress_container)
        
        self.setup_animations()
        
        self.loading_phrases = [
            ("Ottimizzazione algoritmi...", "Analisi strutture in corso"),
            ("Caricamento database...", "Accesso ai record storici"),
            ("Calibrazione parametri...", "Sintonizzazione precisione"),
            ("Preparazione interfaccia...", "Design cinematico in caricamento"),
            ("Verifica connessioni...", "Sincronizzazione cloud"),
            ("Analisi prestazioni...", "Benchmarking sistema"),
            ("Inizializzazione moduli...", "Caricamento componenti essenziali"),
            ("Preparazione report...", "Generazione dashboard"),
            ("Ottimizzazione memoria...", "Allocazione risorse"),
            ("Caricamento completato!", "Pronto per l'analisi")
        ]
        
        self.message_timer = QTimer()
        self.message_timer.timeout.connect(self.update_loading_message)
        self.message_timer.start(2500)
        
        self.progress_timer = QTimer()
        self.progress_timer.timeout.connect(self.simulate_progress)
        self.progress_timer.start(50)
        
        self.current_progress = 0
        self.target_progress = 0
        self.show()
    
    def setup_animations(self):
        self.animation_group = QParallelAnimationGroup()
        
        logo_anim = QPropertyAnimation(self.logo, b"geometry")
        logo_anim.setDuration(2000)
        logo_anim.setKeyValueAt(0, self.logo.geometry())
        logo_anim.setKeyValueAt(0.3, self.logo.geometry().translated(0, -20))
        logo_anim.setKeyValueAt(0.7, self.logo.geometry().translated(0, 10))
        logo_anim.setKeyValueAt(1, self.logo.geometry())
        logo_anim.setLoopCount(-1)
        self.animation_group.addAnimation(logo_anim)
        
        title_anim = QPropertyAnimation(self.title, b"windowOpacity")
        title_anim.setDuration(3000)
        title_anim.setKeyValueAt(0, 1)
        title_anim.setKeyValueAt(0.3, 0.7)
        title_anim.setKeyValueAt(0.6, 1)
        title_anim.setKeyValueAt(0.9, 0.8)
        title_anim.setKeyValueAt(1, 1)
        title_anim.setLoopCount(-1)
        self.animation_group.addAnimation(title_anim)
        
        self.animation_group.start()
    
    def update_loading_message(self):
        if self.target_progress >= 100:
            return
            
        phase = min(int(self.target_progress / 10), 9)
        msg, detail = self.loading_phrases[phase]
        self.message.setText(msg)
        self.details.setText(detail)
        
        if self.target_progress < 100:
            increment = random.uniform(6, 10)
            self.target_progress = min(self.target_progress + increment, 100)
    
    def simulate_progress(self):
        if self.current_progress < self.target_progress:
            step = max(0.5, min(3, (self.target_progress - self.current_progress) / 3))
            self.current_progress += step
            
            if self.target_progress - self.current_progress < 5:
                self.current_progress += 0.1
            
            self.progress.setValue(int(self.current_progress))
            self.percent.setText(f"{int(self.current_progress)}%")
            
            if self.current_progress >= 100:
                self.progress_timer.stop()
                self.message_timer.stop()
                self.message.setText("Caricamento completato!")
                self.details.setText("Pronto per l'analisi")
                self.title.setStyleSheet("font-size: 28px; font-weight: bold; color: #2ecc71;")
                self.logo.setText("🎿")
    
    def paintEvent(self, event):
        painter = QPainter(self)
        gradient = QLinearGradient(0, 0, 0, self.height())
        gradient.setColorAt(0, QColor(15, 15, 26))
        gradient.setColorAt(1, QColor(30, 39, 46))
        painter.fillRect(self.rect(), QBrush(gradient))
        
        highlight = QLinearGradient(0, 0, self.width(), 0)
        highlight.setColorAt(0, QColor(52, 152, 219, 30))
        highlight.setColorAt(0.5, Qt.GlobalColor.transparent)
        painter.fillRect(self.rect(), QBrush(highlight))
        
        super().paintEvent(event)

class ExcelViewer(QWidget):
    def __init__(self, splash):
        super().__init__()
        self.splash = splash
        self.current_url = None
        self.setup_paths()
        self.setup_ui()
        self.load_data()

    def setup_paths(self):
        self.app_data_dir = os.path.join(os.getenv('APPDATA'), "Strutture_XcSkiing")
        self.excel_path = os.path.join(self.app_data_dir, "STRUTTURE.xlsx")
        self.config_path = os.path.join(self.app_data_dir, CONFIG["config_file"])
        
        os.makedirs(self.app_data_dir, exist_ok=True)
        self.load_config()
        self.update_excel_url()
        
        if not self.verify_excel_file():
            if not self.download_excel_file():
                QMessageBox.critical(self, "Errore Critico", 
                    "Impossibile ottenere il file Excel necessario.\n"
                    "Assicurati di essere connesso a internet e riprova.\n"
                    f"Ultimo URL provato: {self.current_url}")
                sys.exit(1)

    def load_config(self):
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r') as f:
                    config = json.load(f)
                    self.current_url = config.get("last_working_url", CONFIG["fallback_url"])
            else:
                self.current_url = CONFIG["fallback_url"]
                self.save_config()
        except Exception:
            self.current_url = CONFIG["fallback_url"]

    def save_config(self):
        try:
            with open(self.config_path, 'w') as f:
                json.dump({"last_working_url": self.current_url}, f)
        except Exception:
            pass

    def update_excel_url(self):
        try:
            for update_url in CONFIG["update_urls"]:
                try:
                    response = requests.get(update_url, timeout=5)
                    if response.status_code == 200:
                        data = response.json()
                        new_url = data.get("excel_url")
                        if new_url and new_url != self.current_url:
                            self.current_url = new_url
                            self.save_config()
                            QMessageBox.information(self, "Aggiornamento", 
                                f"URL del database aggiornato!\n")
                        break
                except Exception:
                    continue
        except Exception:
            pass

    def verify_excel_file(self):
        if os.path.exists(self.excel_path):
            try:
                with pd.ExcelFile(self.excel_path) as xls:
                    if "Foglio1" not in xls.sheet_names:
                        return False
                return True
            except Exception as e:
                print(f"Errore verifica file Excel: {str(e)}")
                os.remove(self.excel_path)
                return False
        return False

    def download_excel_file(self):
        temp_path = self.excel_path + ".tmp"
        try:
            response = requests.get(self.current_url, stream=True, timeout=30)
            response.raise_for_status()
            
            with open(temp_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            
            # Verifica che il file scaricato sia valido
            try:
                with pd.ExcelFile(temp_path) as xls:
                    if "Foglio1" not in xls.sheet_names:
                        raise ValueError("Foglio mancante")
            except Exception as e:
                raise ValueError(f"File scaricato non valido: {str(e)}")
            
            if os.path.exists(self.excel_path):
                os.remove(self.excel_path)
            os.rename(temp_path, self.excel_path)
            
            return True
            
        except Exception as e:
            print(f"Errore download: {str(e)}")
            if os.path.exists(temp_path):
                os.remove(temp_path)
                
            if self.current_url != CONFIG["fallback_url"]:
                self.current_url = CONFIG["fallback_url"]
                return self.download_excel_file()
                
            QMessageBox.critical(self, "Errore", f"Download fallito: {str(e)}")
            return False

# ... (tutto il codice precedente rimane invariato fino alla funzione setup_ui)

    def setup_ui(self):
        self.setWindowTitle("📊 Strutture Sci - Analisi Dati")
        self.setGeometry(100, 100, 1400, 800)
        
        self.setStyleSheet("""
            QPushButton {
                background-color: #5c6bc0;
                color: white;
                border: none;
                padding: 6px 12px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #3949ab;
            }
        """)

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        title = QLabel("📊 Analisi Strutture Sci")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        main_layout.addWidget(title)

        filter_layout = QGridLayout()
        filter_layout.setVerticalSpacing(10)
        filter_layout.setHorizontalSpacing(15)

        self.luogo_filter = QLineEdit()
        self.luogo_filter.setPlaceholderText("Es: Dobbiaco o Livigno")
        
        self.tipo_evento = QComboBox()
        self.tipo_evento.addItems(["Tutti", "TEST", "GARA"])
        
        self.meteo_filter = QLineEdit()
        self.meteo_filter.setPlaceholderText("Es: soleggiato o nuvoloso")
        
        self.temp_aria = QLineEdit()
        self.temp_aria.setPlaceholderText("Es: -5 o 2")
        
        self.temp_neve = QLineEdit()
        self.temp_neve.setPlaceholderText("Es: -3 o 1")
        
        self.tipo_neve = QLineEdit()
        self.tipo_neve.setPlaceholderText("Es: farinosa o umida")
        
        self.umidita = QLineEdit()
        self.umidita.setPlaceholderText("Es: 30 o 70")
        
        self.scelte_filter = QComboBox()
        self.scelte_filter.addItems(["Tutte", "Prima scelta", "Seconda scelta", "Terza scelta"])

        filter_layout.addWidget(QLabel("📍 Luogo:"), 0, 0)
        filter_layout.addWidget(self.luogo_filter, 0, 1)
        filter_layout.addWidget(QLabel("🔧 Tipo Evento:"), 1, 0)
        filter_layout.addWidget(self.tipo_evento, 1, 1)
        filter_layout.addWidget(QLabel("⛅ Condizioni Meteo:"), 2, 0)
        filter_layout.addWidget(self.meteo_filter, 2, 1)
        filter_layout.addWidget(QLabel("🌡️ Temp. Aria:"), 0, 2)
        filter_layout.addWidget(self.temp_aria, 0, 3)
        filter_layout.addWidget(QLabel("🌡️ Temp. Neve:"), 1, 2)
        filter_layout.addWidget(self.temp_neve, 1, 3)
        filter_layout.addWidget(QLabel("❄️ Tipo Neve:"), 2, 2)
        filter_layout.addWidget(self.tipo_neve, 2, 3)
        filter_layout.addWidget(QLabel("💧 Umidità (%):"), 3, 0)
        filter_layout.addWidget(self.umidita, 3, 1)
        filter_layout.addWidget(QLabel("🏆 Scelte:"), 3, 2)
        filter_layout.addWidget(self.scelte_filter, 3, 3)

        main_layout.addLayout(filter_layout)

        btn_layout = QHBoxLayout()
        self.filter_btn = QPushButton("🔍 Applica Filtri")
        self.filter_btn.clicked.connect(self.load_data)
        self.reset_btn = QPushButton("🔄 Reset")
        self.reset_btn.clicked.connect(self.reset_filters)
        btn_layout.addWidget(self.filter_btn)
        btn_layout.addWidget(self.reset_btn)
        main_layout.addLayout(btn_layout)

        self.table = QTableWidget()
        self.table.setSortingEnabled(True)
        main_layout.addWidget(self.table)

        self.status_label = QLabel()
        self.status_label.setStyleSheet("color: #666;")
        main_layout.addWidget(self.status_label)

        self.setLayout(main_layout)

    def reset_filters(self):
        self.luogo_filter.clear()
        self.tipo_evento.setCurrentIndex(0)
        self.temp_aria.clear()
        self.temp_neve.clear()
        self.tipo_neve.clear()
        self.umidita.clear()
        self.meteo_filter.clear()
        self.scelte_filter.setCurrentIndex(0)
        self.load_data()

def load_data(self):
    if not os.path.exists(self.excel_path):
        QMessageBox.critical(self, "Errore", "Il file Excel non esiste!")
        return
        
    try:
        df = pd.read_excel(self.excel_path, sheet_name="Foglio1", engine='openpyxl')
        df = df.fillna("")
        
        # Processa i dati
        groups = []
        current_group = []
        
        for _, row in df.iterrows():
            if all(str(value).strip() == "" for value in row.values):
                if current_group:
                    groups.append(pd.DataFrame(current_group))
                    current_group = []
            else:
                current_group.append(row)
        
        if current_group:
            groups.append(pd.DataFrame(current_group))
        
        # Applica i filtri
        filtered_groups = []
        
        for group in groups:
            include_group = True
            
            # Filtro luogo
            if self.luogo_filter.text():
                luogo_text = self.luogo_filter.text().strip().lower()
                if not any(luogo_text in str(row.get("LUOGO", "")).lower() for _, row in group.iterrows()):
                    include_group = False
            
            # Filtro tipo evento
            if include_group and self.tipo_evento.currentText() != "Tutti":
                tipo_text = self.tipo_evento.currentText().upper()
                if not any(tipo_text in str(row.get("TEST o GARA", "")).strip().upper() for _, row in group.iterrows()):
                    include_group = False
            
            # Filtro meteo
            if include_group and self.meteo_filter.text():
                meteo_text = self.meteo_filter.text().strip().lower()
                if not any(meteo_text in str(row.get("CONDIZIONI METEO E VENTO", "")).lower() for _, row in group.iterrows()):
                    include_group = False
            
            # Filtro temperatura aria (gestisce valori negativi e decimali)
            if include_group and self.temp_aria.text():
                temp_text = self.temp_aria.text().strip().replace(',', '.')
                try:
                    temp_value = float(temp_text)
                    temp_found = False
                    for _, row in group.iterrows():
                        for col in ["TEMP. ARIA INIZIO", "TEMP. ARIA FINE"]:
                            cell_value = str(row.get(col, "")).strip()
                            if cell_value:
                                try:
                                    cell_float = float(cell_value.replace(',', '.'))
                                    if abs(cell_float - temp_value) < 0.1:  # Tolleranza per decimali
                                        temp_found = True
                                        break
                                except ValueError:
                                    continue
                        if temp_found:
                            break
                    if not temp_found:
                        include_group = False
                except ValueError:
                    include_group = False
            
            # Filtro temperatura neve (stessa logica di temperatura aria)
            if include_group and self.temp_neve.text():
                tempn_text = self.temp_neve.text().strip().replace(',', '.')
                try:
                    tempn_value = float(tempn_text)
                    tempn_found = False
                    for _, row in group.iterrows():
                        for col in ["TEMP. NEVE INIZIO", "TEMP. NEVE FINE"]:
                            cell_value = str(row.get(col, "")).strip()
                            if cell_value:
                                try:
                                    cell_float = float(cell_value.replace(',', '.'))
                                    if abs(cell_float - tempn_value) < 0.1:
                                        tempn_found = True
                                        break
                                except ValueError:
                                    continue
                        if tempn_found:
                            break
                    if not tempn_found:
                        include_group = False
                except ValueError:
                    include_group = False
            
            # Filtro tipo neve
            if include_group and self.tipo_neve.text():
                tipon_text = self.tipo_neve.text().strip().lower()
                if not any(tipon_text in str(row.get("TIPO NEVE", "")).lower() for _, row in group.iterrows()):
                    include_group = False
            
            # Filtro umidità (gestisce valori percentuali e decimali)
            if include_group and self.umidita.text():
                umid_text = self.umidita.text().strip().replace(',', '.')
                try:
                    umid_value = float(umid_text)
                    umid_found = False
                    for _, row in group.iterrows():
                        for col in ["UMIDITA % INIZIO", "UMIDITA' % FINE"]:
                            cell_value = str(row.get(col, "")).strip().replace('%', '')
                            if cell_value:
                                try:
                                    cell_float = float(cell_value.replace(',', '.'))
                                    if abs(cell_float - umid_value) < 0.1:
                                        umid_found = True
                                        break
                                except ValueError:
                                    continue
                        if umid_found:
                            break
                    if not umid_found:
                        include_group = False
                except ValueError:
                    include_group = False
            
            # Filtro scelte (ricerca più flessibile)
            if include_group and self.scelte_filter.currentText() != "Tutte":
                scelta_text = self.scelte_filter.currentText().upper()
                scelta_found = False
                for _, row in group.iterrows():
                    considerazioni = str(row.get("CONSIDERAZIONE POST GARA o TEST", "")).upper()
                    # Cerca varie forme di "prima scelta", "seconda scelta", ecc.
                    if scelta_text == "PRIMA SCELTA" and (
                        "PRIMA SCELTA" in considerazioni or 
                        "PRIMO" in considerazioni or 
                        "MIGLIORE" in considerazioni or 
                        "1°" in considerazioni
                    ):
                        scelta_found = True
                        break
                    elif scelta_text == "SECONDA SCELTA" and (
                        "SECONDA SCELTA" in considerazioni or 
                        "SECONDO" in considerazioni or 
                        "2°" in considerazioni
                    ):
                        scelta_found = True
                        break
                    elif scelta_text == "TERZA SCELTA" and (
                        "TERZA SCELTA" in considerazioni or 
                        "TERZO" in considerazioni or 
                        "3°" in considerazioni
                    ):
                        scelta_found = True
                        break
                if not scelta_found:
                    include_group = False
            
            if include_group:
                filtered_groups.append(group)
        
        filtered_df = pd.concat(filtered_groups) if filtered_groups else pd.DataFrame(columns=df.columns)
        
        # Popola la tabella
        self.table.setRowCount(len(filtered_df))
        self.table.setColumnCount(len(filtered_df.columns))
        self.table.setHorizontalHeaderLabels(filtered_df.columns)
        
        for row in range(len(filtered_df)):
            for col in range(len(filtered_df.columns)):
                val = filtered_df.iat[row, col]
                item = QTableWidgetItem(str(val))
                
                # Evidenziazione basata sulle scelte (logica più robusta)
                considerazioni = str(filtered_df.at[row, "CONSIDERAZIONE POST GARA o TEST"]).upper()
                if "PRIMA SCELTA" in considerazioni or "PRIMO" in considerazioni or "MIGLIORE" in considerazioni or "1°" in considerazioni:
                    item.setBackground(QColor(200, 255, 200))  # Verde chiaro
                elif "SECONDA SCELTA" in considerazioni or "SECONDO" in considerazioni or "2°" in considerazioni:
                    item.setBackground(QColor(255, 255, 150))  # Giallo chiaro
                elif "TERZA SCELTA" in considerazioni or "TERZO" in considerazioni or "3°" in considerazioni:
                    item.setBackground(QColor(255, 200, 150))  # Arancione chiaro
                
                self.table.setItem(row, col, item)
        
        self.table.resizeColumnsToContents()
        
        total_records = len(df)
        filtered_records = len(filtered_df)
        self.status_label.setText(
            f"🔹 Record trovati: {filtered_records} | "
            f"🔸 Record totali: {total_records} | "
            f"📌 Developed By: @mattygoi"
        )
        
    except Exception as e:
        QMessageBox.critical(self, "Errore", 
            f"Errore durante il caricamento:\n{str(e)}\n\n"
            f"Assicurarsi che il file Excel sia chiuso e riprovare.")
        print(f"Errore durante il caricamento dei dati: {str(e)}")

# ... (il resto del codice rimane invariato)
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    try:
        requests.get("https://www.google.com", timeout=5)
    except Exception:
        QMessageBox.warning(None, "Avviso", 
            "⚠️ Connessione internet non rilevata\n"
            "L'applicazione funzionerà in modalità offline se hai già il file.")
    
    splash = CinematicLoadingScreen()
    app.processEvents()
    
    while splash.current_progress < 100:
        time.sleep(0.1)
        app.processEvents()
    
    window = ExcelViewer(splash)
    window.show()
    splash.finish(window)
    
    sys.exit(app.exec())