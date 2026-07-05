import sys
import pandas as pd

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QPushButton,
    QLabel,
    QFileDialog,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
    QHBoxLayout
)

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor

from excel_handler import load_patients_from_excel
from scheduler import assign_patients, generate_15_day_schedule


class SpaSchedulerApp(QWidget):

    def __init__(self):

        super().__init__()

        self.setWindowTitle("Spa Scheduler")
        self.setGeometry(100, 50, 1800, 900)

        self.schedule_data = []

        self.layout = QVBoxLayout()

        self.title = QLabel("SPA SCHEDULER SYSTEM")

        self.title.setStyleSheet("""
            QLabel {
                font-size: 28px;
                font-weight: bold;
                color: #0B6E6E;
                padding: 10px;
            }
        """)

        self.layout.addWidget(self.title)

        buttons_layout = QHBoxLayout()

        # IMPORT
        self.import_button = QPushButton("Import Excel") 
        self.import_button.clicked.connect(self.import_excel)
        buttons_layout.addWidget(self.import_button) 

        # EXPORT GLOBAL
        self.export_button = QPushButton(
            "Export Global Schedule"
        )
        self.export_button.clicked.connect(
            self.export_schedule
        )
        buttons_layout.addWidget(self.export_button)

        # SAVE EDITED
        self.save_edited_button = QPushButton(
            "Save Edited Table"
        )

        self.save_edited_button.clicked.connect(
            self.save_edited_table
        )

        buttons_layout.addWidget(
            self.save_edited_button
        )

        # EXPORT INDIVIDUAL
        self.export_individual_button = QPushButton(
            "Export Individual Schedules"
        )

        self.export_individual_button.clicked.connect(
            self.export_individual_schedules
        )

        buttons_layout.addWidget(
            self.export_individual_button
        )

        # EXPORT STATIONS
        self.export_stations_button = QPushButton(
            "Export Stations Schedules"
        )

        self.export_stations_button.clicked.connect(
            self.export_stations_schedules
        )

        buttons_layout.addWidget(
            self.export_stations_button
        )

        # CLEAR
        self.clear_button = QPushButton(
            "Clear Current Schedule"
        )

        self.clear_button.clicked.connect(
            self.clear_schedule
        )

        buttons_layout.addWidget(
            self.clear_button
        )

        # EXIT
        self.exit_button = QPushButton("Exit")

        self.exit_button.clicked.connect(
            self.close
        )

        buttons_layout.addWidget(
            self.exit_button
        )

        self.layout.addLayout(buttons_layout)

        # TABLE
        self.table = QTableWidget()

        self.layout.addWidget(self.table)

        self.setLayout(self.layout)

        # STYLE
        self.setStyleSheet("""

            QWidget {
                background-color: #F4FBFB;
                font-size: 12px;
            }

            QPushButton {

                background-color: #7ED6D4;
                color: black;

                border-radius: 10px;

                padding: 10px;

                font-weight: bold;
            }

            QPushButton:hover {
                background-color: #5CC5C2;
            }

            QTableWidget {

                background-color: white;

                gridline-color: #BFDCDC;

                font-size: 11px;
            }

            QHeaderView::section {

                background-color: #B8E8E8;

                padding: 8px;

                border: 1px solid #9ED4D4;

                font-weight: bold;
            }

        """)

    # =====================================================
    # IMPORT
    # =====================================================

    def import_excel(self):

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Excel File",
            "",
            "Excel Files (*.xlsx *.xls)"
        )

        if not file_path:
            return

        try:

            patients = load_patients_from_excel(
                file_path
            )

            patients = assign_patients(
                patients
            )

            self.schedule_data = generate_15_day_schedule(
                patients
            )

            self.display_schedule(
                self.schedule_data
            )

            QMessageBox.information(
                self,
                "Success",
                "Schedule Generated Successfully"
            )

        except Exception as e:

            QMessageBox.critical(
                self,
                "Error",
                str(e)
            )

    # =====================================================
    # DISPLAY TABLE
    # =====================================================

    def display_schedule(self, data):

        grouped_data = {}

        for row in data:

            patient = row["Patient"]

            if patient not in grouped_data:
                grouped_data[patient] = {}

            grouped_data[patient][row["Day"]] = row

        headers = ["Patient"]

        for day in range(1, 16):
            headers.append(f"Day {day}")

        self.table.clear()

        self.table.setColumnCount(len(headers))

        self.table.setHorizontalHeaderLabels(headers)

        self.table.setRowCount(len(grouped_data))

        for row_index, (patient, schedules) in enumerate(grouped_data.items()):

            patient_item = QTableWidgetItem(patient)

            patient_item.setBackground(
                QColor("#DDF5F5")
            )

            self.table.setItem(
                row_index,
                0,
                patient_item
            )

            for day in range(1, 16):

                if day in schedules:

                    info = schedules[day]

                    text = (
                        f'Session: {info["Session"]}\n\n'
                        f'Group: {info["Group"]}\n'
                        f'Subgroup: {info["Subgroup"]}\n\n'
                        f'{info["Soin"]}\n\n'
                        f'{info["Stations"]}'
                    )

                    item = QTableWidgetItem(text)

                    item.setTextAlignment(
                        Qt.AlignTop | Qt.AlignLeft
                    )

                    group = info["Group"]

                    if group == "G1":
                        item.setBackground(QColor("#FFF4CC"))

                    elif group == "G2":
                        item.setBackground(QColor("#DFFFD8"))

                    elif group == "G3":
                        item.setBackground(QColor("#D9ECFF"))

                    elif group == "G4":
                        item.setBackground(QColor("#FFDDE2"))

                    self.table.setItem(
                        row_index,
                        day,
                        item
                    )

        self.table.setColumnWidth(0, 220)

        for col in range(1, 16):
            self.table.setColumnWidth(col, 300)

        for row in range(self.table.rowCount()):
            self.table.setRowHeight(row, 190)

    # =====================================================
    # SAVE EDITED TABLE
    # =====================================================

    def save_edited_table(self):

        if self.table.rowCount() == 0:

            QMessageBox.warning(
                self,
                "Warning",
                "No Table Data"
            )

            return

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Edited Schedule",
            "edited_schedule.xlsx",
            "Excel Files (*.xlsx)"
        )

        if not save_path:
            return

        try:

            rows_data = []

            for row in range(self.table.rowCount()):

                patient_item = self.table.item(row, 0)

                if patient_item is None:
                    continue

                patient_name = patient_item.text()

                for day in range(1, 16):

                    cell_item = self.table.item(row, day)

                    if cell_item is None:
                        continue

                    text = cell_item.text().strip()

                    if not text:
                        continue

                    rows_data.append({

                        "Patient": patient_name,

                        "Day": day,

                        "Details": text
                    })

            df = pd.DataFrame(rows_data)

            df.to_excel(
                save_path,
                index=False
            )

            QMessageBox.information(
                self,
                "Success",
                "Edited table saved successfully"
            )

        except Exception as e:

            QMessageBox.critical(
                self,
                "Error",
                str(e)
            )

    # =====================================================
    # CLEAR
    # =====================================================

    def clear_schedule(self):

        self.schedule_data = []

        self.table.clear()

        self.table.setRowCount(0)

        self.table.setColumnCount(0)

    # =====================================================
    # EXPORT GLOBAL
    # =====================================================

    def export_schedule(self):

        if not self.schedule_data:
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save File",
            "schedule.xlsx",
            "Excel Files (*.xlsx)"
        )

        if not save_path:
            return

        df = pd.DataFrame(self.schedule_data)

        df.to_excel(save_path, index=False)

        QMessageBox.information(
            self,
            "Success",
            "Exported Successfully"
        )

    # =====================================================
    # EXPORT INDIVIDUAL
    # =====================================================

    def export_individual_schedules(self):

        if not self.schedule_data:
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save File",
            "individual_schedules.xlsx",
            "Excel Files (*.xlsx)"
        )

        if not save_path:
            return

        grouped_data = {}

        for row in self.schedule_data:

            patient = row["Patient"]

            if patient not in grouped_data:
                grouped_data[patient] = []

            grouped_data[patient].append(row)

        with pd.ExcelWriter(
            save_path,
            engine="openpyxl"
        ) as writer:

            for patient, schedules in grouped_data.items():

                df = pd.DataFrame(schedules)

                sheet_name = patient[:31]

                df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False
                )

        QMessageBox.information(
            self,
            "Success",
            "Individual schedules exported"
        )

    # =====================================================
    # EXPORT STATIONS
    # =====================================================

    def export_stations_schedules(self):

        if not self.schedule_data:
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save File",
            "stations_schedules.xlsx",
            "Excel Files (*.xlsx)"
        )

        if not save_path:
            return

        station_data = {}

        for row in self.schedule_data:

            stations = row["Stations"].split(" -> ")

            for station in stations:

                if station not in station_data:
                    station_data[station] = []

                station_data[station].append({

                    "Day": row["Day"],

                    "Patient": row["Patient"],

                    "Session": row["Session"],

                    "Group": row["Group"],

                    "Subgroup": row["Subgroup"],

                    "Soin": row["Soin"]
                })

        with pd.ExcelWriter(
            save_path,
            engine="openpyxl"
        ) as writer:

            for station, rows in station_data.items():

                df = pd.DataFrame(rows)

                sheet_name = station[:31]

                df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False
                )

        QMessageBox.information(
            self,
            "Success",
            "Stations schedules exported"
        )


app = QApplication(sys.argv)

window = SpaSchedulerApp()

window.show()

sys.exit(app.exec())
