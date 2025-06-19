from PySide6.QtCore import (
    QSettings,
    QTimer,
    QAbstractTableModel,
    QModelIndex,
    QSortFilterProxyModel,
    Qt

)
from PySide6.QtWidgets import (
    QApplication,
    QLineEdit,
    QMainWindow,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QMessageBox,
    QHBoxLayout,
    QLabel,
    QTableView,
    QStyledItemDelegate,
    QComboBox,
    QHeaderView,
    QTreeView,
)
from PySide6 import QtGui
from PySide6.QtGui import (
    QStandardItemModel,
    QStandardItem
)
import sys
import csv
from pycomm3 import LogixDriver
import qdarktheme
from dataclasses import dataclass
from datetime import datetime, timedelta

# gets all tags from the PLC and puts them into a dictionary. Tag children will be in keys.
# Attributes are: data_type: str, dimensions: list[x,x,x], structure: bool
def get_tags_from_plc(ip="192.168.1.2"):
    tag_list = {}

    try:
        with LogixDriver(ip) as plc:
            data = plc.tags_json

        for tag_name, tag_info in data.items():
            tag_data_type = tag_info['data_type']
            tag_type = tag_info['tag_type']
            tag_dimensions = tag_info.get('dimensions', [0, 0, 0])

            if tag_type == 'atomic':
                tag_list[tag_name] = {
                    'data_type': tag_data_type,
                    'dimensions': tag_dimensions,
                    'structure': False
                }
            elif tag_type == 'struct':
                # Store the parent structure
                if tag_data_type['name'] == 'STRING':
                    tag_list[tag_name] = {
                        'data_type': tag_data_type['name'],
                        'dimensions': tag_dimensions,
                        'structure': False
                    }
                else:
                    tag_list[tag_name] = {
                        'data_type': tag_data_type['name'],
                        'dimensions': tag_dimensions,
                        'structure': True
                    }
                # Recursively store children
                if tag_data_type['name'] != 'STRING':
                    tag_list = extract_child_data_types(
                        tag_data_type['internal_tags'], tag_list, tag_name)
        return tag_list
    except Exception as e:
        print(f"Error in get_tags_from_plc function: {e}")
        return None


def extract_child_data_types(structure, array, name):
    for child_name, child_info in structure.items():
        child_data_type = child_info['data_type']
        child_tag_type = child_info['tag_type']
        child_array_length = child_info.get('array', 0)

        if child_name.startswith('_') or child_name.startswith('ZZZZZZZZZZ'):
            continue

        full_tag_name = f'{name}.{child_name}'
        if child_tag_type == 'atomic':
            array[full_tag_name] = {
                'data_type': child_data_type,
                'dimensions': [child_array_length, 0, 0],
                'structure': False
            }
        elif child_tag_type == 'struct':
            # Store the structure itself
            if child_data_type['name'] == 'STRING':
                array[full_tag_name] = {
                    'data_type': child_data_type['name'],
                    'dimensions': [child_array_length, 0, 0],
                    'structure': False
                }
            else:
                array[full_tag_name] = {
                    'data_type': child_data_type['name'],
                    'dimensions': [child_array_length, 0, 0],
                    'structure': True
                }
            # Recursively store children
            if child_data_type['name'] != 'STRING':
                array = extract_child_data_types(
                    child_data_type['internal_tags'], array, full_tag_name)

    return array

def format_dimension_label(name: str, dims: list[int]) -> str:
    if dims == [0, 0, 0]:
        return name
    dim_parts = [str(d) for d in dims if d > 0]
    return f"{name}[{','.join(dim_parts)}]"

def build_tag_tree_model(tag_dict: dict) -> QStandardItemModel:
    model = QStandardItemModel()
    model.setHorizontalHeaderLabels(["Tag Name"])

    root = model.invisibleRootItem()
    node_map = {}

    for full_path, meta in tag_dict.items():
        parts = full_path.split(".")
        current_path = ""
        parent = root

        for i, part in enumerate(parts):
            current_path = f"{current_path}.{part}" if current_path else part
            is_leaf = (i == len(parts) - 1)
            label = format_dimension_label(part, meta["dimensions"]) if is_leaf else part
            if current_path not in node_map:
                item = QStandardItem(label)
                item.setEditable(False)

                if is_leaf:
                    item.setData((full_path, meta["dimensions"]), role=Qt.UserRole)
                    item.setToolTip(f'{full_path}({meta["data_type"]})')

                parent.appendRow(item)
                node_map[current_path] = item

            parent = node_map[current_path]

    return model

def round_to_nearest_second(dt: datetime) -> datetime:
    """
    Rounds a datetime object to the nearest second.
    """
    # Check if microseconds are 500,000 or more
    if dt.microsecond >= 500000:
        # Round up: add one second and set microseconds to 0
        dt_rounded = dt.replace(microsecond=0) + timedelta(seconds=1)
    else:
        # Round down: set microseconds to 0
        dt_rounded = dt.replace(microsecond=0)
    return dt_rounded


class TagFilterProxyModel(QSortFilterProxyModel):
    def filterAcceptsRow(self, source_row, source_parent):
        model = self.sourceModel()
        index = model.index(source_row, 0, source_parent)

        pattern = self.filterRegularExpression()

        if pattern.match(model.data(index)).hasMatch():
            return True
        
        for i in range(model.rowCount(index)):
            if self.filterAcceptsRow(i, index):
                return True
            
        if source_parent.isValid():
            parent_text = model.data(source_parent)
            if pattern.match(parent_text).hasMatch():
                return True
            
        return False

@dataclass
class DowntimeEvent:
    start_time: datetime
    end_time: datetime
    cause: str
    nmr: bool

    @property
    def duration(self) -> timedelta:
        return self.end_time - self.start_time


class ComboBoxDelegate(QStyledItemDelegate):
    def __init__(self, items, parent=None):
        super().__init__(parent)
        self.items = items

    def createEditor(self, parent, option, index):
        combo = QComboBox(parent)
        combo.addItems(self.items)
        return combo
    
    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.DisplayRole)
        i = editor.findText(value)
        if i >= 0:
            editor.setCurrentIndex(i)

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), Qt.EditRole)


class DowntimeModel(QAbstractTableModel):
    def __init__(self, events, run_start_time, footer_update_callback=None):
        super().__init__()
        self.events = events
        self.run_start_time = run_start_time
        self.footer_update_callback = footer_update_callback
        

    def rowCount(self, parent=QModelIndex()):
        return len(self.events)
    
    def columnCount(self, parent=QModelIndex()):
        return 5
    
    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or role != Qt.DisplayRole:
            return None
        
        def format_dt(td: timedelta) -> str:
            total_seconds = int(td.total_seconds())
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            seconds = total_seconds % 60

            return f'{hours}:{minutes:02}:{seconds:02}'
        
        event = self.events[index.row()]
        if index.column() == 0:
            return format_dt(event.start_time - self.run_start_time)
        elif index.column() == 1:
            return format_dt(event.end_time - self.run_start_time)
        elif index.column() == 2:
            return format_dt(event.duration)
        elif index.column() == 3:
            return event.cause
        elif index.column() == 4:
            return 'NMR' if event.nmr else 'MR'
        
        return None

    def setData(self, index, value, role=Qt.EditRole):
        if role != Qt.EditRole or not index.isValid():
            return False
        event = self.events[index.row()]

        col = index.column()

        def parse_td(text: str):
            try:
                parts = list(map(int,text.strip().split(":")))
                if len(parts) == 3:
                    return timedelta(hours=parts[0], minutes=parts[1], seconds=parts[2])
                elif len(parts) == 2:
                    return timedelta(minutes=parts[0], seconds=parts[1])
                elif len(parts) == 1:
                    return timedelta(seconds=parts[0])
                else:
                    return None
            except:
                return None
        
        td = parse_td(value)


        if col == 0:
            if td is None:
                return False
            else:
                new_time = self.run_start_time + td
            event.start_time = new_time
            self.dataChanged.emit(index, self.index(index.row(), 2))
        elif col == 1:
            if td is None:
                return False
            else:
                new_time = self.run_start_time + td
            event.end_time = new_time
            self.dataChanged.emit(index, self.index(index.row(), 2))
        elif col == 3:
            event.cause = value
            
            self.dataChanged.emit(index, index)
        elif col == 4:
            if value == "NMR":
                event.nmr = True
            elif value == "MR":
                event.nmr = False
            self.dataChanged.emit(index, index)
            self.footer_update_callback()

        return True
    
    def flags(self, index):
        base_flags = super().flags(index)
        if index.column() in (0,1,3,4,5):
            return base_flags | Qt.ItemIsEditable
        
        return base_flags

    
    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole or orientation != Qt.Horizontal:
            return None
        
        return ['Start Time', 'End Time', 'Duration', 'Cause', 'NMR/MR'][section]


class MainWindow(QMainWindow):
    
    def __init__(self):
        super(MainWindow, self).__init__()
        self.settings = QSettings("PM Development", "OEE Tracker")

        self.setWindowTitle("OEE Tracker")

        self.events = []  # List to hold downtime events
        self.temp_start_time = None
        self.recording = False
        self.run_start_time = None
        self.paused_duration = timedelta()
        self.pause_start = None
        self.is_paused = False
        self.proxy_model = None

        # Buttons
        self.start_button = QPushButton("Start")
        self.pause_button = QPushButton("Pause")
        self.stop_button = QPushButton("Stop")
        self.downtime_button = QPushButton("Record Downtime")
        self.export_downtime_button = QPushButton("Export Downtime Events")
        self.get_tags_button = QPushButton("Get Tag List")
        self.add_fault_tag_button = QPushButton('Add Selected')

        # Labels
        self.quality = QLabel("Quality: 0%")
        self.availability = QLabel("Availability: 0%")
        self.performance = QLabel("Performance: 0%")
        self.oee_components = QLabel('Performance: N/A | Availability: N/A | Quality: N/A')
        self.oee = QLabel("OEE: 0%")
        self.ideal_rate_label = QLabel("Ideal Rate (seconds per part):")
        self.calculated_rate = QLabel("Calculated Rate: 0.00 seconds per part")
        self.time = QLabel('Run Not Started')
        self.footer_label = QLabel()

        # Line Edits
        self.ip_input = QLineEdit()
        self.ideal_rate = QLineEdit()
        self.rejects = QLineEdit()
        self.total_parts = QLineEdit()
        self.rejects_tag = QLineEdit()
        self.total_parts_tag = QLineEdit()
        self.tree_filter = QLineEdit()

        # Timers
        self.timer = QTimer()

        # Tree View
        self.tree = QTreeView()
        self.tree.setFixedHeight(250)
        self.tree.setSelectionMode(QTreeView.ExtendedSelection)

        # Table
        self.table_view = QTableView()

        nmr_delegate = ComboBoxDelegate(['NMR', "MR"])
        self.table_view.setItemDelegateForColumn(4, nmr_delegate)

        self.model = DowntimeModel(self.events, round_to_nearest_second(datetime.now()), self.update_footer)
        self.table_view.setModel(self.model)

        # Set up table header sizing
        header = self.table_view.horizontalHeader()
        self.table_view.resizeColumnsToContents()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(3, QHeaderView.Stretch)

        for col in [0,1,2,4]:
            header.setSectionResizeMode(col, QHeaderView.ResizeToContents)
        
        # Set timer interval
        self.timer.setInterval(1000)

        # Disable enabled state
        self.downtime_button.setEnabled(False)
        self.pause_button.setEnabled(False)
        self.stop_button.setEnabled(False)

        # Set placeholder text
        self.rejects.setPlaceholderText("Enter Rejects")
        self.total_parts.setPlaceholderText("Enter Total Parts")
        self.ideal_rate.setPlaceholderText("Enter Ideal Rate (seconds per part)")
        self.rejects_tag.setPlaceholderText("Rejects Tag (optional)")
        self.total_parts_tag.setPlaceholderText("Total Parts Tag (optional)")
        self.ip_input.setPlaceholderText("Enter PLC IP")
        self.ideal_rate.setPlaceholderText("Ideal Part Rate")
        self.tree_filter.setPlaceholderText("Search Tags...")

        # Set fixed width
        self.total_parts.setFixedWidth(100)
        self.rejects.setFixedWidth(100)
        self.ideal_rate.setFixedWidth(100)
        self.ip_input.setFixedWidth(400)
        self.table_view.setFixedHeight(200)

        # Set alignment
        self.time.setAlignment(QtGui.Qt.AlignCenter)
        self.rejects.setAlignment(QtGui.Qt.AlignCenter)
        self.ideal_rate_label.setAlignment(QtGui.Qt.AlignCenter)
        self.ip_input.setAlignment(QtGui.Qt.AlignCenter)
        self.total_parts.setAlignment(QtGui.Qt.AlignCenter)
        self.rejects_tag.setAlignment(QtGui.Qt.AlignCenter)
        self.total_parts_tag.setAlignment(QtGui.Qt.AlignCenter)
        self.footer_label.setAlignment(Qt.AlignRight)

        # Validators
        self.rejects.setValidator(QtGui.QIntValidator(0, 1000000, self))
        self.total_parts.setValidator(QtGui.QIntValidator(0, 1000000, self))
        self.ideal_rate.setValidator(QtGui.QDoubleValidator(0.01, 1000000, 2, self))

        # Set fonts
        time_font = self.time.font()
        time_font.setBold(True)
        time_font.setPointSize(16)
        self.time.setFont(time_font)

        oee_font = self.oee.font()
        oee_font.setBold(True)
        oee_font.setPointSize(18)
        self.oee.setFont(oee_font)

        oee_components_font = self.oee_components.font()
        oee_components_font.setBold(True)
        oee_components_font.setPointSize(16)
        self.oee_components.setFont(oee_components_font)

        # Create empty and status variables
        self.logging_downtime = False
        self.downtime_start_time = None
        self.downtime_start_time_str = None
        self.downtime_end_time = None
        self.downtime_end_time_str = None

        # Create layouts and add wigets
        self.main_layout = QVBoxLayout()
        self.hor_layout = QHBoxLayout()
        self.part_count_layout = QHBoxLayout()
        self.plc_layout = QHBoxLayout()

        self.plc_layout.addWidget(self.ip_input)
        self.plc_layout.addWidget(self.get_tags_button)
        self.main_layout.addLayout(self.plc_layout)
        self.main_layout.addWidget(self.time)

        self.part_count_layout.addWidget(QLabel('Rejects:'))
        self.part_count_layout.addWidget(self.rejects)
        self.part_count_layout.addWidget(QLabel('Total Parts:'))
        self.part_count_layout.addWidget(self.total_parts)
        self.part_count_layout.addWidget(self.ideal_rate_label)
        self.part_count_layout.addWidget(self.ideal_rate)

        self.part_count_layout.addStretch()

        self.user_entry_layout = QHBoxLayout()
        self.user_entry_layout.addWidget(self.rejects_tag)
        self.user_entry_layout.addWidget(self.total_parts_tag)

        self.main_layout.addLayout(self.part_count_layout)
        self.main_layout.addLayout(self.user_entry_layout)
        self.main_layout.addWidget(self.downtime_button)
        self.main_layout.addWidget(self.export_downtime_button)
        self.main_layout.addWidget(self.table_view)
        self.fault_tag_layout = QHBoxLayout()
        self.fault_tag_layout.addWidget(self.tree_filter)
        self.fault_tag_layout.addWidget(self.add_fault_tag_button)
        self.main_layout.addLayout(self.fault_tag_layout)

        self.main_layout.addWidget(self.tree_filter)
        self.main_layout.addWidget(self.tree)
        self.main_layout.addWidget(self.footer_label)

        self.quality.setAlignment(QtGui.Qt.AlignCenter)
        self.availability.setAlignment(QtGui.Qt.AlignCenter)
        self.performance.setAlignment(QtGui.Qt.AlignCenter)
        self.oee_components.setAlignment(QtGui.Qt.AlignCenter)
        self.oee.setAlignment(QtGui.Qt.AlignCenter)
        self.calculated_rate.setAlignment(QtGui.Qt.AlignCenter)


        self.results_layout = QVBoxLayout()
        self.results_layout.addWidget(self.oee)
        self.results_layout.addWidget(self.oee_components)
        self.results_layout.addWidget(self.calculated_rate)
        self.results_layout.setAlignment(QtGui.Qt.AlignCenter)
        self.main_layout.addLayout(self.results_layout)
        self.hor_layout.addWidget(self.start_button)
        self.hor_layout.addWidget(self.stop_button)
        self.hor_layout.addWidget(self.pause_button)
        self.main_layout.addLayout(self.hor_layout)

        self.read_history()
        
        # Connect events
        self.timer.timeout.connect(self.update_display)
        self.start_button.clicked.connect(
            lambda: self.start_clicked())
        self.stop_button.clicked.connect(
            lambda: self.stop_clicked())
        self.pause_button.clicked.connect(
            lambda: self.pause_clicked())
        self.downtime_button.clicked.connect(
            lambda: self.downtime_clicked())
        self.export_downtime_button.clicked.connect(
            lambda: self.export_downtime_events())
        self.get_tags_button.clicked.connect(
            lambda: self.get_tags_clicked())
        self.add_fault_tag_button.clicked.connect(
            lambda: self.add_fault_tag_clicked())
        
        # Set central widget
        widget = QWidget()
        widget.setLayout(self.main_layout)
        self.setCentralWidget(widget)
    
        self.fault_tags = []
        self.fault_tags_list = []

    def handle_tag_selection(self, index):
        indexes = self.tree.selectedIndexes()

        selected_tags = []

        for index in indexes:
            source_index = self.proxy_model.mapToSource(index)
            item = self.proxy_model.sourceModel().itemFromIndex(source_index)
            tag_path = item.data(Qt.UserRole)
            
            if tag_path:
                selected_tags.append(tag_path)

        print(selected_tags)
        self.fault_tags = selected_tags

        return selected_tags

    def downtime_clicked(self):
        now = round_to_nearest_second(datetime.now())

        if not self.logging_downtime:
            # Start logging downtime
            self.temp_start_time = now
            self.logging_downtime = True
            self.downtime_button.setText("End Downtime")
        else:
            # End logging downtime
            event = DowntimeEvent(start_time=self.temp_start_time, end_time=now, cause="", nmr=False)
            self.events.append(event)
            self.model.layoutChanged.emit()
            self.logging_downtime = False
            self.downtime_button.setText("Record Downtime")
            self.update_footer()

    def start_clicked(self):
        self.downtime_button.setEnabled(True)  # Disable downtime button for now
        self.start_button.setEnabled(False)
        self.pause_button.setEnabled(True)
        self.stop_button.setEnabled(True)
        self.timer.start()
        self.time.setText(f'Run Time: 00:00:00')
        self.events.clear()
        self.model = DowntimeModel(self.events, round_to_nearest_second(datetime.now()), self.update_footer)
        self.run_start_time = datetime.now()
        self.model.layoutChanged.emit()
        self.table_view.setModel(self.model)

        self.save_history()

    def stop_clicked(self):
        self.timer.stop()
        self.start_button.setEnabled(True)
        self.downtime_button.setEnabled(False)
        self.pause_button.setEnabled(False)  # Disable downtime button for now
        self.update_display()
        self.save_history()

    def pause_clicked(self):
        if not self.is_paused:
            self.pause_start = datetime.now()
            self.is_paused = True
        else:
            pause_time = datetime.now() - self.pause_start
            self.paused_duration += pause_time
            self.pause_start = None
            self.is_paused = False

    def update_display(self):
        if self.is_paused:
            elapsed_secs = (self.pause_start - self.run_start_time - self.paused_duration).total_seconds()
        else:
            elapsed_secs = (datetime.now() - self.run_start_time - self.paused_duration).total_seconds()
        # get time in hh:mm:ss format
        hours, remainder = divmod(elapsed_secs, 3600)
        minutes, seconds = divmod(remainder, 60)
        formatted_time = f'{int(hours):02}:{int(minutes):02}:{int(seconds):02}'
        
        self.time.setText(f'Run Time: {formatted_time}')

        # calculate quality, availability, performance, and oee
        try:
            total_parts = int(self.total_parts.text())
            rejects = int(self.rejects.text())
            if total_parts == 0:
                quality = 0
            else:
                quality = ((total_parts - rejects) / total_parts) * 100

            formatted_quality = f"Quality: {quality:.2f}%"
            self.quality.setText(formatted_quality)
        except ValueError:
            formatted_quality = 'Quality: N/A'
            return

        # get the run time and add up the durations of all the downtimes to calcualte availability
        total_downtime = 0
        mr_count = 0
        nmr_count = 0
        mr_total = timedelta()
        nmr_total = timedelta()

        for event in self.events:
            if event.nmr:
                nmr_count += 1
                nmr_total += event.duration
            else:
                mr_count += 1
                mr_total += event.duration

        total_downtime = mr_total

        run_time = elapsed_secs - total_downtime.total_seconds()
        if elapsed_secs == 0:
            availability = 0
        else:
            availability = (run_time / elapsed_secs) * 100
        formatted_availability = f"Availability: {availability:.2f}%"
        self.availability.setText(formatted_availability)

        # get the rate by taking the run time minus the downtime and total parts
        if total_parts == 0:
            performance = 0
            calculated_rate = "N/A"
        else:
            try:
                performance = ((float(self.ideal_rate.text()) * total_parts) / run_time) * 100
            except ValueError:
                QMessageBox.warning(self, "Input Error", "Please enter a valid Ideal Rate.")
                return
            calculated_rate = run_time / total_parts if total_parts > 0 else 0
            calculated_rate = f"{calculated_rate:.2f} seconds per part"
        formatted_performance = f"Performance: {performance:.2f}%"
        self.performance.setText(formatted_performance)
        self.calculated_rate.setText(f"Calculated Rate: {calculated_rate}")

        # calculate oee
        if availability == 0 or performance == 0 or quality == 0:
            oee = 0
        else:
            oee = (availability * performance * quality) / 10000
        formatted_oee = f"{oee:.2f}%"
        self.oee.setText(f"OEE: {formatted_oee}")

        self.oee_components.setText(f"{formatted_performance} | {formatted_availability} | {formatted_quality}")

    def update_footer(self):
        mr_count = 0
        nmr_count = 0
        mr_total = timedelta()
        nmr_total = timedelta()

        for event in self.events:
            if event.nmr:
                nmr_count += 1
                nmr_total += event.duration
            else:
                mr_count += 1
                mr_total += event.duration

        def format_td(td):
            total_seconds = int(td.total_seconds())

            h = total_seconds // 3600
            m = (total_seconds % 3600) // 60
            s = total_seconds % 60

            return f"{h}:{m:02}:{s:02}"
        
        self.footer_label.setText(f"Machine Related: {mr_count} events, {format_td(mr_total)} | Non-Machine Related: {nmr_count} events, {format_td(nmr_total)}")

    def export_downtime_events(self):
        events = []

        for event in self.events:
            events.append([event.start_time, event.end_time, event.duration, event.cause, event.nmr])

        try:
            with open('output.csv', 'w', newline='') as file:
                writer = csv.writer(file)
                writer.writerows(events)
            print("List successfully written to output.csv")
        except IOError:
            print("Error: Could not write to file.")

    def get_tags_clicked(self):
        ip = self.ip_input.text()
        if ip:
            tags = get_tags_from_plc()
            model = build_tag_tree_model(tags)
            self.proxy_model = TagFilterProxyModel()

            self.proxy_model.setSourceModel(model)

            self.proxy_model.setFilterCaseSensitivity(Qt.CaseInsensitive)
            self.proxy_model.setRecursiveFilteringEnabled(True)

            self.tree.setModel(self.proxy_model)
            self.tree_filter.textChanged.connect(self.proxy_model.setFilterFixedString)

            self.tree.clicked.connect(self.handle_tag_selection)

    def add_fault_tag_clicked(self):
        for tag in self.fault_tags:
            if tag not in self.fault_tags_list:
                self.fault_tags_list.append(tag)

        print(self.fault_tags_list)


    def read_history(self):
        self.ip_input.setText(self.settings.value('ip', ''))


    def save_history(self):
        self.settings.setValue('ip', self.ip_input.text())

app = QApplication(sys.argv)
app.processEvents()
app.setWindowIcon(QtGui.QIcon('icon.ico'))
qdarktheme.setup_theme()
window = MainWindow()
window.resize(1000, 500)
window.show()

app.exec()