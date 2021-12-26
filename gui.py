#!/usr/bin/env python3

# Copyright (C) 2021 Biren Patel, GNU General Public License v3.0.
# PySide2 GUI

import sys
from time import sleep

#Qt imports
from PySide2.QtCore import Slot, Qt, QRect
from PySide2.QtGui import QPixmap
from PySide2.QtWidgets import QApplication, QSplashScreen, QMainWindow
from PySide2.QtWidgets import QVBoxLayout, QHBoxLayout, QGridLayout
from PySide2.QtWidgets import QWidget, QStackedWidget
from PySide2.QtWidgets import QAction, QLabel, QPushButton, QFileDialog
from PySide2.QtWidgets import QProgressBar

#local imports
import reconcile 
from style import style_sheet

class mainWindow(QMainWindow):
    def __init__(self):
        """
        Inherits from QMainWindow and setup for window and menu bar
        """
        super().__init__()

        #initialize the application
        self.create_window()
        self.create_status_bar()
        self.create_menu()
        self.create_central_widgets()

    def create_window(self):
        """
        Set the main window attributes on initialization
        """
        self.setWindowTitle("Bluebonnet Computing - Accounting & Legal")
        self.resize(1200,700)

    def create_status_bar(self):
        """
        Set the status bar at the bottom of the main window
        """
        self.status_bar = self.statusBar()

        #add permanent progress bar. has style sheet.
        self.progress_bar_label = QLabel("Status: ")
        self.status_bar.addPermanentWidget(self.progress_bar_label)
        self.progress_bar = QProgressBar()
        self.status_bar.addPermanentWidget(self.progress_bar)

        #add temporary ready message on launch
        self.show_temp_status("Welcome")

    def show_temp_status(self, message, time=10000):
        """
        Helper function for status bar temporary messages
        """
        self.status_bar.showMessage(message, timeout=time)

    def create_menu(self):
        """
        Create the main window menu bar on initialization
        """
        #create menu bar
        self.menu = self.menuBar()

        ### FILE ###

        #file menu
        self.file_menu = self.menu.addMenu("File")

        #home option in file
        home_action = QAction("Home", self)
        home_action.setShortcut("Ctrl+H")
        home_action.triggered.connect(self.menu_file_home)
        self.file_menu.addAction(home_action)

        #exit option in file
        exit_action = QAction("Exit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.menu_file_exit)
        self.file_menu.addAction(exit_action)

        ### TASK ###

        #task menu
        self.task_menu = self.menu.addMenu("Tasks")

        #reconcile statements option in tasks
        reconcile_action = QAction("Reconcile Statements", self)
        reconcile_action.setShortcut("Alt+R")
        reconcile_action.triggered.connect(self.menu_task_reconcile)
        self.task_menu.addAction(reconcile_action)

        ### HELP ###

        #help menu
        self.help_menu = self.menu.addMenu("Help")

        #version option in help
        version_action = QAction("Version Info", self)
        version_action.setShortcut("Ctrl+V")
        version_action.triggered.connect(self.menu_help_version)
        self.help_menu.addAction(version_action)

    def create_central_widgets(self):
        """
        Application uses central widgets to load menu selections
        """
        #create all widget instances to be used as central widgets
        self.launch_widget = launch()
        self.reconcile_widget = reconcile(self.progress_bar)
        self.version_widget = version()

        #add above widget instances to a stack instance
        self.stacked_widgets = QStackedWidget()
        self.stacked_widgets.addWidget(self.launch_widget)
        self.stacked_widgets.addWidget(self.reconcile_widget)
        self.stacked_widgets.addWidget(self.version_widget)

        #on application start set the launch widget to display
        self.stacked_widgets.setCurrentWidget(self.launch_widget)
        self.setCentralWidget(self.stacked_widgets)

    @Slot()
    def menu_file_home(self):
        """
        Helper for home_action object in self.create_menu
        Reset application to launch configuration
        """
        self.stacked_widgets.setCurrentWidget(self.launch_widget)
        self.setCentralWidget(self.stacked_widgets)

    @Slot()
    def menu_file_exit(self):
        """
        Helper for exit_action object in self.create_menu
        """
        QApplication.quit()

    @Slot()
    def menu_task_reconcile(self):
        """
        Helper for reconcile_action in self.create_menu
        """
        self.stacked_widgets.setCurrentWidget(self.reconcile_widget)
        self.setCentralWidget(self.stacked_widgets)

        #alternate popup code, delete reconcile from stack if using this
        #self.test_local = reconcile(self.progress_bar)
        #self.test_local.setGeometry(QRect(100, 100, 400, 200))
        #self.test_local.show()

    @Slot()
    def menu_help_version(self):
        """
        Helper for version_action in self.create_menu
        """
        self.stacked_widgets.setCurrentWidget(self.version_widget)
        self.setCentralWidget(self.stacked_widgets)

#central widget on application launch
class launch(QWidget):
    def __init__(self):
        """
        no main interface since each task is varied, so central widget on the
        application launch prompts user to select a task.
        """
        super().__init__()

        #create label and align text to dead center
        self.text = QLabel("Select a task from the menu to get started.")
        self.text.setAlignment(Qt.AlignCenter)

        #layout
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.text)
        self.setLayout(self.layout)

class version(QWidget):
    def __init__(self):
        """
        Basic software and authoring information on version tab in help menu.
        """
        super().__init__()

        #create labels and align text to dead center
        self.text_1 = QLabel("Biren Patel")
        self.text_1.setAlignment(Qt.AlignCenter)

        self.text_2 = QLabel("biren.dilip.patel@gmail.com")
        self.text_2.setAlignment(Qt.AlignCenter)

        self.text_3 = QLabel("Beta Version 0.2.0")
        self.text_3.setAlignment(Qt.AlignCenter)

        self.text_4 = QLabel("Developed in Python 3.8.3 with PySide2 5.14.2")
        self.text_4.setAlignment(Qt.AlignCenter)

        #layout
        self.layout = QVBoxLayout()
        self.layout.addStretch()
        self.layout.addWidget(self.text_1)
        self.layout.addWidget(self.text_2)
        self.layout.addWidget(self.text_3)
        self.layout.addWidget(self.text_4)
        self.layout.addStretch()
        self.setLayout(self.layout)

class reconcile(QWidget):
    def __init__(self, progress_bar):
        """
        reconciliation item in task menu. initialize sets up widgets and layout
        but the central widget controls are in the mainWindow class.
        """
        super().__init__()

        #prompts for files
        self.bank_label = QLabel("Select Chase Bank Statement: ")
        self.reco_label = QLabel("Select Reconciliation Statement: ")
        self.save_label = QLabel("Select Location to Save the Results: ")

        #bank file browser
        self.bank_file = ""
        self.bank_browse = QPushButton("Browse")
        self.bank_browse.clicked.connect(self.get_bank_file)

        #reconciliation file browser
        self.reco_file = ""
        self.reco_browse = QPushButton("Browse")
        self.reco_browse.clicked.connect(self.get_reco_file)

        #save file browser
        self.save_file = ""
        self.save_browse = QPushButton("Browse")
        self.save_browse.clicked.connect(self.get_save_file)

        #generate reconciliation report
        self.generate = QPushButton("Click to Generate Report")
        self.generate.clicked.connect(self.execute_reconciliation)
        self.generate_notice = QLabel("                         `")
        self.generate_notice.setAlignment(Qt.AlignCenter)

        #load in progress bar from status bar in main window
        self.progress_bar = progress_bar

        #layout management
        self.overall_layout = QVBoxLayout()
        self.bank_layout = QHBoxLayout()
        self.reco_layout = QHBoxLayout()
        self.save_layout = QHBoxLayout()
        self.gen_layout = QHBoxLayout()
        self.gen_layout_bottom = QHBoxLayout()

        self.bank_layout.addStretch()
        self.bank_layout.addWidget(self.bank_label)
        self.bank_layout.addWidget(self.bank_browse)
        self.bank_layout.addStretch()
        self.reco_layout.addStretch()
        self.reco_layout.addWidget(self.reco_label)
        self.reco_layout.addWidget(self.reco_browse)
        self.reco_layout.addStretch()
        self.save_layout.addStretch()
        self.save_layout.addWidget(self.save_label)
        self.save_layout.addWidget(self.save_browse)
        self.save_layout.addStretch()
        self.gen_layout.addStretch()
        self.gen_layout.addWidget(self.generate)
        self.gen_layout.addStretch()
        self.gen_layout_bottom.addStretch()
        self.gen_layout_bottom.addWidget(self.generate_notice)
        self.gen_layout_bottom.addStretch()

        self.overall_layout.addStretch()
        self.overall_layout.addLayout(self.bank_layout)
        self.overall_layout.addLayout(self.reco_layout)
        self.overall_layout.addLayout(self.save_layout)
        self.overall_layout.addLayout(self.gen_layout)
        self.overall_layout.addLayout(self.gen_layout_bottom)
        self.overall_layout.addStretch()

        self.setLayout(self.overall_layout)

    @Slot()
    def get_bank_file(self):
        """
        popup file browser to select appropriate csv file of bank statements
        """
        caption = "Select Bank File"
        filter = "*.csv"
        bank_file = QFileDialog.getOpenFileName(self, \
                    caption=caption, filter=filter)

        self.bank_file = bank_file[0]

    @Slot()
    def get_reco_file(self):
        """
        popup file browser to select appropriate xlsx file of reconciliations
        """
        caption = "Select Reconciliation File"
        filter = "*.xlsx"
        reco_file = QFileDialog.getOpenFileName(self, \
                    caption=caption, filter=filter)

        self.reco_file = reco_file[0]

    @Slot()
    def get_save_file(self):
        """
        popup file browser to select appropriate save location for xlsx result
        """
        caption = "Select Save Location"
        filter = "*.xlsx"
        save_file = QFileDialog.getSaveFileName(self, \
                    caption=caption, filter=filter)

        self.save_file = save_file[0]

    @Slot()
    def execute_reconciliation(self):
        """
        executes the main reconciliation task once both files are provided.

        enumerations from reconcile module:
        NULL_STATUS = -1
        SUCCESS = 0
        FAIL_BANK_ACCT_LEN = 1
        FAIL_RECO_NO_ID = 2
        FAIL_RECO_NO_PLUS_3 = 3
        FAIL_RECO_ACCT_LEN = 4
        FAIL_RECO_LARGE = 5
        """
        if self.bank_file == "":
            self.generate_notice.setText("An Error Occurred: No Bank File")
        elif self.reco_file == "":
            self.generate_notice.setText("An Error Occurred: No Reconciliation File")
        elif self.save_file == "":
            self.generate_notice.setText("An Error Occurred: No Save Location for Results")
        else:
            self.generate_notice.setText("Working...")

            data, status_code = \
            reconcile.reconcile(self.bank_file, self.reco_file, self.progress_bar)

            if status_code == reconcile.status.SUCCESS:
                reconcile.create_xlsx(data, self.save_file)
                self.generate_notice.setText("Finished!")
                sleep(1.5)
            else:
                issue = status_code.name
                self.generate_notice.setText("An Error Occured: {}".format(issue))

            self.progress_bar.reset()

if __name__ == "__main__":
    application = QApplication([])

    #create splash screen and run an animation before launch
    splash_screen = QSplashScreen(QPixmap("splash_image.png"))
    splash_screen.showMessage("(0%) Loading Modules...", alignment=66)
    splash_screen.show()

    for i in range(1,101):
        msg = "({}%) Loading Modules...".format(i)
        splash_screen.showMessage(msg, alignment=66)
        application.processEvents()
        sleep(.01)

    splash_screen.close()

    #create main window and launch application
    window = mainWindow()
    window.setStyleSheet(style_sheet)
    window.show()
    application.exec_()
