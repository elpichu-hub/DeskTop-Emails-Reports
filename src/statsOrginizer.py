import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QVBoxLayout, QTextEdit,
                             QDialog, QLabel, QLineEdit, QDialogButtonBox, QFileDialog)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QDateTime
import re

### Command to create the app: pyinstaller --name=Reports --onefile --windowed --icon=favicon.ico statsOrginizer.py ###

class App(QWidget):

    # initialize the window and set its properties
    def __init__(self):
        super().__init__()
        self.title = 'Reports By Lazaro Gonzalez'
        self.left = 100
        self.top = 100
        self.width = 700
        self.height = 700
        self.initUI()
        self.key = 'target'
        self.key2_crypto = b'\\\x88\x0b\x08\xdd\r\xd6\x1e\xc2\x13\xec\x18\xadm)\x96'

    # initialize the UI
    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        # Set the window icon
        self.setWindowIcon(
            QIcon('C:\MyStuff\Projects\Developing\Tests\StatsOrginizer\src\icon.png'))

        # Create a layout
        layout = QVBoxLayout()


        # Create a button and style it
        button1 = QPushButton(
            'Generate Daily, Weekly, Monthly, Yearly Stats', self)
        button1.setToolTip('Generate statistics')
        button1.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')
        # Set the cursor to a pointer when hovering the button
        button1.setCursor(Qt.PointingHandCursor)
        # Connect the button to the function that generates the statistics
        button1.clicked.connect(self.on_click_stats)
       

        # Create a button and style it
        button2 = QPushButton('None Work Codes', self)
        button2.setToolTip('Break Down Non Work Codes By Agent')
        button2.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')
        # Set the cursor to a pointer when hovering the button
        button2.setCursor(Qt.PointingHandCursor)
        # Connect the button to the function that generates the statistics
        button2.clicked.connect(self.non_work_report_click)


        # Create a button and style it
        button3 = QPushButton('Proponisi Stats', self)
        button3.setToolTip('Creates Proponisi csv file')
        button3.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')
        # Set the cursor to a pointer when hovering the button
        button3.setCursor(Qt.PointingHandCursor)
        # Connect the button to the function that generates the statistics
        button3.clicked.connect(self.proponisi_report_click)


        # Create a button and style it
        button4 = QPushButton('Proponisi QA Stats', self)
        button4.setToolTip('Adds QA Stats To Proponisi File')
        button4.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')
        # Set the cursor to a pointer when hovering the button
        button4.setCursor(Qt.PointingHandCursor)
        # Connect the button to the function that generates the statistics
        button4.clicked.connect(self.qa_stats_proponisi_click)


        button5 = QPushButton('Send Quality Assurance', self)
        button5.setToolTip('Import quality assurance')
        button5.setStyleSheet('''
            QPushButton {
                background-color: grey;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')
        # Set the cursor to a pointer when hovering the button
        button5.setCursor(Qt.PointingHandCursor)
        # Disable the button
        button5.setEnabled(False)
        # Connect the button to the function that sends quality assurance
        button5.clicked.connect(self.send_quality_assurance_click)


        button6 = QPushButton('Proponisi Perfect Attendance', self)
        button6.setToolTip('Mark Pefect Attendance on Proponisi File')
        button6.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')
        # Set the cursor to a pointer when hovering the button
        button6.setCursor(Qt.PointingHandCursor)
        # Connect the button to the function that sends quality assurance
        button6.clicked.connect(self.mark_perfect_attedance_proponisi)


        button7 = QPushButton('Proponisi Profiles', self)
        button7.setToolTip('Creates Proponisi Profiles')
        button7.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')
        # Set the cursor to a pointer when hovering the button
        button7.setCursor(Qt.PointingHandCursor)
        # Connect the button to the function that sends quality assurance
        button7.clicked.connect(self.create_proponisi_profiles)


        button8 = QPushButton('Time Per Queues', self)
        button8.setToolTip('Calculates Time Spent on Specific Queues')
        button8.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')
        # Set the cursor to a pointer when hovering the button
        button8.setCursor(Qt.PointingHandCursor)
        # Connect the button to the function that sends quality assurance
        button8.clicked.connect(self.calculate_time_on_queues)


        button9 = QPushButton('Update Proponisi Profiles', self)
        button9.setToolTip('Update Proponisi Profiles')
        button9.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')
        # Set the cursor to a pointer when hovering the button
        button9.setCursor(Qt.PointingHandCursor)
        # Connect the button to the function that sends quality assurance
        button9.clicked.connect(self.update_proponisi_profiles)


        button10 = QPushButton('Attendance Competition', self)
        button10.setToolTip('Attendance Competition')
        button10.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')
        # Set the cursor to a pointer when hovering the button
        button10.setCursor(Qt.PointingHandCursor)
        # Connect the button to the function that sends quality assurance
        button10.clicked.connect(self.on_click_attendance_competition)


        button11 = QPushButton('Bulk Email Stringify', self)
        button11.setToolTip('Creates string of emails to send bulk emails for Survey & fixes any errors in the email.')
        button11.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')
        # Set the cursor to a pointer when hovering the button
        button11.setCursor(Qt.PointingHandCursor)
        # Connect the button to the function that sends quality assurance
        button11.clicked.connect(self.on_click_stringify_emails)


        button12 = QPushButton('Survey Assigner', self)
        button12.setToolTip('Assigns an equal amount of Surveys to agents by Color')
        button12.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')
        # Set the cursor to a pointer when hovering the button
        button12.setCursor(Qt.PointingHandCursor)
        # Connect the button to the function that sends quality assurance
        button12.clicked.connect(self.assign_surveys_click)


        button13 = QPushButton('SBS Generator', self)
        button13.setToolTip('Generates SBS')
        button13.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')
        # Set the cursor to a pointer when hovering the button
        button13.setCursor(Qt.PointingHandCursor)
        # Connect the button to the function that sends quality assurance
        button13.clicked.connect(self.sbs_generator_click)


        button14 = QPushButton('Reset IE to Defaults', self)
        button14.setToolTip('Reset IE to Defaults')
        button14.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #F4B942;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')
        # Set the cursor to a pointer when hovering the button
        button14.setCursor(Qt.PointingHandCursor)
        # Connect the button to the function that sends quality assurance
        button14.clicked.connect(self.reset_IE_to_defaults_click)


        button15 = QPushButton('GP Update', self)
        button15.setToolTip('GP Update')
        button15.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #F4B942;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')
        # Set the cursor to a pointer when hovering the button
        button15.setCursor(Qt.PointingHandCursor)
        # Connect the button to the function that sends quality assurance
        button15.clicked.connect(self.gp_update_click)


        button16 = QPushButton('Look Ups', self)
        button16.setToolTip('Look Ups')
        button16.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #F4B942;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')

        # Set the cursor to a pointer when hovering the button
        button16.setCursor(Qt.PointingHandCursor)
        # Connect the button to the function that sends quality assurance
        button16.clicked.connect(self.encrypt_decrypt_click)


        # Create a QTextEdit for output
        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)

        # Add the current date and time to the output text
        current_datetime = QDateTime.currentDateTime()
        self.output_text.append(
            f"Date & Time: {current_datetime.toString('yyyy-MM-dd hh:mm:ss')}")
        
        # Add buttons and QTextEdit to the layout
        layout.addWidget(button1)
        layout.addWidget(button2)
        layout.addWidget(button3)
        layout.addWidget(button4)
        layout.addWidget(button5)
        layout.addWidget(button6)
        layout.addWidget(button7)
        layout.addWidget(button8)
        layout.addWidget(button9)
        layout.addWidget(button10)
        layout.addWidget(button11)
        layout.addWidget(button12)
        layout.addWidget(button13)
        layout.addWidget(button14)
        layout.addWidget(button15)
        layout.addWidget(button16)
        layout.addWidget(self.output_text)

        self.setLayout(layout)
        self.show()

    # function that runs to get data from the user
    def get_date(self, QlableText):
        dialog = QDialog()
        dialog.setWindowTitle("Enter Date")

        # label = QLabel("Enter the date range, mm.dd - mm.dd:", dialog)
        label = QLabel(QlableText, dialog)
        line_edit = QLineEdit(dialog)
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dialog)
        layout = QVBoxLayout(dialog)
        layout.addWidget(label)
        layout.addWidget(line_edit)
        layout.addWidget(button_box)

        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)

        result = dialog.exec_()

        if result == QDialog.Accepted:
            return line_edit.text()
        else:
            return None

    # function that runs to get data from the user to send QA to
    def get_email(self):
        dialog = QDialog()
        dialog.setWindowTitle("Enter Email to Receive QAs")

        label = QLabel(
            "Provide an Email Address, All the QAs Will be Sent There:", dialog)
        line_edit = QLineEdit(dialog)
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dialog)
        layout = QVBoxLayout(dialog)
        layout.addWidget(label)
        layout.addWidget(line_edit)
        layout.addWidget(button_box)

        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)

        result = dialog.exec_()

        if result == QDialog.Accepted:
            return line_edit.text()
        else:
            return None

    # this will be used to get WIN ID
    def get_proponisi_winID(self):
        dialog = QDialog()
        dialog.setWindowTitle("Enter Supervisor WIN ID")

        label = QLabel("Enter Supervisor WIN ID", dialog)
        line_edit = QLineEdit(dialog)
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dialog)
        layout = QVBoxLayout(dialog)
        layout.addWidget(label)
        layout.addWidget(line_edit)
        layout.addWidget(button_box)

        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)

        result = dialog.exec_()

        if result == QDialog.Accepted:
            return line_edit.text()
        else:
            self.output_text.append('No WIN ID Entered')
            return None

    # this will be used to get location
    def get_proponisi_location(self):
        dialog = QDialog()
        dialog.setWindowTitle("Enter Location (Ocoee or Boca)")

        label = QLabel("Enter Location (Ocoee or Boca)", dialog)
        line_edit = QLineEdit(dialog)
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dialog)
        layout = QVBoxLayout(dialog)
        layout.addWidget(label)
        layout.addWidget(line_edit)
        layout.addWidget(button_box)

        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)

        result = dialog.exec_()

        if result == QDialog.Accepted:
            return line_edit.text()
        else:
            self.output_text.append('No Location Entered')
            return None

    # simple function to get password
    def get_password(self):
        dialog = QDialog()
        dialog.setWindowTitle("Provide Key")

        label = QLabel("Provide Key to Continue", dialog)
        line_edit = QLineEdit(dialog)
        line_edit.setEchoMode(QLineEdit.Password)  # Hide the entered text
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dialog)
        layout = QVBoxLayout(dialog)
        layout.addWidget(label)
        layout.addWidget(line_edit)
        layout.addWidget(button_box)

        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)

        result = dialog.exec_()

        if result == QDialog.Accepted:
            return line_edit.text()
        else:
            self.output_text.append('No Key Entered, Cancelled')
            return None


    # this will be used to get a team
    def get_team(self):
        dialog = QDialog()
        dialog.setWindowTitle("Enter Team (Optional)")

        label = QLabel("Enter Team (Optional)", dialog)
        line_edit = QLineEdit(dialog)
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dialog)
        layout = QVBoxLayout(dialog)
        layout.addWidget(label)
        layout.addWidget(line_edit)
        layout.addWidget(button_box)

        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)

        result = dialog.exec_()

        if result == QDialog.Accepted:
            return line_edit.text()
        else:
            self.output_text.append('No Team Entered')
            return None
        
    # click to encrypt file for ROV look up
    def encrypt_decrypt_click(self):
        import encrypting

        # ask for password
        password = self.get_password()

        if password != self.key:
            self.append_output("Invalid Key\n")
            return 
        else:
            self.append_output("Access Granted\n")

        # Ask user to select survey file
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)

        file_path, _ = file_dialog.getOpenFileName(
            self, "Select CSV File", "", "XLSX Files (*.xlsx);;XLS Files (*.xls);;CSV Files (*.csv);;All Files (*)"
        )

        if not file_path:
            self.append_output("File not selected. Process Cancelled.\n")
            return
        
        try:
            encrypting.main_func_encrypt(file_path=file_path, key=self.key2_crypto)
            self.append_output("Success.\n")
        except Exception as e:
            self.append_output("Error:")
            self.append_output(f'{str(e)}\n')

    # to reset internet explorer to defaults
    def get_times_to_reset(self):
        dialog = QDialog()
        dialog.setWindowTitle("Enter Number of Times to Reset (1 = 1, 0 = 10 hours)")

        label = QLabel("Enter Number of Times to Reset (1 = 1, 0 = 10 hours)", dialog)
        line_edit = QLineEdit(dialog)
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dialog)

        layout = QVBoxLayout(dialog)
        layout.addWidget(label)
        layout.addWidget(line_edit)
        layout.addWidget(button_box)

        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)

        result = dialog.exec_()

        if result == QDialog.Accepted:
            return line_edit.text()
        else:
            self.output_text.append('No Number of times to reset Entered')
            return None
        
    # reapplies all policy settings
    def gp_update_click(self):
        import subprocess

        # ask for password
        password = self.get_password()

        if password != self.key:
            self.append_output("Invalid Key\n")
            return 
        else:
            self.append_output("Access Granted\n")

        try:
            # Run the gpupdate command
            # /force reapplies all policy settings
            subprocess.run('gpupdate /force', shell=True, check=True)
            self.append_output("Group Policy Update executed successfully.\n")
        except subprocess.CalledProcessError as e:
            print(f"An error occurred: {e}")
            self.append_output(f"An error occurred: {e}\n")
        
    # reset internet explorer to defaults
    def reset_IE_to_defaults_click(self):
        import reset_proxy
        import time

        # ask for password
        password = self.get_password()

        if password != self.key:
            self.append_output("Invalid Key\n")
            return 
        else:
            self.append_output("Access Granted\n")

        # get number of times to reset 1 = 1, 0 = 10 hours
        times_to_reset = self.get_times_to_reset()

        if not times_to_reset:
            self.append_output("No Times to Reset Entered, Process Cancelled.\n")
            return
        
        if times_to_reset == '1':
            self.append_output("Times to Reset: 1\n")
            reset_proxy.reset_IE_to_defaults()
            self.append_output("Success, IE Reset.\n")

        elif times_to_reset == '0':
            try:
                # Calculate the time when the script should stop running (10 hours from now)
                end_time = time.time() + 10*60*60  # 10 hours in seconds
                # Run the function every 5 minutes until the end time is reached
                while time.time() < end_time:
                    reset_proxy.reset_IE_to_defaults()
                    print(f"Waiting 5 minutes before the next reset at {time.ctime(time.time() + 5*60)}.")
                    time.sleep(5*60)  # Wait for 5 minutes (300 seconds)
                    self.append_output("Success, IE Reset.\n")
            except Exception as e:
                self.append_output("Error with IE Reset:")
                self.append_output(f'{str(e)}\n')
        
    # generates SBS
    def sbs_generator_click(self):
        import sbs_generator
        
        # ask for password
        password = self.get_password()

        if password != self.key:
            self.append_output("Invalid Key\n")
            return 
        else:
            self.append_output("Access Granted\n")

        self.append_output(
            "Steps to Follow To Generate SBS:\n1. Select SBS File with Mentors and Mentees.\n2. SBS File Will Be Created.\n"
        )

        # Ask user to select survey file
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)

        file_path, _ = file_dialog.getOpenFileName(
            self, "Select SBS File", "", "XLSX Files (*.xlsx);;XLS Files (*.xls);;CSV Files (*.csv);;All Files (*)"
        )

        if not file_path:
            self.append_output("SBS File Not Selected. Process Cancelled.\n")
            return
        
        team = self.get_team()

        if not team:
            self.append_output("SBS Team Not Selected. Process Cancelled.\n")
            return
        
        try:
            sbs_generator.create_sbs(file_path=file_path, email_address='none@none.com', team=team, throughEmail=False)
            self.append_output("Success, SBS Generated.\n")
        except Exception as e:
            self.append_output("Error with SBS:")
            self.append_output(f'{str(e)}\n')

        return
        
    # Assigns surveys to agents
    def assign_surveys_click(self):
        import survey_assignments

        # ask for password
        password = self.get_password()

        if password != self.key:
            self.append_output("Invalid Key\n")
            return 
        else:
            self.append_output("Access Granted\n")

        self.append_output(
            "Steps to Follow To Assign Surveys to Agents:\n1. Select Survey File.\n2. Provide List of Agents Who Will Work on Surveys.\n3. File Will Be Created.\n"
        )

        # Ask user to select survey file
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)

        file_path, _ = file_dialog.getOpenFileName(
            self, "Select Survey File", "", "XLSX Files (*.xlsx);;XLS Files (*.xls);;CSV Files (*.csv);;All Files (*)"
        )

        if not file_path:
            self.append_output("Survey File Not selected. Process Cancelled.\n")
            return

        agents_string = self.get_agents_for_surveys()

        if not agents_string:
            self.append_output("Agents Not selected. Process Cancelled.\n")
            return
        
        # Split the agent names string into a list, and capitalize the first letter of each name
        agents_for_survey = [name.strip().capitalize() for name in agents_string.split(',') if name.strip()]
        self.append_output(f'Agent who will work on Surveys: {agents_for_survey}')

        
        try:
            survey_assignments.survey_email_generator(file_path1=file_path, email_address='none@none.com', agents_for_survey=agents_for_survey, throughApp=True)
            self.append_output("Success, Surveys Assigned.\n")
        except Exception as e:
            self.append_output("Error with Assigning Surveys:")
            self.append_output(f'{str(e)}\n')


    # This will collect the names of the agents who will work on the surveys
    def get_agents_for_surveys(self):
        dialog = QDialog()
        dialog.setWindowTitle("Enter Agents for Survey")
        dialog.resize(400, 300)  # Adjust the size of the dialog

        label = QLabel("Enter Agents for Survey", dialog)
        text_edit = QTextEdit(dialog)  # Use QTextEdit for multiline input
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dialog)

        layout = QVBoxLayout(dialog)
        layout.addWidget(label)
        layout.addWidget(text_edit)
        layout.addWidget(button_box)

        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)

        result = dialog.exec_()
        if result == QDialog.Accepted:
            return text_edit.toPlainText()  # Return the input text if Ok is pressed
        return None
        
    # this will be used to get emails to stringify
    def get_emails_to_stringify(self):
        dialog = QDialog()
        dialog.setWindowTitle("Enter Emails to Stringify")
        dialog.resize(400, 300)  # Adjust the size of the dialog

        label = QLabel("Enter Emails to Stringify", dialog)
        text_edit = QTextEdit(dialog)  # Use QTextEdit for multiline input
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dialog)

        layout = QVBoxLayout(dialog)
        layout.addWidget(label)
        layout.addWidget(text_edit)
        layout.addWidget(button_box)

        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)

        result = dialog.exec_()
        if result == QDialog.Accepted:
            return text_edit.toPlainText()  # Return the input text if Ok is pressed
        return None

    # this will stringify emails 
    def on_click_stringify_emails(self):
        import bulkEmailStringCreator

        # ask for password
        password = self.get_password()

        if password != self.key:
            self.append_output("Invalid Key\n")
            return 
        else:
            self.append_output("Access Granted\n")
        
        self.append_output(
            "Steps to Follow To Create String Of Emails:\n1. Paste in text field the emails.\n2. A list will be generated ready for copy and paste.\n"
        )

        emails = self.get_emails_to_stringify()

        try:
            email_string = bulkEmailStringCreator.create_bulk_email_string(emails=emails)
            self.append_output("*****************\nSunPass Survey Acknowledgement\n*****************")
            self.append_output("Hi, below you can find the emails for the surveys acknowledged:\n")
            email_count = len(email_string.split(';'))
            self.append_output(f'Email Count: {email_count}')
            self.append_output(f'{email_string}\n')
        
        except Exception as e:
            self.append_output("Email Stringify error:")
            self.append_output(f'{str(e)}\n')
        
    # Attendance Competition
    def on_click_attendance_competition(self):
        import AttendanceCompetition

        # ask for password
        password = self.get_password()

        if password != self.key:
            self.append_output("Invalid Key\n")
            return 
        else:
            self.append_output("Access Granted\n")

        # Ask user to select the call outs file
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        self.append_output(
            "Steps to Follow To Update Attendance for Competition:\n1. Select C/O list file.\n2. Select Weekly Attendance File.\n3. Select Attendance Competition.xlsx to Update.\n4. Attendance Competition.xlsx will be Updated and Opened."
        )


        file_path, _ = file_dialog.getOpenFileName(
            self, "Select Daily C/O File", "", "CSV Files (*.csv)")

        if not file_path:
            self.append_output("C/O File Not selected. Process Cancelled.\n")
            return
        
        # Ask user to select the weekly attendance file
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)


        weekly_attendance_file, _ = file_dialog.getOpenFileName(
            self, "Select Weekly Attendance File in CSV Format", "", "CSV Files (*.csv)")

        if not weekly_attendance_file:
            self.append_output("Weekly Attendance File Not selected. Processing Only Daily Attendance.\n")
        
        # Ask user to select the file to update
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)

        competitionFile, _ = file_dialog.getOpenFileName(
            self, "Select Attendance Competition File to update", "", "XLSX Files (*.xlsx);;XLS Files (*.xls);;All Files (*)"
        )


        if not competitionFile:
            self.append_output("Attendance Competition File Not selected. Process Cancelled.\n")
            return

        try:
            AttendanceCompetition.attendance_competition(file_path=file_path, competitionFile=competitionFile, weekly_attendance=weekly_attendance_file,)
            self.append_output("Success, Attendance Competition Updated.\n")
        except Exception as e:
            self.append_output("Error Attendance Competition:")
            self.append_output(f'{str(e)}\n')

    # this will generate daily, weekly, monthly or yearly stats given xlsx or xls file
    def on_click_stats(self):
        import statsXLS

        # ask for password
        password = self.get_password()
        
        if password != self.key:
            self.append_output("Invalid Key\n")
            return 
        else:
            self.append_output("Access Granted\n")
        
        # Ask user to select the stats file
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        self.append_output(
            "Steps to Follow To Run Stats:\n1. Enter Team (Optional).\n2. Pick the User Productivity File in .xlxs or .xls.\n3. A workbook will be created and opened with the stats formatted already."
        )

        team = self.get_team()

        file_path, _ = file_dialog.getOpenFileName(
            self, "Select Stats in CSV Format", "", "Excel Files (*.xls *.xlsx)")

        if not file_path:
            self.append_output("File Not selected. Process Cancelled.\n")
            return
        self.append_output("Selected file: " + file_path)

        # Set the cursor to WaitCursor
        QApplication.setOverrideCursor(Qt.WaitCursor)

        try:
            statsXLS.run_stats(file_path=file_path, team=team)
            self.append_output("Success, stats created.\n")
        except Exception as e:
            self.append_output("Error stats:")
            self.append_output(f'{str(e)}\n')

        # Set the cursor back to ArrowCursor
        QApplication.restoreOverrideCursor()

    # Run ProponisiUpdateProfiles.py to update profiles
    def update_proponisi_profiles(self):
        import ProponisiUpdateProfiles

        # ask for password
        password = self.get_password()
        
        if password != self.key:
            self.append_output("Invalid Key\n")
            return 
        else:
            self.append_output("Access Granted\n")

        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        self.append_output(
            "Steps to Follow To Update Proponisi Profiles:\n1. Choose a location: Boca or Ocoee.\n2. Pick the file that has Agents to be updated.\n3. Select the file detailing team alignments.\n4. A .txt document will be generated, providing strings suitable for Proponisi input."
        )

        location = self.get_proponisi_location()

        if not location:
            self.append_output("No Location Selected Process Cancelled.\n")
            return

        file_path, _ = file_dialog.getOpenFileName(
            self, "Pick the file that has Agents to be updated.", "", "Excel Files (*.xlsx)")
        if not file_path:
            self.append_output("File Not selected. Process Cancelled\n")
            return

        file_path_leads, _ = file_dialog.getOpenFileName(
            self, "Select the file detailing team alignments.", "", "Excel Files (*.xlsx)")
        if not file_path_leads:
            self.append_output("File Not selected. Process Cancelled\n")
            return

        try:
            ProponisiUpdateProfiles.update_proponisi_profiles_nesting_to_team(
                location, file_path, file_path_leads)
            self.append_output("Proponisi Update Profiles Ran.\n")
        except Exception as e:
            self.append_output("Error on Proponisi Update Profiles: ")
            self.append_output(f'{str(e)}\n')

    # Run time_on_phone.py to calculate time on queues
    def calculate_time_on_queues(self):
        import time_on_phone

        # ask for password
        password = self.get_password()
        
        if password != self.key:
            self.append_output("Invalid Key\n")
            return 
        else:
            self.append_output("Access Granted\n")

        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        self.append_output(
            "Steps to Follow To Calculate Agent's Total Time Per Queue:\n1. Select User Productivity Report in CSV format.\n2. A CSV File Will Be Created and Opened.")

        file_path1, _ = file_dialog.getOpenFileName(
            self, "Select Productivity Report CSV Format", "", "CSV Files (*.csv)")
        if not file_path1:
            self.append_output("File Not selected. Process Cancelled.\n")
            return
        try:
            response = time_on_phone.run_agent_time_report(file_path1)
            if response == 'Total Time Per Agent In Queues Report exported to: Total_Time_Per_Agent_In_Queues.csv':
                self.append_output("Time per Queue Ran.\n")
                pass
            else:
                self.append_output(
                    f'Error on time Per Queue: {str(response)}\n')
        except Exception as e:
            self.append_output(f'Error on time Per Queue: {str(e)}\n')

    # Pefect Attendance for Proponisi
    def mark_perfect_attedance_proponisi(self):
        import proponisi_perfect_attendance

        # ask for password
        password = self.get_password()
        
        if password != self.key:
            self.append_output("Invalid Key\n")
            return 
        else:
            self.append_output("Access Granted\n")

        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        self.append_output(
            "Steps to Follow To Mark Perfect Attendance on Proponisi Daily File:\n1. Select Proponisi File Created by the Proponisi Stats Button.\n2. Select File with Perfect Attendance.")

        file_path1, _ = file_dialog.getOpenFileName(
            self, "Select Proponisi File Created By 'Proponisi Stats' Button", "", "CSV Files (*.csv)")
        if not file_path1:
            self.append_output(
                "Proponisi File Not selected. Process Cancelled\n")
            return

        file_path2, _ = file_dialog.getOpenFileName(
            self, "Select the list of Agents with perfect attendance (CSV File)", "", "CSV Files (*.csv)")
        if not file_path2:
            self.append_output(
                "Perfect Attendance Agents File Not Selected. Process Cancelled\n")
            return

        try:
            proponisi_perfect_attendance.mark_attendance(
                file_path1, file_path2)
            self.append_output("Pefect Attendance Ran.\n")
        except Exception as e:
            self.append_output("Error sending quality assurance: ")
            self.append_output(f'{str(e)}\n')

    # Created proponisi profiles click
    def create_proponisi_profiles(self):
        import proponisiProfiles

        # ask for password
        password = self.get_password()
        
        if password != self.key:
            self.append_output("Invalid Key\n")
            return 
        else:
            self.append_output("Access Granted\n")

        self.append_output(
            "Steps to Follow To create Proponisi Profiles:\n1. Type Supervisor WIN ID.\n2. Type Location Boca or Ocoee.\n3. Team is Optional, Can Be Skipped\n4. Select File With the new agent's Profiles to Create.\n5. The file selected will be updated and opened with emails addresses, and a .txt file will be created and opened with a string ready for Proponsi.")

        supervisor = self.get_proponisi_winID()
        if not supervisor:
            self.append_output(
                "Supervisor WIN ID Not selected. Process Cancelled.\n")
            return

        location = self.get_proponisi_location()
        if not location:
            self.append_output("Location Not selected. Process Cancelled.\n")
            return

        team = self.get_team()

        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)

        file_path1, _ = file_dialog.getOpenFileName(
            self, "Select Profiles File in (.xlsx) format", "", "Excel Files (*.xlsx)")
        if not file_path1:
            self.append_output(
                "Proponisi profiles File Not selected. Process Cancelled.\n")
            return

        try:
            proponisiProfiles.createProponisiProfilesStringAndExcel(
                supervisor, location, file_path1, team)
            self.append_output("Success, Profiles .txt and .xlsx created.\n")
        except Exception as e:
            self.append_output("Error Profiles .txt and .xlsx created.")
            self.append_output(f'{str(e)}\n')

    # function to append output to the QTextEdit
    def append_output(self, text):
        self.output_text.append(text)

    # will orginize the none work codes report
    def non_work_report_click(self):
        from NonWorkCodes import run_report

        try:
            file_dialog = QFileDialog()
            file_dialog.setFileMode(QFileDialog.ExistingFile)
            self.append_output(
                "Steps to Follow To create Non Work Codes Report:\n1. Run User Availability Report from ICBM and export in CSV Format.\n2. Select out of the three files exported by ICBM, the one named table after the name provided when exporting it.\n3. Find and Open the Replica File Created, which contains all the Non Work Codes.")

            file_path, _ = file_dialog.getOpenFileName(
                self, "Select file", "", "CSV Files (*.csv)")

            if not file_path:
                self.append_output('No File Selected, Process Cancelled.\n')
                return

            output_text = run_report(file_path=file_path)
            self.append_output(f'{output_text}')
            self.append_output(
                'None Work Codes Report Created Successfully.\n')

        except Exception as e:
            self.append_output(f'{str(e)}\n')

    # will run main proponisi report
    def proponisi_report_click(self):
        import proponisi

        self.append_output(
            "Steps to Follow To Create Proponisi Daily File:\n1. Run User Productivity Report for Both Call Centers and Export in CSV Format.")

        date = self.get_date(QlableText='Enter Stats Date')
        if not date:
            self.append_output('No Date Entered, Process Cancelled.\n')
            return
        # date_object = datetime.strptime(date, "%m.%d.%Y")
        # Select first file
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file_path1, _ = file_dialog.getOpenFileName(
            self, "Select Stats File CSV Format", "", "CSV Files (*.csv)")
        if not file_path1:
            self.append_output('No File Selected, Process Cancelled.\n')
            return

        # Run your proponisi function with only one file
        try:
            output_text = proponisi.run_report(
                date_string=date, file_path1=file_path1)
            self.append_output(output_text)
            self.append_output('Proponisi Daily Report Ran Succesfully.\n')
        except Exception as e:
            self.append_output('Error Running Proponisi Daily Report.')
            self.append_output(f'{str(e)}\n')

    # add qa to proponisi report
    def qa_stats_proponisi_click(self):
        import proponisiQA

        # ask for password
        password = self.get_password()
        
        if password != self.key:
            self.append_output("Invalid Key\n")
            return 
        else:
            self.append_output("Access Granted\n")

        self.append_output(
            "Steps to Follow to Add QA Stats to the Proponisi Firday File:\n"
            "1. Select the 'Proponisi File' created by clicking the 'Proponisi Stats' button.\n"
            "2. Select the file named 'agents_list_do_not_delete', located in the same directory as this program.\n"
            "3. Export QA for Both Call Centers and Export them in CSV Format.\n"
            "4. Select the QA files."
        )

        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)

        file_path1, _ = file_dialog.getOpenFileName(
            self, "Select Proponisi File Created By 'Proponisi Stats' Button", "", "CSV Files (*.csv)")
        if not file_path1:
            self.append_output('Proponisi File Not Selected, Process Cancelled.\n')
            return
        self.append_output('1 / 4')

        file_path2, _ = file_dialog.getOpenFileName(
            self, "Select agents_list_do_not_delete Created By 'Proponisi Stats' Button", "", "CSV Files (*.csv)")
        if not file_path2:
            self.append_output('agents_list_do_not_delete File Not Selected, Process Cancelled.\n')
            return
        self.append_output('2 / 4')

        file_path3, _ = file_dialog.getOpenFileName(
            self, "Select the file Ocoee/Boca QA list, exported by ICBM, must be CSV format", "", "CSV Files (*.csv)")
        if not file_path3:
            self.append_output('QA File(s) Not Selected, Process Cancelled.\n')
            return
        self.append_output('3 / 4')

        file_path4, _ = file_dialog.getOpenFileName(
            self, "Select the file Ocoee/Boca QA list, exported by ICBM, must be CSV format", "", "CSV Files (*.csv)")
        if not file_path4:
            self.append_output(
                'You only selected 1 QA file, make sure to select Boca QA list and Ocoee QA list')
            try:
                output_text = proponisiQA.qa_stats_proponisi_click(
                file_path1=file_path1, file_path2=file_path2, file_path3=file_path3)
                self.append_output(output_text)
                self.append_output('Proponisi File with QA Ran Successfully.\n')
            except Exception as e:
                self.append_output('Error Running Proponisi Daily Report QA.')
                self.append_output(f'{str(e)}\n')
        else:
            self.append_output('4 / 4')
            try:
                output_text = proponisiQA.qa_stats_proponisi_click(
                file_path1=file_path1, file_path2=file_path2, file_path3=file_path3)
                self.append_output(output_text)
                self.append_output('Proponisi File with QA Ran Successfully.\n')
            except Exception as e:
                self.append_output('Error Running Proponisi Daily Report QA.')
                self.append_output(f'{str(e)}\n')

        return

    # report to send quality assurance
    def send_quality_assurance_click(self):
        import QA

        email_address = self.get_email()
        self.append_output(f'You Provided email: {email_address}')
        if not email_address:
            return

        date_range = self.get_date("Enter the date range, mm.dd - mm.dd:")
        self.append_output(f'You Provided date range: {date_range}')
        if not date_range:
            return

        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)

        file_path_1, _ = file_dialog.getOpenFileName(
            self, "File 1 is the QA in CSV format from ICBM", "", "CSV Files (*.csv)")
        if not file_path_1:
            return
        self.append_output('Selected QA file in CSV')

        file_path_2, _ = file_dialog.getOpenFileName(
            self, "Select agents_list_do_not_delete Created By 'Proponisi Stats' Button", "", "CSV Files (*.csv)")
        if not file_path_2:
            return
        self.append_output('Selected agents_list_do_not_delete file in CSV')

        try:
            QA.send_qa_to_work_email(
                file_path_1, file_path_2, date_range, email_address)
            self.append_output("Quality assurance sent successfully.")
        except Exception as e:
            self.append_output("Error sending quality assurance:")
            self.append_output(str(e))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
