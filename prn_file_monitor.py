import time

from watchdog.observers import Observer

from watchdog.events import FileSystemEventHandler

from reportlab.lib import colors

from reportlab.lib.pagesizes import letter

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph

from reportlab.lib.styles import getSampleStyleSheet

import os

import shutil

from datetime import datetime

 

class PRNHandler(FileSystemEventHandler):

    def __init__(self, input_folder, output_folder):

        self.input_folder = input_folder

        self.output_folder = output_folder

 

    def on_created(self, event):

        if not event.is_directory and event.src_path.endswith('.PRN'):

            print(f"New PRN file detected: {event.src_path}")

            self.process_prn_file(event.src_path)

 

    def parse_prn_file(self, file_path):

        data = {

            'test_number': '',

            'date': '',

            'product_details': '',

            'statistics': [],

            'weight_ranges': []

        }

 

        try:

            with open(file_path, 'r') as file:

                lines = file.readlines()

               

                # This is a placeholder parsing logic - adjust according to actual PRN format

                for line in lines:

                    if line.startswith('Test#'):

                        data['test_number'] = line.split(':')[1].strip()

                    elif line.startswith('Date'):

                        data['date'] = line.split(':')[1].strip()

                    elif line.startswith('Product'):

                        data['product_details'] = line.split(':')[1].strip()

                    # Add more parsing logic based on PRN file structure

                   

        except Exception as e:

            print(f"Error parsing PRN file: {e}")

           

        return data

 

    def generate_pdf(self, data, output_path):

        doc = SimpleDocTemplate(output_path, pagesize=letter)

        styles = getSampleStyleSheet()

        elements = []

 

        # Title

        title = Paragraph(f"Test Report #{data['test_number']}", styles['Heading1'])

        elements.append(title)

 

        # Basic Information

        basic_info = [

            ['Date:', data['date']],

            ['Product:', data['product_details']]

        ]

       

        basic_table = Table(basic_info)

        basic_table.setStyle(TableStyle([

            ('GRID', (0, 0), (-1, -1), 1, colors.black),

            ('PADDING', (0, 0), (-1, -1), 6),

        ]))

        elements.append(basic_table)

 

        # Statistics Table

        if data['statistics']:

            stats_title = Paragraph("Statistics", styles['Heading2'])

            elements.append(stats_title)

            stats_table = Table(data['statistics'])

            stats_table.setStyle(TableStyle([

                ('GRID', (0, 0), (-1, -1), 1, colors.black),

                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),

                ('PADDING', (0, 0), (-1, -1), 6),

            ]))

            elements.append(stats_table)

 

        # Build the PDF

        doc.build(elements)

 

    def process_prn_file(self, file_path):

        try:

            # Parse the PRN file

            data = self.parse_prn_file(file_path)

 

            # Generate output filename

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            output_filename = f"Report_{timestamp}.pdf"

            output_path = os.path.join(self.output_folder, output_filename)

 

            # Generate PDF

            self.generate_pdf(data, output_path)

            print(f"PDF report generated: {output_path}")

 

        except Exception as e:

            print(f"Error processing file: {e}")

 

def start_monitoring(input_folder, output_folder):

    # Create output folder if it doesn't exist

    os.makedirs(output_folder, exist_ok=True)

 

    # Initialize event handler and observer

    event_handler = PRNHandler(input_folder, output_folder)

    observer = Observer()

    observer.schedule(event_handler, input_folder, recursive=False)

    observer.start()

 

    try:

        print(f"Starting monitoring of folder: {input_folder}")

        print(f"PDF reports will be saved to: {output_folder}")

        while True:

            time.sleep(1)

    except KeyboardInterrupt:

        observer.stop()

        print("\nMonitoring stopped")

   

    observer.join()
