from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
import os
import time
import re
from datetime import datetime

# Define the folder to monitor
WATCH_DIRECTORY = r"C:\Users\client27\Desktop\Converter"

print(f"Watching directory: {WATCH_DIRECTORY}")

# PDF generation function
def create_pdf_from_prn(prn_file):
    pdf_file = prn_file.replace(".PRN", ".pdf")
    c = canvas.Canvas(pdf_file, pagesize=letter)
    
    left_margin = 72
    right_margin = 72
    top_margin = 72
    bottom_margin = 72
    
    width, height = letter

    # Add page number and generation timestamp to each page
    generation_time_string = "Generated on: 5-03-11 15:02:45"
    page_number = 1
    
    def add_footer():
        nonlocal page_number
        c.setFont("Helvetica", 8)
        # Add generation timestamp to the left side
        c.drawString(left_margin, bottom_margin - 10, generation_time_string)
        # Add page number to the right side
        page_text = f"Page {page_number}"
        c.drawRightString(width - right_margin, bottom_margin - 10, page_text)
        page_number += 1
    
    try:
        with open(prn_file, 'r', encoding='latin-1') as file:
            prn_content = file.readlines()
    except Exception as e:
        print(f"Error reading {prn_file}: {e}")
        return
    
    data_fields = {}

    for line in prn_content:
        test_match = re.search(r"Test#:\s*(\d+)", line)
        if test_match:
            data_fields["Test#"] = test_match.group(1)

        match = re.match(r"^(.*?):\s*(.*)$", line.strip())
        if match:
            key, value = match.groups()
            data_fields[key.strip()] = value.strip() if value else "(Not Provided)"
    
    y_position = height - top_margin
    
    # Add Logo to Top Right Corner (Above Title)
    logo_path = "logo.png"
    if os.path.exists(logo_path):
        logo_width = 80
        logo_height = 40
        c.drawImage(logo_path, width - right_margin - logo_width, height - top_margin + 20, width=logo_width, height=logo_height, preserveAspectRatio=True, mask='auto')
    
    # Set title font and color to blue
    c.setFont("Helvetica-Bold", 16)
    c.setFillColor(colors.blue)  # Title in blue
    title_text = "MOCON Automatic Balance Analysis System"
    title_width = c.stringWidth(title_text, "Helvetica-Bold", 16)
    c.drawString((width - title_width) / 2, y_position, title_text)

    # Reset text color to black for the rest of the document
    c.setFillColor(colors.black)
    
    y_position -= 10
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)
    c.line(left_margin, y_position, width - right_margin, y_position)
    y_position -= 20
    
    def draw_section(title):
        nonlocal y_position
        if y_position < bottom_margin:
            add_footer()
            c.showPage()
            y_position = height - top_margin
        c.setFont("Helvetica-Bold", 12)
        title_width = c.stringWidth(title, "Helvetica-Bold", 12)
        c.drawString((width - title_width) / 2, y_position, title)
        y_position -= 15
    
    def draw_text(text, font="Helvetica", size=10, bold=False):
        nonlocal y_position
        if y_position < bottom_margin:
            add_footer()
            c.showPage()
            y_position = height - top_margin
        if bold:
            c.setFont("Helvetica-Bold", size)
        else:
            c.setFont(font, size)
        text_width = c.stringWidth(text, font if not bold else "Helvetica-Bold", size)
        c.drawString((width - text_width) / 2, y_position, text)
        y_position -= 12

    draw_section("Report Details")
    draw_text(f"Test#: {data_fields.get('Test#', '(Not Provided)')}")
    for key in ["Date", "Time"]:
        if key in data_fields:
            draw_text(f"{key}: {data_fields[key]}")
    y_position -= 10
    
    draw_section("Product Information")
    for key in ["Product Name", "Product Code", "Batch Number", "Machine Number", "Container/Bin"]:
        if key in data_fields:
            draw_text(f"{key}: {data_fields[key]}")
    y_position -= 10
    
    draw_section("Analysis Data")
    for key in ["Assay", "Target", "Sort %", "Shell Weight", "Cut Point Analysis Based On"]:
        if key in data_fields:
            draw_text(f"{key}: {data_fields[key]}")
    y_position -= 10
    
    draw_section("Sort Analysis Based on Target")
    sort_analysis_lines = []
    sort_analysis_start = False
    for line in prn_content:
        if "Sort Analysis Based on Target" in line:
            sort_analysis_start = True
            continue
        if sort_analysis_start:
            if "Samples:" in line:
                break
            clean_line = re.sub(r'[^\x20-\x7E]', '', line).strip()
            clean_line = clean_line.replace("?", "").strip()
            clean_line = clean_line.replace("-- ", "(<)")
            clean_line = clean_line.replace("-->", "(>)")
            if clean_line:
                sort_analysis_lines.append(clean_line)
    for text in sort_analysis_lines:
        draw_text(text)
    y_position -= 10
    
    # Add custom field "Reweigh of Rejected Caps"
    draw_text("Reweigh of Rejected Caps: ____________", bold=True)
    y_position -= 10
    
    # Remove the "Statistical Summary" title and push the table up
    samples_value = data_fields.get("Samples", "(Not Provided)")
    stats_data = [
        [f"Samples: {samples_value}"],
        ["Mean: " + data_fields.get("Mean", "(Not Provided)")],
        ["Maximum: " + data_fields.get("Maximum", "(Not Provided)")],
        ["Minimum: " + data_fields.get("Minimum", "(Not Provided)")]
    ]
    
    prn_text = "\n".join(prn_content)
    max_dev_match = re.search(r"Max Dev:\s*([\d.]+ mg\s+\d+\.\d+% of target)", prn_text)
    std_dev_match = re.search(r"Std Dev:\s*([\d.]+ mg\s+\.\d+% Rel Std Dev)", prn_text)

    if max_dev_match:
        stats_data.append(["Max Dev: " + max_dev_match.group(1)])
    if std_dev_match:
        stats_data.append(["Std Dev: " + std_dev_match.group(1)])
    
    stats_table = Table(stats_data, colWidths=[200])  # Reduced the width of the table columns
    stats_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Center the whole table
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ]))

    # Make the "Samples" value bold
    for i, row in enumerate(stats_data):
        if "Samples:" in row[0]:
            stats_table.setStyle(TableStyle([
                ('FONTNAME', (0, i), (0, i), 'Helvetica-Bold')
            ]))
    
    stats_table.wrapOn(c, width - left_margin - right_margin, height - top_margin - bottom_margin)
    _, stats_table_height = stats_table.wrap(width - left_margin - right_margin, height - top_margin - bottom_margin)  # Calculate table height
    
    # Calculate x position to center the table
    table_width = 200
    table_x_position = (width - table_width) / 2
    
    if y_position < stats_table_height + bottom_margin:
        add_footer()
        c.showPage()
        y_position = height - top_margin
    stats_table.drawOn(c, table_x_position, y_position - stats_table_height)
    y_position -= stats_table_height + 20
    
    weight_sections = re.split(r"WEIGHTS IN (.*?) RANGE", prn_text)[1:]
    for i in range(0, len(weight_sections), 2):
        range_name = weight_sections[i].strip()
        values = weight_sections[i + 1].strip()

        # Create a table for the weights, removing lines with "----------------------------------------------------------------"
        table_data = [[f"Weights in {range_name} Range"]]
        wrapped_lines = re.findall(r"(.{1,80})(?:\s|$)", values)
        clean_lines = [line.strip() for line in wrapped_lines if "----------------------------------------------------------------" not in line]
        
        if not clean_lines:
            # If there's no data, insert a row with "N/A"
            clean_lines.append("N/A")
        
        for line in clean_lines:
            table_data.append([line])
        
        table = Table(table_data, colWidths=[width - left_margin - right_margin])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))
        
        table.wrapOn(c, width - left_margin - right_margin, height - top_margin - bottom_margin)
        _, table_height = table.wrap(width - left_margin - right_margin, height - top_margin - bottom_margin)  # Calculate table height
        if y_position < table_height + bottom_margin:
            add_footer()
            c.showPage()
            y_position = height - top_margin
        table.drawOn(c, left_margin, y_position - table_height)
        y_position -= table_height + 20
    
    add_footer()
    c.save()
    print(f"PDF saved as {pdf_file}")

class PRNHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        if event.src_path.endswith(".PRN"):
            print(f"Event detected: created, Path: {event.src_path}")
            create_pdf_from_prn(event.src_path)

observer = Observer()
handler = PRNHandler()
observer.schedule(handler, WATCH_DIRECTORY, recursive=False)
observer.start()
try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()
observer.join()
