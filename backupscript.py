from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
import os
import time
import re

# Define the folder to monitor
WATCH_DIRECTORY = r"C:\Users\malno\OneDrive\Desktop/Converter"

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
    center_x = width / 2
    
    try:
        with open(prn_file, 'r', encoding='latin-1') as file:
            prn_content = file.readlines()
    except Exception as e:
        print(f"Error reading {prn_file}: {e}")
        return
    
    data_fields = {}

    # Fix: Ensure "Test#" is captured properly
    for line in prn_content:
        test_match = re.search(r"Test#:\s*(\d+)", line)
        if test_match:
            data_fields["Test#"] = test_match.group(1)

        match = re.match(r"^(.*?):\s*(.*)$", line.strip())
        if match:
            key, value = match.groups()
            data_fields[key.strip()] = value.strip() if value else "(Not Provided)"
    
    y_position = height - top_margin
    
    c.setFont("Helvetica-Bold", 16)
    title_text = "MOCON Automatic Balance Analysis System"
    title_width = c.stringWidth(title_text, "Helvetica-Bold", 16)
    c.drawString(center_x - (title_width / 2), y_position, title_text)
    y_position -= 10
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)
    c.line(left_margin, y_position, width - right_margin, y_position)
    y_position -= 20
    
    def draw_section(title):
        nonlocal y_position
        if y_position < bottom_margin:
            c.showPage()
            y_position = height - top_margin
        c.setFont("Helvetica-Bold", 12)
        title_width = c.stringWidth(title, "Helvetica-Bold", 12)
        c.drawString(center_x - (title_width / 2), y_position, title)
        y_position -= 15
    
    def draw_text(text):
        nonlocal y_position
        if y_position < bottom_margin:
            c.showPage()
            y_position = height - top_margin
            c.setFont("Helvetica", 10)
        c.setFont("Helvetica", 10)
        text_width = c.stringWidth(text, "Helvetica", 10)
        c.drawString(center_x - (text_width / 2), y_position, text)
        y_position -= 12

    draw_section("Report Details")
    draw_text(f"Test#: {data_fields.get('Test#', '(Not Provided)')}")  # Fix applied
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
    
    # Extract Sort Analysis Based on Target
    sort_analysis_lines = []
    sort_analysis_start = False
    for line in prn_content:
        if "Sort Analysis Based on Target" in line:
            sort_analysis_start = True
            continue
        if sort_analysis_start:
            if "Samples:" in line:
                break
            clean_line = re.sub(r'[^\x20-\x7E]', '', line).strip()  # Remove non-ASCII characters
            clean_line = clean_line.replace("?", "").strip()  # Remove extra "?" characters
            clean_line = clean_line.replace("-- ", "(<)")  # Replace "-- " with "(<)"
            clean_line = clean_line.replace("-->", "(>)")  # Replace "-->" with "(>)"
            if clean_line:
                sort_analysis_lines.append(clean_line)

    sort_analysis_text = "\n".join(sort_analysis_lines) if sort_analysis_lines else "(No Data Available)"
    
    draw_section("Sort Analysis Based on Target")
    c.setFont("Helvetica", 10)
    for text in sort_analysis_text.split("\n"):
        draw_text(text)
    y_position -= 10
    
    draw_section("Statistical Summary")
    for key in ["Samples", "Mean", "Maximum", "Minimum"]:
        if key in data_fields:
            draw_text(f"{key}: {data_fields[key]}")

    prn_text = "\n".join(prn_content)
    max_dev_match = re.search(r"Max Dev:\s*([\d.]+ mg\s+\d+\.\d+% of target)", prn_text)
    std_dev_match = re.search(r"Std Dev:\s*([\d.]+ mg\s+\d+\.\d+% Rel Std Dev)", prn_text)

    if max_dev_match:
        draw_text(f"Max Dev: {max_dev_match.group(1)}")
    if std_dev_match:
        draw_text(f"Std Dev: {std_dev_match.group(1)}")

    y_position -= 20

    # Extract and format "WEIGHTS IN RANGE" sections dynamically
    weight_sections = re.split(r"WEIGHTS IN (.*?) RANGE", prn_text)[1:]
    for i in range(0, len(weight_sections), 2):
        range_name = weight_sections[i].strip()
        values = weight_sections[i + 1].strip()

        # Draw section header
        draw_section(f"Weights in {range_name} Range")
        
        # Remove the solid line **ONLY** for specified ranges
        if range_name not in ["-T3", "-T2", "-T1", "T0", "+T1", "+T2", "+T3"]:
            c.setStrokeColor(colors.black)
            c.setLineWidth(1)
            c.line(left_margin, y_position, width - right_margin, y_position)
            y_position -= 12

        # Wrap text to keep it readable
        wrapped_lines = re.findall(r"(.{1,80})(?:\s|$)", values)  
        for line in wrapped_lines:
            draw_text(line.strip())

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
