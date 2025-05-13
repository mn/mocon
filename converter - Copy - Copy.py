################################
#                              #
#  ADVANCEDRX MOCON PROJECT    #
#     AUTOMATION SCRIPT        #
#                              #
################################
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
import errno

def wait_for_file_access(file_path, timeout=10, interval=0.5):
    """Wait until the file is accessible for reading."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            with open(file_path, 'r', encoding='latin-1') as file:
                return file.readlines()
        except IOError as e:
            if e.errno == errno.EACCES:
                time.sleep(interval)
            else:
                raise
    raise TimeoutError(f"Could not access {file_path} after {timeout} seconds.")

# Define the folder to monitor
WATCH_DIRECTORY = r"\\SERVER02\DATA\Profiles\Profile_CLIENT67\Desktop\Converter"
# Define the output folder for PDFs
OUTPUT_DIRECTORY = r"\\server01\DATA\Completed Formula Worksheets\Bulk Processing Batches - QA Sheets"

# Ensure the output directory exists
if not os.path.exists(OUTPUT_DIRECTORY):
    os.makedirs(OUTPUT_DIRECTORY)

print(f"\033[93mWatching directory: {WATCH_DIRECTORY}\033[0m")
print(f"\033[93mSaving PDFs to: {OUTPUT_DIRECTORY}\033[0m")

# Define a darker pastel yellow color
darker_pastel_yellow = colors.Color(0.77, 0.68, 0.25)

# PDF generation function
def create_pdf_from_prn(prn_file):
    base_filename = os.path.basename(prn_file).replace(".PRN", "")
    pdf_file = os.path.join(OUTPUT_DIRECTORY, base_filename + ".pdf")
    c = canvas.Canvas(pdf_file, pagesize=letter)
    
    left_margin = 72
    right_margin = 72
    top_margin = 72
    bottom_margin = 72
    
    width, height = letter
    
    def add_footer():
        pass  # Footer is currently disabled
        
    def new_page():
        add_footer()
        c.showPage()
        nonlocal y_position
        y_position = height - top_margin

    try:
        prn_content = wait_for_file_access(prn_file)
    except Exception as e:
        print(f"Error reading {prn_file}: {e}")
        return
    
    data_fields = {}
    test_number_found = False

    for line in prn_content:
        if not test_number_found:
            test_match = re.search(r"Test#:\s*(\d+)", line)
            if test_match:
                data_fields["Test#"] = test_match.group(1)
                test_number_found = True

        match = re.match(r"^(.*?):\s*(.*)$", line.strip())
        if match:
            key, value = match.groups()
            data_fields[key.strip()] = value.strip() if value else "(Not Provided)"
    
    if "Test#" not in data_fields:
        data_fields["Test#"] = "(Not Provided)"
    
    y_position = height - top_margin
    
    # Add Logo to Top Right Corner (Above Title)
    logo_path = "logo.png"
    if os.path.exists(logo_path):
        logo_width = 80
        logo_height = 40
        c.drawImage(logo_path, width - right_margin - logo_width, height - top_margin + 20, width=logo_width, height=logo_height, preserveAspectRatio=True, mask='auto')
    
    # Set title font and color to blue
    c.setFont("Helvetica-Bold", 20)
    c.setFillColor(colors.blue)  # Title in blue
    title_text = "MOCON Automatic Balance Analysis System"
    title_width = c.stringWidth(title_text, "Helvetica-Bold", 20)
    c.drawString((width - title_width) / 2, y_position, title_text)

    # Reset text color to black for the rest of the document
    c.setFillColor(colors.black)
    
    y_position -= 10
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)
    c.line(left_margin, y_position, width - right_margin, y_position)
    y_position -= 20
    
    # Report Details Section as Table
    report_details_data = [
        ["Report Details"],
        ["Test#: " + data_fields.get('Test#', '(Not Provided)'), "Date: " + data_fields.get("Date", "(Not Provided)"), "Time: " + data_fields.get("Time", "(Not Provided)")]
    ]
    
    # Adjust column widths to balance the table within the PDF
    report_details_table = Table(report_details_data, colWidths=[150, 150, 150])  # Adjust column widths as needed
    report_details_table.setStyle(TableStyle([
        ('SPAN', (0, 0), (-1, 0)),  # Merge all columns in the first row
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # Use grey color for the title row
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),  # Use light grey color for the data rows
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
    ]))
    
    report_details_table.wrapOn(c, width - left_margin - right_margin, height - top_margin - bottom_margin)
    report_details_table_width, report_details_table_height = report_details_table.wrap(width - left_margin - right_margin, height - top_margin - bottom_margin)
    
    if y_position < report_details_table_height + bottom_margin + 20:  # Add buffer space for the footer
        new_page()
    
    # Center the table on the page
    table_x_position = (width - report_details_table_width) / 2
    report_details_table.drawOn(c, table_x_position, y_position - report_details_table_height)
    y_position -= report_details_table_height + 20
    
    # Product Information Section as Mini Table
    product_info_data = [
        ["Product Information"],
        ["Product Name", data_fields.get("Product Name", "(Not Provided)")],
        ["Product Code", data_fields.get("Product Code", "(Not Provided)")],
        ["Batch Number", data_fields.get("Batch Number", "(Not Provided)")],
        ["Machine Number", data_fields.get("Machine Number", "(Not Provided)")],
        ["Container/Bin", data_fields.get("Container/Bin", "(Not Provided)")]
    ]
    
    product_info_table = Table(product_info_data, colWidths=[(width - left_margin - right_margin) / 2] * 2)
    product_info_table.setStyle(TableStyle([
        ('SPAN', (0, 0), (-1, 0)),  # Merge all columns in the first row
        ('BACKGROUND', (0, 0), (-1, 0), darker_pastel_yellow),  # Use darker pastel yellow color for the title row
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),  # Use light grey color for the data rows
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
    ]))
    
    product_info_table.wrapOn(c, width - left_margin - right_margin, height - top_margin - bottom_margin)
    product_info_table_width, product_info_table_height = product_info_table.wrap(width - left_margin - right_margin, height - top_margin - bottom_margin)
    
    # Analysis Data Section as Mini Table
    analysis_data = [
        ["Analysis Data"],
        ["Assay", data_fields.get("Assay", "(Not Provided)")],
        ["Target", data_fields.get("Target", "(Not Provided)")],
        ["Sort %", data_fields.get("Sort %", "(Not Provided)")],
        ["Shell Weight", data_fields.get("Shell Weight", "(Not Provided)")]
    ]
    
    # Calculate row heights to stretch the Analysis Data table to match the height of the Product Information table
    num_analysis_rows = len(analysis_data)
    row_heights_analysis = [product_info_table_height / num_analysis_rows] * num_analysis_rows
    
    analysis_table = Table(analysis_data, colWidths=[(width - left_margin - right_margin) / 2] * 2, rowHeights=row_heights_analysis)
    analysis_table.setStyle(TableStyle([
        ('SPAN', (0, 0), (-1, 0)),  # Merge all columns in the first row
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#A7C7E7")),  # pastel blue
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),  # Use light grey color for the data rows
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
    ]))
    
    analysis_table.wrapOn(c, width - left_margin - right_margin, height - top_margin - bottom_margin)
    analysis_table_width, analysis_table_height = analysis_table.wrap(width - left_margin - right_margin, height - top_margin - bottom_margin)
    
    # Ensure both tables have the same height
    max_table_height = max(product_info_table_height, analysis_table_height)
    
    if y_position < max_table_height + bottom_margin + 20:  # Add buffer space for the footer
        new_page()
    
    # Calculate the total width of both tables including the gap between them
    total_width = product_info_table_width + analysis_table_width + 20
    start_x = (width - total_width) / 2
    
    # Draw the Product Information table
    table_x_position = (width - product_info_table_width) / 2
    product_info_table.drawOn(c, table_x_position, y_position - product_info_table_height)
    y_position -= product_info_table_height + 20  # Add spacing between tables
    
    # Draw the Analysis Data table below
    table_x_position = (width - analysis_table_width) / 2
    if y_position < analysis_table_height + bottom_margin + 20:
        new_page()
    analysis_table.drawOn(c, table_x_position, y_position - analysis_table_height)
    y_position -= analysis_table_height + 20
    
    
    # Remove the black title "Sort Analysis Based on Target"
    
    # Extract and clean sort analysis data
    headers = []
    values = []
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
            if clean_line:  # Ensure clean_line is not empty
                if not headers:
                    headers = re.split(r'\s{2,}', clean_line)  # Split using two or more spaces
                else:
                    values = re.split(r'\s{2,}', clean_line)  # Split using two or more spaces

    # Ensure the table data is formatted correctly
    if len(headers) >= 4 and len(values) >= 3:
        # Correctly format the headers
        headers[1] = headers[1].replace("<", "").replace(">", "").replace("--", "").strip()
        headers[3] = headers[3].replace("<", "").replace(">", "").replace("--", "").strip()
        headers = [f"Reject (<) {headers[1]}", "Accepted Range", f"{headers[3]} (>) Reject"]
        
        table_data = [
            ["Sort Analysis Based on Target"],  # Add the new first row with the title
            headers,  # Use the entire formatted headers list
            values[:3]  # We only need the first three values for the table
        ]
        col_widths = [(width - left_margin - right_margin) / 3] * 3
        table = Table(table_data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('SPAN', (0, 0), (-1, 0)),  # Merge all columns in the first row
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),  # Use darkblue color for the title row
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('BACKGROUND', (0, 1), (-1, 2), colors.lightgrey),  # Make the second and third row the same color
            ('TEXTCOLOR', (0, 1), (-1, 2), colors.black),  # Make the text color black for the second and third row
            ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
        ]))
        table.wrapOn(c, width - left_margin - right_margin, height - top_margin - bottom_margin)
        _, table_height = table.wrap(width - left_margin - right_margin, height - top_margin - bottom_margin)
        table_x_position = (width - (3 * col_widths[0])) / 2
        if y_position < table_height + bottom_margin + 20:  # Add buffer space for the footer
            new_page()
        table.drawOn(c, table_x_position, y_position - table_height)
        y_position -= table_height + 20
    else:
        c.setFont("Helvetica-Bold", 10)
        draw_text("No valid sort analysis data found.", bold=True)
        y_position -= 20
    
    # Add custom field "Reweigh of Rejected Caps" in a table
    reweigh_data = [["Reweigh of Rejected Caps Results:"]]
    reweigh_table = Table(reweigh_data, colWidths=[width - left_margin - right_margin])
    reweigh_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6)
    ]))
    reweigh_table.wrapOn(c, width - left_margin - right_margin, height - top_margin - bottom_margin)
    _, reweigh_table_height = reweigh_table.wrap(width - left_margin - right_margin, height - top_margin - bottom_margin)
    if y_position < reweigh_table_height + bottom_margin + 20:  # Add buffer space for the footer
        new_page()
    reweigh_table.drawOn(c, left_margin, y_position - reweigh_table_height)
    y_position -= reweigh_table_height + 20

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
    std_dev_match = re.search(r"Std Dev:\s*([\d.]+ mg\s+\d+% Rel Std Dev)", prn_text)

    if max_dev_match and "Max Dev" not in data_fields:
        data_fields["Max Dev"] = max_dev_match.group(1)
    if std_dev_match and "Std Dev" not in data_fields:
        data_fields["Std Dev"] = std_dev_match.group(1)

    # Include Max Dev and Std Dev in the stats_data list
    if "Max Dev" in data_fields:
        stats_data.append(["Max Dev: " + data_fields["Max Dev"]])
    if "Std Dev" in data_fields:
        stats_data.append(["Std Dev: " + data_fields["Std Dev"]])

    # Format the data for the table
    formatted_stats_data = [[f"Samples: {samples_value}", ""]]  # Keep the first row unchanged
    for row in stats_data[1:]:
        parts = row[0].split(' ')
        if "Max Dev:" in row[0]:
            column1 = "Max Dev: " + ' '.join(parts[2:3]) + ' mg'
            column2 = ' '.join(parts[3:]).replace('mg ', '')  # Remove 'mg' from the right column if it exists
        elif "Std Dev:" in row[0]:  # Ensure Std Dev is handled properly
            column1 = "Std Dev: " + ' '.join(parts[2:3]) + ' mg'
            column2 = ' '.join(parts[3:]).replace('mg ', '')  # Remove 'mg' from the right column if it exists
        else:
            column1 = ' '.join(parts[:2]) + ' mg'
            column2 = ' '.join(parts[2:]).replace('mg ', '')  # Remove 'mg' from the right column if it exists
        formatted_stats_data.append([column1, column2])

    # Create the stats table
    stats_table = Table(formatted_stats_data, colWidths=[width/2 - left_margin, width/2 - right_margin])  # Adjust column widths
    stats_table.setStyle(TableStyle([
        ('SPAN', (0, 0), (-1, 0)),  # Merge all columns in the first row
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),  # Center the data in the left column
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),  # Center the data in the right column
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ]))

    # Make the "Samples" value bold
    for i, row in enumerate(formatted_stats_data):
        if "Samples:" in row[0]:
            stats_table.setStyle(TableStyle([
                ('FONTNAME', (0, i), (0, i), 'Helvetica-Bold')
            ]))

    stats_table.wrapOn(c, width - left_margin - right_margin, height - top_margin - bottom_margin)
    _, stats_table_height = stats_table.wrap(width - left_margin - right_margin, height - top_margin - bottom_margin)  # Calculate table height

    # Calculate x position to center the table
    table_width = width - left_margin - right_margin
    table_x_position = left_margin

    if y_position < stats_table_height + bottom_margin + 20:  # Add buffer space for the footer
        new_page()
    stats_table.drawOn(c, table_x_position, y_position - stats_table_height)
    y_position -= stats_table_height + 20

    weight_sections = re.findall(r"WEIGHTS IN\s+([^\n]+?)\s+RANGE\s*\n(.*?)(?=(?:\n\s*WEIGHTS IN\s+[^\n]+?\s+RANGE|\Z))", prn_text, re.DOTALL | re.IGNORECASE)
    print("Weight Sections Found:", len(weight_sections))
    for title, body in weight_sections:
        print(f"Section: {title} - Preview: {body[:100]}...")
        
    for range_name, values in weight_sections:
        values = values.strip()

        # Clean the weight lines and remove header noise
        clean_lines = [
            line.strip()
            for line in values.splitlines()
            if "MOCON Automatic Balance Analysis System" not in line
            and "Report includes" not in line
            and "Test#:" not in line
            and "Date:" not in line
            and "Time:" not in line
            and "--------" not in line
            and line.strip()
]

        if not clean_lines:
            clean_lines = ["N/A"]

        # Create the section title as a table
        title_data = [[f"Weights in {range_name} Range"]]
        title_table = Table(title_data, colWidths=[width - left_margin - right_margin])
        title_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ]))

        title_table.wrapOn(c, width, height)
        _, title_height = title_table.wrap(width, height)

        # Determine how many lines can fit on the current page
        line_height = 14
        available_height = y_position - bottom_margin - 20
        lines_per_page = int(available_height // line_height)
        chunked_lines = [clean_lines[i:i + lines_per_page] for i in range(0, len(clean_lines), lines_per_page)]

        for i, chunk in enumerate(chunked_lines):
            if y_position < title_height + line_height * len(chunk) + bottom_margin + 20:
                new_page()

            # Draw the title
            title_table.drawOn(c, left_margin, y_position - title_height)
            y_position -= title_height + 10

            if chunk:
                data_rows = [[line] for line in chunk]
                weight_table = Table(data_rows, colWidths=[width - left_margin - right_margin])
                weight_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), colors.beige),
                    ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))
                weight_table.wrapOn(c, width, height)
                _, table_height = weight_table.wrap(width, height)

                if y_position < table_height + bottom_margin + 20:
                    new_page()
                weight_table.drawOn(c, left_margin, y_position - table_height)
                y_position -= table_height + 20


    # Add custom field "Reweigh of Rejected Caps Weights" in a properly sized table
    reweigh_weights_data = [["Reweigh of Rejected Caps Weights:"]]
    # Adjust the number of rows to accommodate up to 10 weights
    for _ in range(10):
        reweigh_weights_data.append([""])

    reweigh_weights_table = Table(reweigh_weights_data, colWidths=[width - left_margin - right_margin])
    reweigh_weights_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6)
    ]))
    reweigh_weights_table.wrapOn(c, width - left_margin - right_margin, height - top_margin - bottom_margin)
    _, reweigh_weights_table_height = reweigh_weights_table.wrap(width - left_margin - right_margin, height - top_margin - bottom_margin)
    if y_position < reweigh_weights_table_height + bottom_margin + 20:  # Add buffer space for the footer
        new_page()
    reweigh_weights_table.drawOn(c, left_margin, y_position - reweigh_weights_table_height)
    y_position -= reweigh_weights_table_height + 20

    # Ensure "Reviewed by" and "Date Reviewed" appear just above the footer
    y_position = bottom_margin + 30  # Adjust position to be just above the footer
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin, y_position, "Reviewed by: ____________")
    c.drawRightString(width - right_margin, y_position, "Date Reviewed: ____________")

    add_footer()
    c.save()

    # Ensure the PDF opens on page 1
    with open(pdf_file, "r+b") as f:
        pdf_content = f.read()
        pdf_content = pdf_content.replace(b'/PageMode /UseNone', b'/PageMode /UseNone /Page 1')
        f.seek(0)
        f.write(pdf_content)
        f.truncate()

class PRNHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        if event.src_path.endswith(".PRN"):
            print(f"\033[92mEvent detected: created, Path: {event.src_path}\033[0m")
            retries = 5
            for attempt in range(retries):
                try:
                    create_pdf_from_prn(event.src_path)
                    print(f"\033[94mPDF saved to {OUTPUT_DIRECTORY}\033[0m")
                    break
                except PermissionError as e:
                    print(f"Retrying... ({attempt + 1}/{retries})")
                    time.sleep(2)  # Wait 2 seconds before retrying
                except Exception as e:
                    print(f"Error processing {event.src_path}: {e}")
                    break

observer = Observer()
handler = PRNHandler()
observer.schedule(handler, WATCH_DIRECTORY, recursive=False)
observer.start()
try:
    while True:
        time.sleep(1)  # Sleep for 1 second
except KeyboardInterrupt:
    observer.stop()
observer.join()