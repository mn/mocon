MOCON Project’s End of Day Update Version:
## Latest Version Update 3/13
+ Updated Sections and Graphs
  
![Alt text](https://i.imgur.com/M1Qug5v.png)

## Latest Version Update 3/13
+ Updated Graphs
+ Moved Date Signed & Reviewed Placeholders
  
![Alt text](https://i.imgur.com/bHQ6wpb.png)

## Latest Version Update 3/12
+ Updated Graphs
+ Updated Custom Fields
+ Added Custom Fields
  
![Alt text](https://i.imgur.com/zAInUlr.png)

## Latest Version Update 3/11
+ Added Graphs
+ Added Custom Field
  
![Alt text](https://i.imgur.com/F3y1UWq.png)

## (Archived) Old Version Update 3/10
![Alt text](https://i.imgur.com/UaND1PU.png)

## Example Logic
![Alt text](https://i.imgur.com/nkBmLUO.png)

# PRN-to-PDF Automated Reporting System

## Project Overview

### Scope/Objective
The objective of this project is to develop a Python-based automation script (bot) to handle the conversion of `.prn` files into `.pdf` format. The solution will run on Windows and will be executed from the command prompt, with task scheduling managed by **Windows Scheduler**.

### Features

#### 1. **File Monitoring & Conversion:**
- The script will continuously monitor a specified directory (filepath) for incoming `.prn` files.
- Upon detecting a new `.prn` file, the script will process the file, extracting data by reading it in a sequential manner (top to bottom, left to right).
- The extracted data will then be converted into a `.pdf` file and saved to a designated folder that syncs with **Google Drive** for easy access and storage.

#### 2. **Configuration File:**
- A configuration file will be created where admins can specify:
  - The directory to scan for `.prn` files.
  - The target location for saving the converted `.pdf` files.
- The script will read and follow these configurations to ensure it operates as expected.

#### 3. **Error Handling::**
- The script will handle unexpected errors, such as issues with file encoding or structure, by skipping the problematic `.prn` file and continuing to monitor for new files.
- In case of an error, the script will raise an error message, log it for debugging, and stop further processing of the current file, but will not halt the entire monitoring loop. A delay mechanism (e.g., 10 seconds) will be implemented to prevent system overload during the scanning process.

#### 4. **Continuous Operation:**
- The script will run in an infinite loop, continuously scanning the configured folder for new `.prn` files.
- The loop will have a slight delay (roughly 10 seconds) to prevent resource overload while ensuring files are processed in a timely manner.

### Goal:
The goal of this project is to automate the `.prn` to `.pdf` conversion process in an efficient and fault-tolerant manner, ensuring reliable and seamless integration with **Google Drive**. The script will be simple to configure and flexible enough to handle potential errors, enabling continuous operation without manual intervention.
