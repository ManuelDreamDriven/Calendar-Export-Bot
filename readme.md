# Calendar Events Bot

This bot helps you manage and process calendar events from an ICS file and outputs them into an Excel spreadsheet.

## Overview

The bot takes an ICS file (a common format for calendar data) and converts the events into an Excel file named `calendar_events.xlsx`. If the output file already exists, it will be deleted and regenerated.

## Software

To run this bot, you need the following software:

1. **Python**: Make sure you have Python 3 installed on your computer. You can download it from [python.org](https://www.python.org/downloads/).
2. **Code Editor**: You will need a code editor to edit the `main.py` file. Recommended options include:
   - [Visual Studio Code](https://code.visualstudio.com/)
   - [Sublime Text](https://www.sublimetext.com/)
   - [Atom](https://atom.io/)

## Instructions

### Step 1: Download the Bot

1. Download the ZIP file from the GitHub repository.
2. Extract the contents of the ZIP file to a folder on your computer.

### Step 2: Set Up Your Environment

1. **Install a Code Editor**: If you don’t have a code editor, you can download one
   
2. **Install Required Packages**: 
   - Open your terminal (Command Prompt for Windows, Terminal for macOS).
   - Navigate to the folder where you extracted the files.
   - Install the necessary packages by running the following command:
     ```bash
     pip install -r requirements.txt
     ```

### Step 3: Prepare Your ICS File

1. Locate your ICS file.
2. Open the `main.py` file in your code editor.
3. Find the line where `ics_file` is defined, and insert the path to your ICS file:
   ```python
   ics_file = 'path/to/your/file.ics'


### Step 3: Run the bot
For macOS, use the following command in your terminal:
    ```python
    python3 main.py 

For Windows, use this command:
    ```python
    python main.py


### Output
After running the bot, you will find a new file named calendar_events.xlsx in the same directory. This file contains all the events from your ICS file formatted for easy viewing.

### Troubleshooting
If you encounter any errors, ensure that you have installed all the required packages and that the path to your ICS file is correct.
Feel free to reach out if you have any questions or need further assistance!