ITM-352 Final Project: Automated Car Log Uploader

🚗 Project Summary
This program automates the process of entering car log data from an Excel file into the UH travel log website. It features a simple graphical interface (GUI) and creates a detailed log file for tracking results.

🛠️ Required Setup (How to Run the App)
The application requires Python 3 and the Google Chrome browser to be installed.

Step 1: Get the Code
Open your computer's terminal (or Command Prompt) and run these two commands to download the project and enter the correct folder:


# 1. Downloads the project files
git clone https://github.com/ankabayanbat/ITM-352-Final-Project

# 2. Move into the folder where the program is located
cd ITM-352-Final-Project
Step 2: Install Necessary Tools
Now, install all the required Python libraries (like Pandas and Selenium) in a clean environment.


# 1. Creates a clean environment (Optional, but best practice)
python -m venv venv

# 2. Activates the clean environment (run ONE of the following based on your system)
# Windows:
.\venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# 3. Installs all required libraries
pip install pandas selenium webdriver-manager openpyxl
Step 3: Run the Program

Start the application:
python cargui19.py

Step 4: Using the GUI
Login: Use Username: anka and Password: anka123

Run: Click the "▶ Start Upload" button. The application will automatically launch Chrome and begin filling out the web form using the excel data.

✅ Final Output
After the program finishes, a file named Submission_Log.csv will appear in the project folder.

This log is important because it shows you two things side-by-side:

The value that was in the Excel file.
The exact value that the program successfully selected on the website.

🤖 Use of AI
I used three different AI—ChatGPT, Gemini, and Claude for key aspects of this project. AI was critical for:

Debugging: It helped fix complex errors in finding specific elements on the website by suggesting robust Selenium XPath solutions.

Safety & Auditability: It ensured the final logging feature works correctly by architecting the solution to track and compare the actual data selected by the program versus the initial Excel input.

Structure: It helped structure the code to ensure the app runs smoothly, specifically advising on the use of threading to prevent the graphical interface (GUI) from freezing while the automation runs.
