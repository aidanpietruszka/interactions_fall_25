#### Note: This script is configured for the Fall 24 form. It will be updated to the Spring 25 form once Roompact is updated.
# Setup
In order to setup your environment, you need to open your terminal command or command line and do the following:
1. Clone this repo to your local machine. \
   ```git clone https://github.com/gmanaster54/interactions_spring_25.git ```
2. Cd into the repo with \
   ```cd interactions_spring_25```
4. Makse sure python is installed on your computer.  \
   Running ```python --version``` or ```python3 --version``` should display `python 3.*`
5. Install a few Python packages. These are found in requirments.txt, and can be downloaded with \
```pip install -r requirements.txt```
6. Download the chromedriver binary matching both your computers chrome version and OS, and place the binary in your PATH. It would be easiest to place the driver in the same directory as this repo. The drivers can be found at \
https://googlechromelabs.github.io/chrome-for-testing/
7. In runner.py, update `building` to your building name. It is currently set to `North Avenue East`.

# Running
1. Fill out the `residents.xlsx` excel spreadsheet with your interactions row by row. Be sure not to change the headers and that everything is spelled and formatted correctly.
2. Run the runner.py python file. \
   ```python runner.py ``` or ```python3 runner.py```
3. If things are working correctly, a new tab should open to the roompact sign in page. Sign in, and then type `y` and press enter in your terminal.
4. Selenium should take over your chrome tab, filling in the information from your spreadsheet into the roompact form.
5. Check the output in the terminal to ensure that each interaction is properly submitted. It will output: \
``` Submitted: <Resident Name> ``` for each sucessful submission.  
# Troubleshooting
Is you run into issues, feel free to email or teams me at gmanaster3@gatech.edu
