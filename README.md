# Setup
In order to setup your environemnt, you need to: \
0. Clone this repo to your local machine.
1. Have python installed on your computer.
2. Install a few Python packages. These are found in requirments.txt, and can be downloaded with \
```pip install -r requirments.txt```
3. Download the chromedriver binary matching your computers chrome version and OS, and place the binary in your PATH. It would be easiest to place the driver in the same directory as this repo. The drivers can be found at \
https://googlechromelabs.github.io/chrome-for-testing/
4. In runner.py, update building to your building name.

# Running
1. Fill out the `residents.xlsx` excel spreadsheet with your interactions row by row. Be not to change the headers and that everythign is spelled and formatted correctly.
2. Run the runner.py python file. \
   ```python runner.py ``` or ```python3 runner.py```
3. If things are working correctly, a new tab should open to the roompact sign in page. Sign in, and then type `y` and press enter in your terminal.
4. Selenium should take over your chrome tab, filling in the information from your spreadsheet into the roompact form.

# Troubleshooting
Is you run into issues, feel free to email or teams me at gmanaster3@gatech.edu
