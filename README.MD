### This script uses custom-made excel document to download tests and merge them based on their corresponding lables

To run this script:
- Download and install python (https://www.youtube.com/watch?v=4Rx_JRkwAjY, watch up to 3:30) don't forget "Add Python to PATH" checkbox
- Download chromedriver (https://chromedriver.chromium.org/downloads) appropriate for your Google Chrome version (you can check verion in Settings -> Help -> About Google Chrome)
- Copy chromedriver.exe to your python installation folder (so it is alongside python.exe file)
- Download and start PyCharm
- Setup your configuration and interpreter (https://www.youtube.com/watch?v=56bPIGf4us0, watch up to 3:50) 
- Open in PyCharm Jira_downloads.py file, you should get prompt about missing requirements, install them
- At the 22'nd and 23'rd line of the script you will find empty quotes, fill them with your username and password
- Run command prompt and type: pip install xlrd
- Run the script