/*
*=========================================
*            GAS-ICS-Sync v6.0
*-----------------------------------------
*              INSTRUCTIONS
*=========================================
*
* 1) Make a local Copy:
*      - Click on the project overview icon on the left (ⓘ icon).
*      - Click the "Copy" icon (two files on top of each other) on the top right to make a copy of the project.
*
* 2) Configure Settings:
*      - Open the `Settings.gs` file.
*      - Adjust the configuration options as per your requirements (e.g., sync interval, email settings).
*
* Note: For the following steps, ensure `Main.gs` is open.
*
* 3) Install the Script:
*      - Ensure that the dropdown menu to the right of "Debug" says "install".
*      - Click the "Run" button to install the script.
*      - After Installation is complete, a sync will be started automatically.
*
* 4) Authorize the Script:
*      - You will be prompted to authorize the script.
*      - Click "Advanced" > "Go to GAS-ICS-Sync (unsafe)".
*      - Follow the on-screen prompts to complete the authorization.
*      - (For detailed steps, refer to this video: https://youtu.be/_5k10maGtek?t=1m22s).
*
* 5) Sync Calendars:
*      - To start the sync immediately, change the dropdown menu to the right of "Debug" from "install" to "startSync" and click "Run".
*      - The script will now sync your calendars based on the configured settings.
*
* 6) Uninstall the Script:
*      - To stop the script, change the dropdown menu to the right of "Debug" from "install" to "uninstall" and click "Run".
*
*=========================================
*              HOW TO UPDATE
*=========================================
*
* To update the script, you have two options:
*
* Option 1: Manually Update Files
*      - Go to the GitHub repository: https://github.com/derekantrican/GAS-ICS-Sync
*      - Open your existing Google Apps Script project.
*      - Manually replace the old files with the new ones from GitHub.
*
* Option 2: Create a New Copy
*      - Go to the GitHub repository: https://github.com/derekantrican/GAS-ICS-Sync
*      - Click on the link to the latest release.
*      - Make a copy of the latest release.
*      - Open your existing `Settings.gs` file and copy your current settings.
*      - Paste the copied settings into the `Settings.gs` file of the new project.
*      - Follow the installation steps above to reinstall and authorize the new script.
*
* Option 3: Use Git and clasp (Advanced)
*      - Install Git: https://git-scm.com/
*      - Install clasp (Command Line Apps Script): https://github.com/google/clasp
*      - Clone the repository locally: 'git clone https://github.com/derekantrican/GAS-ICS-Sync.git'
*      - Navigate to the project directory: 'cd GAS-ICS-Sync'
*      - Log in to your Google account with clasp: 'clasp login'
*      - Create a new Google Apps Script project or use an existing one and link it with clasp: 'clasp create' or 'clasp clone <SCRIPT_ID>'
*      - Create a '.claspignore' file in the project directory and add 'Settings.gs' to prevent your settings from being overwritten
*      - Push the local files to Google Apps Script: 'clasp push'
*      - Follow the installation steps above to reinstall and authorize the updated script.
*
*=========================================
*           ABOUT THE AUTHOR
*=========================================
*
* This program was created by Derek Antrican
*
* If you would like to see other programs Derek has made, you can check out
* his website: derekantrican.com or his github: https://github.com/derekantrican
*
*=========================================
*            BUGS/FEATURES
*=========================================
*
* Please report any issues at https://github.com/derekantrican/GAS-ICS-Sync/issues
*
*=========================================
*           $$ DONATIONS $$
*=========================================
*
* If you would like to donate and support the project,
* you can do that here: https://www.paypal.me/jonasg0b1011001
*
*=========================================
*             CONTRIBUTORS
*=========================================
* Andrew Brothers
* Github: https://github.com/agentd00nut
* Twitter: @abrothers656
*
* Joel Balmer
* Github: https://github.com/JoelBalmer
*
* Blackwind
* Github: https://github.com/blackwind
*
* Jonas Geissler
* Github: https://github.com/jonas0b1011001
*/