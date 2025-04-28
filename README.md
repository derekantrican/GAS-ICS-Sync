# GAS-ICS-Sync

### ⚠️⚠️ This project is looking for contributors and people to help answer questions! Please message @developers on the Discord! ⚠️⚠️

This is a standalone script (that consists of multiple files). The purpose is to sync ics/ical calendars to Google Calendar. Google Calendar *can* already do this, but updates only happen once every 12 or even 24 hrs. This script can be run much more frequently.

[If you want to use this, please copy the script from here](https://script.google.com/d/1BOk8MDLbLaHh6SwG1M1tsgNXjkcC-79LE0QoipRuTDxbO3fMVvqoROQD/edit?newcopy=true)

**To make a copy in the new Google Apps Script interface:**
1. Go to the project overview icon on the left (looks like this: ⓘ)
2. Click the "copy" icon on the top right (looks like two files on top of each other)

**NOTE:** If too many people are accessing the file at the same time, Google may lock you out. You can follow these instructions to set up the script: https://github.com/derekantrican/GAS-ICS-Sync/wiki/Setting-up-the-script-manually

### Install Using Clasp

For the ones who prefer CLI installation, please try the steps below, utilizing [Clasp](https://github.com/google/clasp).

1. Install dependencies: `npm i`.
2. Login: `npm run login`.
3. Enable Apps Script API on the page https://script.google.com/home/usersettings.
4. Create project: `npm run create`.
5. Update `Code.gs` and `filters.gs` with custom configuration.
6. Push local code: `npm run push`.
7. Execute `npm run open` to visit your Apps Script project on UI and run `Code.install` function.
    * If you wish to run this step through CLI command `npx clasp run install`, you need to deploy this script as API executable. Visit https://github.com/google/clasp/blob/master/docs/run.md for more details.

---------------

### Questions? Comments? Anything else?
[Join the Discord!](https://discord.gg/DRBpb4k)

![Discord](https://img.shields.io/discord/612735135120490496)

----------------

### Contributing

If you would like to contribute to this repository, please fork the repository, make your changes, and start a pull request. If your pull request is approved, I will add you as a contributer directly to the repository


**If you would like to fund an issue, you can do that through here: https://issuehunt.io/repos/136078981/**
