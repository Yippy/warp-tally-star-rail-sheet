# Honkai Star Rail Warp Tracker
Script to help manage Honkai Star Rail warp history using Google Sheet Document.

## Project Website
Visit the Genshin Impact collection of Google Sheets tools:

https://gensheets.co.uk 

## Google Add-on
Not available yet

## Preview
<img src="https://raw.github.com/Yippy/warp-tally-star-rail-sheet/master/images/warp-tally-preview.png?sanitize=true">

## Tutorial
[Install Add-on](docs/INSTALL_ADD_ON.md)

[Setup Add-on](docs/SETUP_ADD_ON.md)

[Change Language](docs/CHANGE_LANGUAGE.md)

[Get README](docs/GET_README.md)

[Use Auto Import](docs/USE_AUTO_IMPORT.md)

## Template Document
If you prefer to use the Wish Tally document with embedded script, you can make a copy here:
https://docs.google.com/spreadsheets/d/1o3stZ7bCO2Wzz7d80VRKIQvr8GxCjhLhbsPAT3yaaIU/edit

## How to compile script
This project uses https://github.com/google/clasp to help compile code to Google Script.

1. Run ```npm install -g @google/clasp```
2. Edit the file .clasp.json with your Google Script
3. Run ```clasp login```
4. Run ```clasp push -w```