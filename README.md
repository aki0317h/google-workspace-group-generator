# Google Workspace Group Auto-Creation Tool

This Google Apps Script automates the creation and configuration of email groups in Google Workspace.

## Features
- Create groups with name, email, and description
- Configure group settings (visibility, permissions, posting rules, etc.)
- Assign a group manager
- Log output to a spreadsheet
- Skip already-processed rows

## How It Works
1. Spreadsheet contains group data (name, email, manager, etc.)
2. The script reads each row and performs:
   - Group creation via Admin SDK
   - Settings configuration via Groups Settings API
   - Manager assignment via Admin SDK
   - Log output to the sheet

## Requirements
- Google Workspace admin privileges
- Enabled Admin SDK and Groups Settings API
- Spreadsheet setup with required columns

## Author
Created by Yuya (ゆうや)
