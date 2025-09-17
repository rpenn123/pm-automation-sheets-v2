# Script Configuration Guide

This document provides a brief overview of the key configuration options within the `src/main.gs` Apps Script for developers or technical users who may need to modify the script's behavior.

## 1. Required Headers

- **Variable**: `CFG.REQUIRED_PIPELINE_HEADERS`, `CFG.REQUIRED_TASK_HEADERS`, `CFG.REQUIRED_OPS_HEADERS`
- **Purpose**: These arrays list the column headers that the script *must* find in the corresponding sheets to function correctly.
- **How it Works**: The script is resilient to column order and the addition of new, unused columns. However, if a column listed in one of these arrays is missing or renamed in the sheet, the script will log an error to the `Ops Inbox` and halt execution for that function to prevent further errors.
- **When to Modify**: If you add a new feature to the script that requires reading from or writing to a new column, you must add that column's header name to the appropriate `REQUIRED` array.

## 2. Zapier-Only Columns

- **Variable**: `CFG.ZAPIER_ONLY_COLUMNS`
- **Purpose**: This array lists columns that are expected to be written to exclusively by Zapier integrations.
- **How it Works**: The script has a guard clause that checks if an edit event occurred in one of these columns. If it did, the script immediately stops processing. This prevents the script's validation logic from running unnecessarily when, for example, Zapier simply updates a notification timestamp.
- **When to Modify**: If you create a new Zapier workflow that writes to a new, dedicated column in the sheet and you do *not* want this write to trigger the script's validation logic, add the column's header name to this array.

## 3. Admin Configuration

- **Mechanism**: Script Properties
- **Purpose**: To securely store the list of users who are authorized to use the "Override: Allow Advance" feature.
- **How it Works**: The list of admin emails is stored as a comma-separated string in a Script Property named `OVERRIDE_ADMINS`. The script uses the `getAdmins_()` and `setAdmins_()` utility functions to interact with this property. This is more secure than hardcoding the list in the script, as only users with edit access to the script project can modify script properties.
- **How to Modify**: The easiest way is to use the built-in UI. In the Google Sheet, go to the **Mobility123 PM → Configure Admins** menu item. This will open a prompt where you can edit the list of admin emails.
