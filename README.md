# Quick-AD PowerShell Script

## Overview

Quick-AD is a PowerShell script designed to streamline common Active Directory (AD) management tasks. It provides a graphical user interface (GUI) for efficiently handling user accounts, computer accounts, and other AD-related operations.

## Features

- **User and Computer Management:** Easily find and manage user and computer accounts within the AD environment.
- **CSV Support:** Import user information from CSV files for bulk operations.
- **Password Operations:** Generate, reset, and manage user passwords effortlessly.
- **Group Management:** Copy, remove, and manipulate user group memberships.
- **Organizational Unit (OU) Operations:** Move users or computers to different OUs within the AD structure.

## How to Use

1. **Launching the Script:**
   - Run the main script file (`Quick-AD.ps1`).
   - The script initiates a GUI for intuitive interaction.

2. **Finding AD Users:**
   - Enter the username in the "User Name" field and click "Find User."

3. **Finding AD Computers:**
   - Enter the computer name in the "Computer Name" field and click "Find Computer."

4. **CSV Operations:**
   - Load user information from a CSV file by clicking "Browse" and selecting the file.
   - Perform various operations on the loaded data.

5. **Password Management:**
   - Generate or reset passwords for users.

6. **Group Management:**
   - Copy, remove, or manipulate group memberships.

7. **OU Operations:**
   - Move users or computers to different OUs.

## Requirements

- PowerShell 5.1 or later.

## Notes

- The script logs activities in a file named "Quick-AD.log" in the script directory.
