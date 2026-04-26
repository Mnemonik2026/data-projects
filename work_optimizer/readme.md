# Project Name: AZS Control Panel (Automated Fuel Inventory System)

## Description:
I developed a standalone desktop application designed to streamline and automate daily fuel stock management and reporting for a fuel and lubricants storage facility. The application eliminates manual data entry errors and accelerates the daily reporting cycle.

## Key Features:

Data Aggregation: Automatically parses and consolidates daily fuel balances from multiple Excel files.

Tank Calibration: Includes a custom module with a GUI to generate calibration tables for horizontal fuel tanks based on specified physical dimensions.

Automated Notifications: Integrates with Telegram API and SMTP servers (supporting various providers like Gmail and Ukr.net) to automatically distribute consolidated reports to stakeholders.

Dynamic Configuration: Features a user-friendly interface to manage API tokens, email credentials, and recipient lists without altering the underlying code.

## Tools & Technologies Used:

Python: Core logic and automation.

CustomTkinter: Building a modern, dark-mode graphical user interface.

Pandas & Openpyxl: Data extraction and manipulation in Excel.

Telegram Bot API & smtplib: Setting up automated messaging and email distribution.

## Impact:
By fully automating the routine data collection, calculation, and reporting processes, this application saves me approximately one hour of manual work every single day.
