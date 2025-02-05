# Door Dec Creator

## Overview
This Python script generates personalized door decorations for residents using images and names. It creates a PDF file with dynamically cropped images, white nameplates, and formatted text.

## Features
- Reads names from an Excel file or user input.
- Dynamically crops images to maintain a consistent size.
- Adds a white box for name placement.
- Places images in a grid layout on a PDF page.
- Supports both **Roster Mode** (predefined list from Excel) and **Custom Mode** (manual input).

## Dependencies
Ensure you have the following installed:
```sh
pip install pillow reportlab pandas openpyxl
