# Office365excel Extension for Krista

## Overview

The **Krista Office 365 Excel Extension** enables seamless integration with Microsoft Office 365 Excel, allowing you to automate spreadsheet operations, manage workbooks, sheets, and tables, and perform data manipulation tasks directly from your Krista workflows.

This extension leverages the Microsoft Graph API to provide comprehensive Excel automation capabilities for both shared drives and personal OneDrive files.

## Key Features

- ✅ **Workbook Management**: Access and manage Excel files from OneDrive and shared drives
- ✅ **Sheet Operations**: Create, delete, read, and modify worksheets
- ✅ **Table Operations**: Perform CRUD operations on Excel tables
- ✅ **Data Manipulation**: Add, update, delete rows and columns with flexible input options
- ✅ **Advanced Functions**: Execute formulas, VLOOKUP operations, and range calculations
- ✅ **Batch Operations**: Asynchronous processing for large datasets
- ✅ **File Conversion**: Convert Excel files to PDF format
- ✅ **Flexible Authentication**: Support for both service account and user account authentication
- ✅ **Retry Mechanism**: Built-in retry logic for handling transient failures

## Quick Start Guide

### 1. Prerequisites
- Microsoft Azure account with administrative access
- Office 365 subscription with Excel Online
- Access to create Microsoft Entra ID applications

### 2. Setup Steps
1. [Create a Microsoft Azure Application](pages/CreatingMicrosoftApp.md)
2. [Configure the Extension](pages/ExtensionConfiguration.md)
3. [Set Up Authentication](pages/Authentication.md)

### 3. Start Using
Once configured, you can use any of the [catalog requests](#catalog-requests) in your Krista workflows.



## License

This project is licensed under the **GNU General Public License v3.0**.

```
Office365excel Extension for Krista
Copyright (C) 2025 Krista Software

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
```

For the full license text, see the [LICENSE](LICENSE) file or visit https://www.gnu.org/licenses/gpl-3.0.html.
