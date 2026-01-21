# Get Rows For Range (Sheet)

## Overview

Retrieves all rows of data within a specified cell range from an Excel worksheet in OneDrive.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Sheet Name | Text | Yes | Sheet name | "Sheet1" |
| Range | Text | Yes | Cell range (e.g., A1:D10) | "A1:D10" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Rows | List&lt;Composite&gt; | Two-dimensional array of cell values |

## Usage Examples

### Example 1: Get Range Data

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Employees"
Range: "A1:C5"
```

**Output**:
```json
{
  "Rows": [
    ["ID", "Name", "Department"],
    ["EMP001", "John Doe", "Engineering"],
    ["EMP002", "Jane Smith", "Sales"],
    ["EMP003", "Bob Johnson", "Marketing"],
    ["EMP004", "Alice Williams", "HR"]
  ]
}
```

## Business Rules

1. **Range Notation**: Supports Excel range notation (A1:D10)
2. **Empty Cells**: Empty cells returned as empty strings
3. **All Data Types**: Returns all cell values as strings

## Related Catalog Requests

- [Get Rows for Range](pages/GetRowsForRange.md) - Shared drive version
- [Clear Sheet For Range](pages/ClearSheetForRange.md) - Clear range

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /me/drive/items/{fileId}/workbook/worksheets/{sheetName}/range(address='{range}')`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

