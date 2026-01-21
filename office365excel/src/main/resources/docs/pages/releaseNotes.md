# Release Notes

## Overview

This document provides version history, bug fixes, new features, and known issues for the Krista Office 365 Excel
Extension.

---

## Version 2.0.28

**Release Date**: October 2025

### Version Information

| Component                  | Version           |
|----------------------------|-------------------|
| Extension Version          | 2.0.28            |
| Developer                  | Vaibhav Choudhary |
| Krista Service APIs (Java) | 1.0.115           |
| Global Catalog Version     | GC-2025.10.3      |
| Microsoft Graph API        | v1.0              |

### What's New

#### **Major Error Message Enhancement Initiative**

This release focuses entirely on improving user experience through comprehensive error message enhancements across the entire extension.

#### **User-Friendly Error Messages**

- **Validator Messages**: Completely rewrote all validation error messages to be more helpful and actionable
- **Authentication Errors**: Enhanced authentication error messages with clear next steps and administrator guidance
- **Service Layer Errors**: Improved error messages in WorkbookService and WorksheetService with better context
- **Consistent Tone**: All error messages now use friendly, helpful language instead of technical jargon

#### **Enhanced Documentation**

- **New Troubleshooting Guide**: Comprehensive troubleshooting documentation with solutions for common issues
- **Error Handling Section**: Added detailed error handling guidance to main README
- **Common Error Codes**: Reference table for quick error resolution
- **Self-Service Support**: Users can now resolve most issues independently

### Detailed Improvements

#### **Validator Error Messages Enhanced**

**File and Resource Validators**:
- **FileIdValidator**: Now explains how to find file IDs and provides permission guidance
- **DriveIdValidator**: Clarifies how to obtain drive IDs using "Get All Shared Drives Information"
- **FileNameValidator**: Lists all supported Excel extensions with descriptions

**Sheet and Table Validators**:
- **SheetNameValidator**: Explains Excel naming rules and suggests using "Get Sheets Information"
- **TableNameValidator**: Emphasizes case-sensitivity and provides clear accessibility guidance
- **ColumnNameValidator**: Highlights case-sensitivity with helpful examples

**Data Format Validators**:
- **CellIdValidator**: Shows proper Excel cell reference formats (A1, B2, AA10)
- **RangeValidator**: Provides multiple valid Excel range format examples
- **RowIndexValidator**: Clarifies Excel row numbering starting from 1
- **FormulaValidator**: Includes formula examples and mentions 8,192 character limit

#### **Authentication Error Messages Improved**

**Enhanced MicrosoftAuthErrorConstants**:
- **Session Expired**: Clear re-authorization guidance
- **Permission Issues**: Specific administrator contact instructions
- **Account Problems**: Detailed explanations for disabled/deleted accounts
- **Application Issues**: Clear guidance for Azure AD configuration problems
- **Service Unavailable**: Helpful retry and escalation guidance

#### **Service Layer Enhancements**

**WorkbookService Improvements**:
- Better workbook not found messages with permission context
- Enhanced batch processing error messages with data format guidance
- Improved CSV processing errors with file format validation tips
- More helpful formula execution errors with syntax guidance

**WorksheetService Improvements**:
- Clearer worksheet deletion error messages
- Enhanced sheet not found errors with troubleshooting steps
- Better permission-related error explanations

#### **Documentation Enhancements**

**New Troubleshooting Guide** (`pages/Troubleshooting.md`):
- **Authentication Issues**: Step-by-step resolution for all auth problems
- **File Access Issues**: Comprehensive file permission and access guidance
- **Data Format Issues**: Clear examples of correct formats for all data types
- **Formula Issues**: Excel formula troubleshooting with examples
- **Performance Issues**: Guidance for handling large datasets and timeouts
- **Error Code Reference**: Quick lookup table for common error codes

**Updated README.md**:
- **Error Handling Section**: New comprehensive error handling guidance
- **Common Error Types**: Categorized error types with solutions
- **Getting Help Section**: Enhanced with troubleshooting guide reference

### User Experience Improvements

#### **Key Benefits**

1. **Faster Problem Resolution**: Users can quickly understand and fix issues
2. **Reduced Support Burden**: Clear guidance reduces administrator intervention needs
3. **Better Learning Experience**: Error messages help users understand Excel integration
4. **Self-Service Capability**: Comprehensive troubleshooting enables independent problem-solving
5. **Consistent Experience**: All error messages follow the same helpful, friendly tone

#### **Error Message Principles Applied**

- **Clear and Concise**: No technical jargon, plain English explanations
- **Actionable Guidance**: Every error includes specific steps to resolve
- **Helpful Context**: Users understand what went wrong and why
- **Examples Provided**: Correct formats and usage examples included
- **Appropriate Tone**: Helpful and supportive, never blaming

### Breaking Changes

None in this version. All existing functionality remains unchanged.

### Deprecations

None in this version.

### Security Updates

No security changes in this version.

### Performance Improvements

No performance changes in this version.

### API Changes

No API changes in this version. All existing catalog requests remain fully compatible.

---

## Version 2.0.26

**Release Date**: October 2025

### Version Information

| Component                  | Version          |
|----------------------------|------------------|
| Extension Version          | 2.0.26           |
| Developer                  | Vrushali Gaikwad |
| Krista Service APIs (Java) | 1.0.115          |
| Global Catalog Version     | GC-2025.10.2     |
| Microsoft Graph API        | v1.0             |

### What's New

#### Documentation Overhaul

- Complete documentation rewrite following comprehensive guidelines
- Individual pages for each catalog request with detailed information
- Enhanced authentication documentation with OAuth 2.0 flows and token management
- Comprehensive configuration guide with troubleshooting
- Organized sidebar navigation by functional categories

#### Enhanced Features

- Improved retry mechanism for handling rate limiting
- Better error messages and validation
- Support for larger datasets in batch operations

### Resolved Bugs

#### KE-2394: Entity Layout Not Honored

- **Issue**: Office 365 Excel does not honor entity layout when writing to spreadsheet
- **Impact**: Data written to Excel may not match expected column order
- **Resolution**: Fixed entity layout processing to maintain correct column order
- **Status**: Resolved

#### KE-2500: Retry and Rate Limiting

- **Issue**: Extension does not properly handle rate limiting from Microsoft Graph API
- **Impact**: Operations fail when rate limits are exceeded
- **Resolution**: Implemented automatic retry with exponential backoff for rate limiting (429 errors)
- **Status**: Resolved

#### KE-2504: Unable to Add Larger Datasets

- **Issue**: Extension fails when attempting to add large datasets to Excel
- **Impact**: Bulk data operations fail or timeout
- **Resolution**: Optimized batch processing and increased timeout limits
- **Status**: Resolved

### Known Issues

#### KE-1226: DateTime Display Issue

- **Issue**: Time is not displayed inside the DateTime range field in "Add Rows With List Of Entity" request
- **Impact**: Time component of DateTime values may not be visible in Excel
- **Workaround**: Format DateTime values as text before adding to Excel
- **Status**: Under Investigation
- **Planned Fix**: Version 2.1.0

### Breaking Changes

None in this version.

### Deprecations

None in this version.

### Security Updates

- Updated OAuth 2.0 token handling for improved security
- Enhanced credential storage encryption
- Improved error messages to avoid exposing sensitive information

### Performance Improvements

- Optimized batch operations for large datasets
- Reduced API call overhead through better caching
- Improved retry logic to minimize delays

### API Changes

No API changes in this version. All existing catalog requests remain compatible.

---

## Version History

### Version 2.0.23

**Release Date**: September 2025

**Changes**:

- Initial support for asynchronous operations
- Added batch processing capabilities
- Bug fixes and stability improvements

### Version 2.0.22

**Release Date**: November 2024

**Changes**:

- Enhanced error handling
- Improved authentication flow
- Performance optimizations

### Version 2.0.21

**Release Date**: October 2024

**Changes**:

- Added support for additional Excel operations
- Bug fixes
- Documentation updates

---

## Upgrade Instructions

### Upgrading from 2.0.23 to 2.0.24

1. **Backup Configuration**:
    - Export current extension configuration
    - Document any custom settings

2. **Update Extension**:
    - Navigate to Extensions in Krista
    - Select Office 365 Excel Extension
    - Click "Update" to version 2.0.24

3. **Verify Configuration**:
    - Review extension configuration
    - Test connection
    - Validate authentication

4. **Test Workflows**:
    - Test existing workflows using the extension
    - Verify all catalog requests function correctly
    - Check error handling and retry behavior

5. **Review Documentation**:
    - Review updated documentation for new features
    - Update any custom documentation or training materials

### No Action Required

- Existing configurations remain compatible
- No changes to catalog request signatures
- Authentication settings are preserved

---

## Compatibility

### Supported Platforms

| Platform            | Version | Status      |
|---------------------|---------|-------------|
| Krista Appliance    | 3.5.1+  | Supported |
| Microsoft Graph API | v1.0    | Supported |
| Office 365 Excel    | Online  | Supported |
| Microsoft Entra ID            | v2.0    | Supported |

### Browser Compatibility (for OAuth)

| Browser | Version | Status      |
|---------|---------|-------------|
| Chrome  | Latest  | Supported |
| Firefox | Latest  | Supported |
| Edge    | Latest  | Supported |
| Safari  | Latest  | Supported |

### Required Permissions

- **Microsoft Graph API**:
    - `Files.ReadWrite` - Read and write user's files
    - `Files.ReadWrite.All` - Read and write all files user can access
    - `openid` - OpenID Connect authentication
    - `offline_access` - Refresh token support

---

## Support

### Getting Help

If you encounter issues with this version:

1. **Review Documentation**:
    - [Extension Configuration](pages/ExtensionConfiguration.md)
    - [Authentication](pages/Authentication.md)
    - [Troubleshooting sections in catalog request pages](../README.md#catalog-requests)

2. **Check Known Issues**:
    - Review known issues section above
    - Check if your issue is already documented

3. **Contact Support**:
    - Provide extension version (2.0.27)
    - Include error messages and logs
    - Describe steps to reproduce the issue

### Reporting Bugs

When reporting bugs, please include:

- Extension version: 2.0.24
- Krista Appliance version
- Catalog request being used
- Complete error message
- Steps to reproduce
- Expected vs actual behavior

---

## Feedback

We value your feedback! Please share:

- Feature requests
- Documentation improvements
- Usability suggestions
- Bug reports

Contact your Krista support representative or submit feedback through the Krista platform.

---

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration guide
- [Authentication](pages/Authentication.md) - OAuth 2.0 authentication and token management
- [Creating Microsoft App](pages/CreatingMicrosoftApp.md) - Microsoft Entra ID application setup
- [Catalog Requests](../README.md#catalog-requests) - Complete list of available operations