# Authentication

## Overview

The Office 365 Excel Extension uses OAuth 2.0 authentication protocol to securely access Microsoft Graph API and interact with Excel files. This document provides comprehensive information about authentication modes, OAuth flows, token management, and security best practices.

## Authentication Modes

The extension supports two authentication modes to accommodate different use cases:

### Service Account Authentication

Service account authentication uses a single administrator account to execute all operations on behalf of the extension.

**When to Use**:
- Automated workflows that don't require user-specific permissions
- Background processes and scheduled tasks
- Centralized data management scenarios
- Operations on shared organizational resources

**Configuration**:
- Uncheck the **"Have people authenticate with their account"** checkbox in the invoker configuration
- The service account configured during extension setup will be used for all operations

![AskASystem](../_media/Authentication_askASystemUnchecked.png)

**Behavior**:
- All catalog requests execute using the administrator account credentials
- Access is limited to files and resources accessible by the service account
- Single refresh token is maintained for the service account
- No per-user authentication required

**Advantages**:
- Simplified user experience (no authentication prompts)
- Consistent permissions across all operations
- Easier to manage and monitor

**Limitations**:
- All operations use the same account's permissions
- Cannot access user-specific private files
- Audit trails show service account, not individual users

### User Account Authentication

User account authentication requires each user to authenticate with their own Microsoft account.

**When to Use**:
- User-specific file access requirements
- Workflows requiring individual user permissions
- Compliance scenarios requiring user-level audit trails
- Operations on personal OneDrive files

**Configuration**:
- Check the **"Have people authenticate with their account"** checkbox in the invoker configuration
- Each user will be prompted to authenticate when executing the conversation

![AskASystem](../_media/Authentication_askASystemChecked.png)

**Behavior**:
- Users authenticate with their own Microsoft credentials
- Each user's refresh token is stored separately
- Operations execute with the authenticated user's permissions
- Access is limited to files the user can access

**Advantages**:
- User-specific permissions and access control
- Accurate audit trails showing individual users
- Access to personal OneDrive files
- Better security isolation

**Limitations**:
- Users must authenticate before first use
- Each user needs appropriate Office 365 licenses
- More complex token management

## OAuth 2.0 Scopes and Permissions

The extension requires specific Microsoft Graph API permissions to function properly.

### Required Scopes

| Scope | Permission Level | Description |
|-------|------------------|-------------|
| `https://graph.microsoft.com/Files.ReadWrite` | Delegated | Read and write access to user's files |
| `https://graph.microsoft.com/Files.ReadWrite.All` | Delegated | Read and write access to all files user can access |

### Permission Details

#### Files.ReadWrite
- **What it allows**:
  - Read files in the user's OneDrive
  - Create, update, and delete files in the user's OneDrive
  - Access file metadata and content
- **What it doesn't allow**:
  - Access to files shared by others
  - Access to SharePoint site files
  - Access to other users' OneDrive files

#### Files.ReadWrite.All
- **What it allows**:
  - All capabilities of Files.ReadWrite
  - Access to files shared with the user
  - Access to SharePoint site files (if user has permissions)
  - Access to shared drives and team files
- **What it doesn't allow**:
  - Access to files the user doesn't have permissions for
  - Administrative access to all organizational files

### Why Not Use Alternative Scopes?

**Files.Read vs Files.ReadWrite**:
- Files.Read only allows reading files
- The extension requires write capabilities for operations like adding rows, updating cells, etc.
- Files.ReadWrite is the minimum scope needed for full functionality

**Sites.ReadWrite.All**:
- Provides broader SharePoint access
- Not required for Excel-specific operations
- Files.ReadWrite.All is sufficient and more restrictive

## OAuth 2.0 Grant Types

The extension uses two OAuth 2.0 grant types during the authentication lifecycle.

### Authorization Code Grant

Used for initial user authentication and obtaining the first set of tokens.

**Grant Type**: `authorization_code`

**Used For**: Initial authentication when user first authorizes the extension

**Flow**:
1. User is redirected to Microsoft login page
2. User authenticates and grants permissions
3. Microsoft redirects back with authorization code
4. Extension exchanges code for access and refresh tokens

**Request Example**:
```http
POST https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
Content-Type: application/x-www-form-urlencoded

client_id={client_id}
&scope=https://graph.microsoft.com/Files.ReadWrite.All
&code={authorization_code}
&redirect_uri={redirect_uri}
&grant_type=authorization_code
&client_secret={client_secret}
```

**Response Example**:
```json
{
  "token_type": "Bearer",
  "scope": "Files.ReadWrite.All",
  "expires_in": 3600,
  "access_token": "eyJ0eXAiOiJKV1QiLCJhbGc...",
  "refresh_token": "0.AXsA1..."
}
```

### Refresh Token Grant

Used to obtain new access tokens when the current access token expires.

**Grant Type**: `refresh_token`

**Used For**: Obtaining new access tokens without user interaction

**Flow**:
1. Access token expires (after 1 hour)
2. Extension automatically uses refresh token
3. New access token is obtained
4. New refresh token may be issued

**Request Example**:
```http
POST https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
Content-Type: application/x-www-form-urlencoded

client_id={client_id}
&scope=https://graph.microsoft.com/Files.ReadWrite.All
&refresh_token={refresh_token}
&grant_type=refresh_token
&client_secret={client_secret}
```

**Response Example**:
```json
{
  "token_type": "Bearer",
  "scope": "Files.ReadWrite.All",
  "expires_in": 3600,
  "access_token": "eyJ0eXAiOiJKV1QiLCJhbGc...",
  "refresh_token": "0.AXsA1..."
}
```

## Authentication Flow

### Initial Authentication Flow

```
1. User initiates workflow
   ↓
2. Extension checks for valid access token
   ↓
3. No token found → Redirect to Microsoft login
   ↓
4. User authenticates with Microsoft credentials
   ↓
5. User grants requested permissions
   ↓
6. Microsoft redirects to callback URL with authorization code
   ↓
7. Extension exchanges code for access + refresh tokens
   ↓
8. Tokens stored securely in Krista
   ↓
9. Workflow continues with authenticated session
```

### Subsequent Request Flow

```
1. User initiates workflow
   ↓
2. Extension checks for valid access token
   ↓
3. Valid token found → Use token for API request
   ↓
4. If token expired → Use refresh token to get new access token
   ↓
5. Execute workflow with valid token
```

## Token Management

### Access Token

**Lifetime**: 1 hour (3600 seconds)

**Purpose**: Used to authenticate API requests to Microsoft Graph

**Storage**: Temporarily cached in memory during workflow execution

**Expiration Handling**:
- Extension automatically detects expired access tokens
- Refresh token is used to obtain a new access token
- Process is transparent to the user
- No user interaction required

### Refresh Token

**Lifetime**: 
- **Verified apps**: 90 days of inactivity (can be extended with use)
- **Unverified apps**: 24 hours

**Purpose**: Used to obtain new access tokens without user re-authentication

**Storage**: Encrypted and stored securely in Krista database

**Expiration Handling**:
- If refresh token expires, user must re-authenticate
- Extension prompts for re-authentication automatically
- New refresh token is obtained and stored

### Token Refresh Mechanism

The extension implements automatic token refresh:

1. **Before Each Request**: Check if access token is expired or will expire soon
2. **If Expired**: Use refresh token to obtain new access token
3. **Update Storage**: Store new access token (and refresh token if provided)
4. **Retry Request**: Execute the original API request with new token
5. **Transparent**: No user interaction required

**Refresh Timing**:
- Tokens are refreshed proactively before expiration
- Refresh occurs approximately 5 minutes before access token expires
- Prevents mid-request token expiration

## Token Expiration Scenarios

### Access Token Expiration

**When**: After 1 hour of issuance

**Handling**:
1. Extension detects expired access token
2. Automatically uses refresh token to get new access token
3. Retries the failed request with new token
4. User experiences no interruption

**Error Message** (if refresh fails):
```
"Failed to refresh access token. Please re-authenticate."
```

### Refresh Token Expiration

**When**: 
- After 90 days of inactivity (verified apps)
- After 24 hours (unverified apps)
- When explicitly revoked by user or administrator

**Handling**:
1. Extension detects invalid refresh token
2. User is prompted to re-authenticate
3. New authorization flow is initiated
4. New tokens are obtained and stored

**Error Message**:
```
"Your session has expired. Please sign in again."
```

### Token Revocation Scenarios

Refresh tokens can be revoked in several scenarios:

#### User-Initiated Revocation
- User removes app permissions in Microsoft account settings
- User changes Microsoft account password
- User explicitly revokes access

#### Administrative Revocation
- Administrator removes app from Microsoft Entra ID
- Administrator revokes user's permissions
- Conditional access policies block access

#### Security Events
- Suspicious activity detected
- Account compromise suspected
- Security policy violations

#### Token Limit Exceeded
- Microsoft limits the number of refresh tokens per user/app combination
- Oldest tokens are automatically revoked when limit is reached
- Typical limit: 25-50 tokens per user/app

**Handling Revocation**:
1. Extension receives authentication error from Microsoft Graph
2. User is notified that re-authentication is required
3. User is redirected to Microsoft login
4. New tokens are obtained after successful authentication

## Security Best Practices

### For Administrators

1. **Rotate Client Secrets Regularly**
   - Recommended: Every 90 days
   - Set expiration dates when creating secrets
   - Update extension configuration before expiration

2. **Monitor Token Usage**
   - Review authentication logs regularly
   - Investigate unusual authentication patterns
   - Monitor for failed authentication attempts

3. **Limit Permissions**
   - Grant only required scopes (Files.ReadWrite.All)
   - Avoid granting unnecessary admin permissions
   - Review and audit permissions quarterly

4. **Secure Configuration**
   - Store client secrets in Azure Key Vault when possible
   - Restrict access to extension configuration
   - Use strong passwords for service accounts

5. **Enable Conditional Access**
   - Configure Microsoft Entra ID conditional access policies
   - Require multi-factor authentication for sensitive operations
   - Restrict access based on location or device compliance

### For Users

1. **Protect Credentials**
   - Use strong, unique passwords
   - Enable multi-factor authentication
   - Never share account credentials

2. **Review Permissions**
   - Understand what permissions you're granting
   - Review granted permissions periodically
   - Revoke access if no longer needed

3. **Monitor Account Activity**
   - Review Microsoft account activity logs
   - Report suspicious activity immediately
   - Sign out from unused sessions

## Troubleshooting Authentication

### Issue: "Invalid Client" Error

**Cause**: Incorrect Client ID or Client Secret

**Resolution**:
1. Verify Client ID in Azure portal matches configuration
2. Regenerate Client Secret if expired
3. Update extension configuration with new secret
4. Ensure no extra spaces in copied values

### Issue: "Redirect URI Mismatch"

**Cause**: Redirect URI mismatch between Krista and Azure

**Resolution**:
1. Copy exact redirect URI from Krista extension details
2. Add to Microsoft Entra ID app's redirect URIs
3. Ensure protocol, domain, and path match exactly
4. Save changes in Azure portal

### Issue: "Insufficient Permissions"

**Cause**: Required API permissions not granted

**Resolution**:
1. Navigate to Microsoft Entra ID app > API permissions
2. Add Files.ReadWrite and Files.ReadWrite.All
3. Click "Grant admin consent"
4. Wait a few minutes for permissions to propagate

### Issue: "Token Expired" Repeatedly

**Cause**: Refresh token expired or revoked

**Resolution**:
1. Clear stored tokens in Krista
2. Re-authenticate through extension setup
3. Validate configuration
4. Test connection

### Issue: "User Not Found"

**Cause**: User account doesn't exist or lacks Office 365 license

**Resolution**:
1. Verify user has active Office 365 account
2. Ensure user has appropriate license
3. Check user is in correct Microsoft Entra ID tenant
4. Verify user account is not disabled

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration guide
- [Creating Microsoft App](pages/CreatingMicrosoftApp.md) - Microsoft Entra ID application setup
- [Release Notes](pages/releaseNotes.md) - Version information and updates

