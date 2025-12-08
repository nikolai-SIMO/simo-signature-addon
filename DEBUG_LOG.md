# SIMO Signatures Add-in - Debug Log

## Issue
Profile photo not displaying in email signatures

## Status
IN PROGRESS - Testing REST API fallback

## Root Cause
- SSO token obtained successfully
- Graph API returns 401 (Unauthorized) for `/me/photo/$value`
- The Office SSO token doesn't have `User.Read` scope for Microsoft Graph photo access

## Fixes Applied

### Commit 1: 04e43eb
- Changed from embedding base64 directly to using CID inline attachments
- Added `addFileAttachmentFromBase64Async` with `isInline: true`
- HTML now uses `<img src="cid:profile-photo.png">` instead of data URL
- Added custom photo URL option in settings as fallback
- Added `inlineAttachment` API detection in debug panel

### Commit 2: 460e44c
- Added "Copy Full Log" button in debug panel for easier troubleshooting
- Shows photo size when loaded
- Displays last 10 log entries

### Commit 3: 5975a08
- Fixed fallback: now tries REST API when Graph API fails (401/error)
- Previously it only fell back when SSO token acquisition failed

## Last Debug Output (2024-12-08 21:23)
```
API:
  Version: 1.10
  setSelectedData: ✅
  setSignature: ✅
  inlineAttachment: ✅

Profile:
  Photo: ❌ Not loaded

Log:
  [21:23:48] Fetching photo via SSO...
  [21:23:48] SSO token obtained
  [21:23:48] Graph photo failed: 401
  [21:23:48] Apply: default
  [21:23:49] Success: setSelectedData
```

## Next Steps
1. Refresh add-in and test again
2. Check if REST API fallback works (should see "trying REST API..." in log)
3. If REST API also fails, possible solutions:
   - **Option A**: Register app in Azure AD with proper Graph permissions
   - **Option B**: Use EWS (Exchange Web Services) to get photo
   - **Option C**: Host employee photos on GitHub and use custom URL setting

## How to Test
1. Refresh/restart Outlook
2. Open the add-in
3. Go to Settings > Debug section
4. Click "📋 Copy Full Log" button
5. Check if photo loaded or what error occurred

## Files Modified
- `simo-signatures.html` - Main add-in code

## Relevant Code Locations
- Photo fetch: `fetchProfilePhoto()` ~line 434
- Graph API call: `fetchPhotoFromGraph()` ~line 461
- REST API fallback: `tryRestApiPhoto()` ~line 494
- Signature generation: `getSignatureHTML()` ~line 607
- Inline attachment: `tryInsert()` ~line 757
