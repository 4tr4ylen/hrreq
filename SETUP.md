# HR Requests Tool - Setup Guide

## Node.js Version Requirements

This SPFx solution requires Node.js version 16.x for compatibility with SPFx 1.17.4.

### Install Node.js 16.x

**Option 1: Using Node Version Manager (nvm) - Recommended**

```bash
# Install nvm if you don't have it
curl -o- https://raw.githubusercontent.com/nvm-sh/nvm/v0.39.0/install.sh | bash

# Restart your terminal or run
source ~/.bashrc

# Install and use Node.js 16
nvm install 16
nvm use 16

# Verify version
node --version  # Should show v16.x.x
```

**Option 2: Direct Download**

Download Node.js 16.x from [nodejs.org](https://nodejs.org/dist/latest-v16.x/)

**Option 3: Using Homebrew (macOS)**

```bash
# Install Node.js 16
brew install node@16

# Link it
brew link node@16 --force

# Verify version
node --version
```

## Project Setup

1. **Switch to Node.js 16**:
   ```bash
   nvm use 16  # or use your preferred method
   ```

2. **Install Dependencies**:
   ```bash
   npm install
   ```

3. **Trust Development Certificate**:
   ```bash
   gulp trust-dev-cert
   ```

4. **Build the Project**:
   ```bash
   npm run build
   ```

5. **Package for Production**:
   ```bash
   npm run package
   ```

## Development Commands

```bash
# Development server
npm run dev

# Build
npm run build

# Package for deployment
npm run package

# Clean build artifacts
npm run clean
```

## Deployment

1. **Upload to App Catalog**:
   - Navigate to your tenant app catalog
   - Upload the generated `.sppkg` file from `sharepoint/solution/`
   - Check "Make this solution available to all sites in the organization"
   - Deploy the solution

2. **Approve API Permissions**:
   - Go to SharePoint Admin Center > Advanced > API access
   - Approve the requested Microsoft Graph permissions:
     - User.Read.All (delegated)
     - User.ReadBasic.All (delegated)

3. **Add Web Parts to Pages**:
   - Edit your target pages
   - Add the web parts from the "HR Requests" category:
     - HR Request Form
     - Department Requests Grid
     - HR Admin Requests Grid

## SharePoint List Configuration

Create a list named "HR Requests" with these columns:

| Column Name | Internal Name | Type | Required |
|-------------|---------------|------|----------|
| Title | Title | Single line of text | Yes |
| Request Type | RequestType | Choice | Yes |
| Description | Description | Multiple lines of text | Yes |
| Department | Department | Choice | Yes |
| Requestor | Requestor | Person | Yes |
| Manager | Manager | Person | No |
| Status | Status | Choice | Yes |
| Approval Outcome | ApprovalOutcome | Choice | No |
| Approver Comments | ApproverComments | Multiple lines of text | No |

**Choice Values**:
- **RequestType**: Leave Request, Equipment Request, Policy Question, Benefits Question, Other
- **Status**: Draft, Submitted, Pending Approval, Approved, Rejected, Completed
- **ApprovalOutcome**: Approved, Rejected

## Site Groups

Create these site groups:
- **HR Admins**: Full Control
- **Department Members**: Read (or create per-department groups)

## Power Automate Flow

Create a flow triggered when items are created in the HR Requests list:

1. **Validate Attachment**: Check for exactly one attachment
2. **Break Inheritance**: Remove inherited permissions
3. **Set Permissions**: 
   - Requestor: Edit access
   - HR Admins: Edit access
   - Department group: Read access
4. **Start Approval**: Initiate approval process

## Troubleshooting

### Node.js Version Issues
```bash
# Check current version
node --version

# Switch to Node.js 16
nvm use 16

# Clear npm cache if needed
npm cache clean --force
```

### Build Errors
```bash
# Clean and rebuild
npm run clean
npm install
npm run build
```

### Permission Errors
- Verify user is in correct site groups
- Check Microsoft Graph API permissions
- Ensure Power Automate flow is running

## Project Structure

```
spfx-hr-requests/
├── src/
│   ├── webparts/
│   │   ├── requestForm/          # Request submission form
│   │   ├── deptGrid/             # Department-specific grid
│   │   └── hrAdminGrid/          # HR admin dashboard
│   ├── services/
│   │   ├── SharePointService.ts  # SharePoint operations
│   │   └── GraphService.ts       # Microsoft Graph operations
│   ├── models/
│   │   ├── IRequest.ts           # Request interfaces
│   │   └── IRole.ts              # Permission interfaces
│   └── utils/
│       ├── validation.ts         # Form validation
│       └── attachments.ts        # File handling
├── config/                       # Build configuration
├── sharepoint/                   # Solution package
└── README.md                     # Full documentation
```

## Features Implemented

✅ **RequestForm Web Part**
- Form validation with required fields
- Single file attachment requirement
- Manager selection (optional)
- Success dialog with item link

✅ **DeptGrid Web Part**
- Department-aware filtering
- Advanced filtering options
- Request details dialog
- Attachment download links

✅ **HrAdminGrid Web Part**
- Full administrative control
- Bulk approval capabilities
- Individual approval workflow
- Export functionality (placeholder)

✅ **Services**
- SharePointService for CRUD operations
- GraphService for user/department data
- Permission management
- File upload handling

✅ **Security**
- Department-based access control
- Item-level permissions
- HR admin override capabilities
- Least-privilege access model

## Next Steps

1. **Deploy to SharePoint Online**
2. **Configure Power Automate workflow**
3. **Test with real users**
4. **Customize request types and workflows**
5. **Add additional features as needed**

## Support

For issues or questions:
1. Check the troubleshooting section
2. Review SharePoint logs
3. Verify configuration settings
4. Contact your SharePoint administrator

---

**Note**: This solution is designed for SharePoint Online environments. On-premises SharePoint may require additional configuration. 