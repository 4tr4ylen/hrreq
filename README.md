# HR Requests Tool for SharePoint Online

A comprehensive HR Requests management solution built with SharePoint Framework (SPFx), React, and TypeScript. This solution provides a complete workflow for submitting, managing, and approving HR requests with department-aware security and attachment support.

## Features

### 🎯 Core Functionality
- **Custom Request Form**: User-friendly form with validation and single attachment requirement
- **Department-Aware Security**: Users see only their department's requests with appropriate permissions
- **HR Admin Dashboard**: Full administrative control with bulk approval capabilities
- **Approval Workflow**: Integrated approval process with comments and status tracking
- **File Management**: Secure attachment handling with validation

### 🔒 Security Features
- **Least-Privilege Access**: Users can only edit their own requests until terminal status
- **Department Isolation**: Department members get read-only access to their department's requests
- **HR Admin Override**: HR administrators have full visibility and control
- **Permission Inheritance**: Automatic permission management through Power Automate

### 🎨 User Experience
- **Modern UI**: Built with Fluent UI React for consistent SharePoint experience
- **Responsive Design**: Works across desktop and mobile devices
- **Real-time Validation**: Form validation with helpful error messages
- **Status Tracking**: Visual status indicators with color coding
- **Filtering & Search**: Advanced filtering capabilities for all grids

## Architecture

### Web Parts
1. **RequestForm**: Submission form with validation and file upload
2. **DeptGrid**: Department-specific request viewing with filtering
3. **HrAdminGrid**: Administrative dashboard with approval controls

### Services
- **SharePointService**: Handles all SharePoint operations (CRUD, permissions, attachments)
- **GraphService**: Manages Microsoft Graph operations (user data, departments)

### Models & Interfaces
- **IRequest**: Core request data structure
- **IUser**: User information model
- **IRole**: Permission and role definitions

## Prerequisites

### Development Environment
- Node.js 16.x or later
- npm or yarn package manager
- SharePoint Framework development tools
- Visual Studio Code (recommended)

### SharePoint Environment
- SharePoint Online tenant
- Tenant App Catalog enabled
- Microsoft Graph API access
- Power Automate (for approval workflows)

### Required Permissions
- **SharePoint**: Full Control on target site
- **Microsoft Graph**: User.Read.All (delegated)
- **Power Automate**: Create and manage flows

## Installation & Setup

### 1. Clone and Install Dependencies

```bash
git clone <repository-url>
cd spfx-hr-requests
npm install
```

### 2. Configure Development Environment

```bash
# Install SPFx development tools globally
npm install -g @microsoft/generator-sharepoint

# Trust the development certificate
gulp trust-dev-cert
```

### 3. Build and Package

```bash
# Development build
npm run dev

# Production build and package
npm run package
```

### 4. Deploy to SharePoint

1. **Upload to App Catalog**:
   - Navigate to your tenant app catalog
   - Upload the generated `.sppkg` file
   - Check "Make this solution available to all sites in the organization"
   - Deploy the solution

2. **Approve API Permissions**:
   - Go to SharePoint Admin Center > Advanced > API access
   - Approve the requested Microsoft Graph permissions

3. **Add Web Parts to Pages**:
   - Edit your target pages
   - Add the web parts from the "HR Requests" category

## Site Configuration

### 1. Create SharePoint List

Create a list named "HR Requests" with the following columns:

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

### 2. Configure Site Groups

Create the following site groups:
- **HR Admins**: Full Control
- **Department Members**: Read (or create per-department groups)

### 3. Enable Features

- Enable attachments on the list
- Enable versioning
- Configure item-level permissions

## Power Automate Workflow

### Flow: HR Request Approval

**Trigger**: When an item is created in HR Requests list

**Steps**:
1. **Validate Attachment**: Check for exactly one attachment
2. **Break Inheritance**: Remove inherited permissions
3. **Set Permissions**: 
   - Requestor: Edit access
   - HR Admins: Edit access
   - Department group: Read access
4. **Start Approval**: Initiate approval process

### Required Actions:
- Get attachments
- Break role inheritance
- Add role assignments
- Start approval

## Usage Guide

### For End Users

#### Submitting a Request
1. Navigate to the HR Requests page
2. Fill out the request form:
   - **Title**: Brief description of the request
   - **Request Type**: Select appropriate category
   - **Description**: Detailed explanation
   - **Manager**: Optional manager assignment
   - **Attachment**: Upload required documentation
3. Click "Submit Request"
4. Receive confirmation and tracking information

#### Viewing Requests
1. Use the Department Grid to view requests
2. Filter by status, type, or date range
3. Click on any request to view details
4. Download attachments as needed

### For HR Administrators

#### Managing Requests
1. Access the HR Admin Dashboard
2. View all requests across the organization
3. Use advanced filtering and search
4. Process individual or bulk approvals

#### Approval Process
1. Review request details and attachments
2. Select approve or reject
3. Add comments explaining the decision
4. Submit the decision

#### Bulk Operations
1. Select multiple requests using checkboxes
2. Choose bulk approve or reject
3. Add comments for all selected items
4. Process the batch

## Configuration Options

### Web Part Properties

#### RequestForm
- **Title**: Form header text
- **Description**: Form description
- **Request Types**: Comma-separated list of request types
- **Max File Size**: Maximum attachment size in MB
- **Show Manager Field**: Toggle manager selection
- **Require Manager Approval**: Enable manager approval workflow

#### DeptGrid
- **Title**: Grid header text
- **Description**: Grid description
- **Show Filters**: Enable filtering options
- **Items Per Page**: Number of items to display

#### HrAdminGrid
- **Title**: Dashboard header text
- **Description**: Dashboard description
- **Show Filters**: Enable advanced filtering
- **Items Per Page**: Number of items to display
- **Enable Bulk Actions**: Enable bulk approval features

## Security Considerations

### Permission Model
- **Requestors**: Can edit their own requests until approved/rejected
- **Department Members**: Read-only access to department requests
- **HR Admins**: Full control over all requests
- **Item-Level Security**: Automatic permission management

### Data Protection
- All data stored in SharePoint with native security
- File attachments validated for type and size
- Audit trail maintained through SharePoint versioning
- No sensitive data stored in client-side code

## Troubleshooting

### Common Issues

#### Build Errors
```bash
# Clear cache and rebuild
npm run clean
npm install
npm run build
```

#### Permission Errors
- Verify user is in correct site groups
- Check Microsoft Graph API permissions
- Ensure Power Automate flow is running

#### File Upload Issues
- Check file size limits
- Verify allowed file types
- Ensure SharePoint attachments are enabled

### Debug Mode
```bash
# Enable debug logging
gulp serve --nobrowser --debug
```

## Development

### Project Structure
```
spfx-hr-requests/
├── src/
│   ├── webparts/
│   │   ├── requestForm/
│   │   ├── deptGrid/
│   │   └── hrAdminGrid/
│   ├── services/
│   ├── models/
│   └── utils/
├── config/
├── sharepoint/
└── package.json
```

### Adding New Features
1. Create new web part or component
2. Add to solution package
3. Update documentation
4. Test thoroughly

### Code Standards
- TypeScript strict mode enabled
- ESLint configuration included
- Fluent UI React components
- Async/await for all API calls
- Proper error handling

## Support

### Documentation
- [SharePoint Framework Documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Fluent UI React](https://developer.microsoft.com/en-us/fluentui)
- [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview)

### Issues
- Check the troubleshooting section
- Review SharePoint logs
- Verify configuration settings
- Contact your SharePoint administrator

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## Version History

- **v1.0.0**: Initial release with core functionality
  - Request form with validation
  - Department grid with filtering
  - HR admin dashboard
  - Approval workflow
  - Attachment support

---

**Note**: This solution is designed for SharePoint Online environments. On-premises SharePoint may require additional configuration. 