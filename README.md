# HR Requests Tool for SharePoint Online

## Developer Handoff Summary (What’s built and how to run it fast)

This repository contains a minimal, production-ready HR Requests tool built with SPFx 1.17 (React + TypeScript), SharePoint Online, Fluent UI, and Microsoft Graph. During this session we:

- Scaffolded three SPFx web parts: RequestForm, DeptGrid, HrAdminGrid
- Implemented shared services: `SharePointService` (REST for list CRUD, single-attachment upload) and `GraphService` (MSGraphClientV3 for user/department)
- Enforced form validation with exactly one attachment (custom file input)
- Built department-aware UX (DeptGrid) and an admin grid (HrAdminGrid) with basic status updates
- Resolved build/runtime environment: Node 16.20.x (SPFx 1.17 requirement), dev cert trust, hosted workbench usage
- Added provisioning scripts to auto-create the SharePoint list and columns (PowerShell PnP + CLI for Microsoft 365)
- Pushed the working repo to GitHub so you can move devices quickly

Quick start for the next developer
- Node: use v16.20.x (SPFx 1.17 requirement)
- Install: `npm install`
- Serve: `gulp serve` (Hosted workbench required on SPFx 1.17+)
- Open hosted workbench (same browser session as tenant):
  - `https://<tenant>.sharepoint.com/sites/<site>/_layouts/15/workbench.aspx?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js`
- List provisioning (choose one):
  - PowerShell (PnP): `./scripts/setup-hr-requests.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<site>" -ListTitle "HR Requests" -Departments @("HR","IT","Finance","Sales","Operations")`
  - Bash (m365 CLI): `./scripts/setup-hr-requests.sh -s "https://<tenant>.sharepoint.com/sites/<site>" -l "HR Requests"`

Common gotchas solved
- Use Node 16.13–16.x; newer Node (18/20) will fail with SPFx 1.17
- SPFx 1.17 removes local workbench by default; use the hosted workbench URL above
- HTTPS dev cert prompts: if blocked, run `gulp untrust-dev-cert && gulp trust-dev-cert`, then accept the cert by opening `https://localhost:4321/temp/manifests.js`
- If mixed content blocked, ensure both hosted workbench and manifests use HTTPS


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

## Implementation Notes (Deep Dive)

### Web parts
- `RequestForm`: Validates required fields; enforces single attachment. On submit: create item via REST then upload binary to `AttachmentFiles/add`. Success dialog shows the list item link.
- `DeptGrid`: Reads via REST, client-side filters by department and status; details dialog; attachment links use server-relative URLs.
- `HrAdminGrid`: Shows all items for HR Admins; inline status approve/reject actions using `SharePointService.updateRequest` (Power Automate flow can enforce item permissioning and final state transitions).

### Services
- `SharePointService`:
  - Create: `/_api/web/lists/getbytitle('HR Requests')/items`
  - Upload: `/_api/web/lists/getbytitle('HR Requests')/items({id})/AttachmentFiles/add(FileName='name.ext')`
  - Get/filter/paging, simple permission hints (owner vs admin)
- `GraphService` (MSGraphClientV3): `/me` for department, `/users` with `$filter=department eq 'X'` for filtered pickers.

### Provisioning
- PowerShell PnP script: `scripts/setup-hr-requests.ps1` (requires `PnP.PowerShell` and PS 7.4+ or pin 2.2.0 on PS 7.1)
- Bash m365 CLI script: `scripts/setup-hr-requests.sh` (no admin required; uses device code login)

### Next Steps / Roadmap
- Department input
  - Switch `Department` to a dropdown (source from Graph or a `Departments` list)
  - Add a web part property to override department options without a rebuild
- Permissions & flow
  - Implement Power Automate “On Create” flow to break inheritance, grant Requestor/HR Admins/Dept Read, and start approval
  - Add “reapply permissions” step on item updates
- UX polish
  - Add status badges in grids, inline edits with confirmation
  - Add paging/infinite scroll to grids for large lists
- Packaging
  - Add pipeline YAML for `bundle --ship` and `package-solution --ship`
  - Add release notes automation

### Testing checklist
- Form blocks submit unless exactly one attachment present
- Item is created and exactly one file is uploaded
- DeptGrid shows the newly created item (based on Department)
- HrAdminGrid shows all items and can update Status for Submitted items

### Project link
- GitHub repo: https://github.com/4tr4ylen/hrreq

## LLM-ready prompts (copy/paste to keep momentum)

- You are a senior SPFx engineer. In `RequestForm`, replace the free-text Department with a Dropdown. Source options from a new web part property `departments` (comma-separated). Parse to `IChoiceGroupOption[]` and persist selected value to the `Department` field.
- You are a senior SPFx + Graph engineer. In `GraphService`, add `getDistinctDepartmentsFromGraph()` (paginated) returning unique department strings. In `RequestForm`, if the property `departments` is empty, call this and render a Dropdown.
- You are a senior SPFx engineer. Add a web part property to `RequestFormWebPart` named `listTitle` (default `HR Requests`). Update `SharePointService` to accept `listTitle` from props and use it for all REST calls.
- You are a senior SPFx engineer. Implement optimistic UI updates in `HrAdminGrid` when changing Status (Submitted → Approved/Rejected). Disable actions while update is in-flight and show a success/error `MessageBar`.
- You are a senior SharePoint + Power Automate engineer. Create a flow that triggers on item created in list `HR Requests`, validates exactly one attachment, breaks permissions, grants Edit to Requestor and HR Admins, Read to Department group, and starts approval. Document endpoints used in the README.
- You are a senior SPFx engineer. Add paging to `DeptGrid` using `$top` + `$skiptoken` or `$skip` on REST. Surface “Next/Previous” in the grid footer.
- You are a senior SPFx engineer. Add an export to CSV function in `HrAdminGrid` that serializes current filtered items client-side and triggers a download.
- You are a senior SPFx security engineer. Centralize list/field internal names in `src/models/constants.ts`, remove magic strings throughout services and components.
- You are a senior SPFx engineer. Add CI to package the solution: GitHub Actions workflow that runs `npm ci`, `gulp bundle --ship`, and `gulp package-solution --ship`, then uploads the `.sppkg` as an artifact.
- You are a senior SPFx engineer. Add feature flags (web part properties) to toggle: showManagerField, requireManagerApproval, and enforceDepartmentLock (prevents editing Department by requestor).