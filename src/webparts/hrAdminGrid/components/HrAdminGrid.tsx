import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  // SelectionMode,
  IColumn,
  Stack,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  IconButton,
  Link,
  Dialog,
  DialogType,
  DialogFooter,
  Persona,
  PersonaSize,
  Label,
  // Toggle,
  CommandBar,
  ICommandBarItemProps,
  // duplicate removal handled
  Checkbox,
  Selection,
  ISelection
} from '@fluentui/react';
import { IHrAdminGridWebPartProps } from '../IHrAdminGridWebPartProps';
import { SharePointService } from '../../../services/SharePointService';
// import { GraphService } from '../../../services/GraphService';
import { IRequest, RequestStatus, IRequestFilters, ApprovalOutcome } from '../../../models/IRequest';
import { AttachmentUtils } from '../../../utils/attachments';

export interface IHrAdminGridState {
  requests: IRequest[];
  filteredRequests: IRequest[];
  isLoading: boolean;
  currentUser: any;
  departments: string[];
  requestTypes: string[];
  filters: IRequestFilters;
  selectedRequest: IRequest | null;
  showDetailsDialog: boolean;
  showApprovalDialog: boolean;
  showBulkApprovalDialog: boolean;
  userPermissions: any | null;
  currentPage: number;
  itemsPerPage: number;
  selection: ISelection;
  selectedItems: IRequest[];
  approvalComments: string;
  approvalOutcome: ApprovalOutcome;
  isProcessing: boolean;
}

export const HrAdminGrid: React.FC<IHrAdminGridWebPartProps> = (props) => {
  const [state, setState] = useState<IHrAdminGridState>({
    requests: [],
    filteredRequests: [],
    isLoading: true,
    currentUser: null,
    departments: [],
    requestTypes: [],
    filters: {},
    selectedRequest: null,
    showDetailsDialog: false,
    showApprovalDialog: false,
    showBulkApprovalDialog: false,
    userPermissions: null,
    currentPage: 1,
    itemsPerPage: parseInt(props.itemsPerPage) || 50,
    selection: new Selection({
      onSelectionChanged: () => {
        const selectedItems = state.selection.getSelection() as IRequest[];
        setState(prev => ({ ...prev, selectedItems }));
      }
    }),
    selectedItems: [],
    approvalComments: '',
    approvalOutcome: ApprovalOutcome.Approved,
    isProcessing: false
  });

  const sharePointService = new SharePointService(props.context);
  // const graphService = new GraphService(props.context);

  useEffect(() => {
    loadData();
  }, []);

  useEffect(() => {
    applyFilters();
  }, [state.requests, state.filters]);

  const loadData = async (): Promise<void> => {
    try {
      setState(prev => ({ ...prev, isLoading: true }));

      // Check if user is HR admin
      const isHRAdmin = await sharePointService.isHRAdmin();
      if (!isHRAdmin) {
        setState(prev => ({
          ...prev,
          isLoading: false,
          requests: [],
          filteredRequests: []
        }));
        return;
      }

      // Get current user
      const currentUser = await sharePointService.getCurrentUser();
      
      // Get all requests (HR admins see everything)
      const requests = await sharePointService.getRequests();
      
      // Extract unique departments and request types
      const departments = [...new Set(requests.map(r => r.Department))].sort();
      const requestTypes = [...new Set(requests.map(r => r.RequestType))].sort();

      setState(prev => ({
        ...prev,
        requests,
        currentUser,
        departments,
        requestTypes,
        isLoading: false
      }));
    } catch (error) {
      console.error('Error loading data:', error);
      setState(prev => ({ ...prev, isLoading: false }));
    }
  };

  const applyFilters = (): void => {
    let filtered = [...state.requests];

    if (state.filters.status && state.filters.status.length > 0) {
      filtered = filtered.filter(r => state.filters.status!.includes(r.Status));
    }

    if (state.filters.department) {
      filtered = filtered.filter(r => r.Department === state.filters.department);
    }

    if (state.filters.requestType) {
      filtered = filtered.filter(r => r.RequestType === state.filters.requestType);
    }

    if (state.filters.dateFrom) {
      filtered = filtered.filter(r => new Date(r.Created) >= state.filters.dateFrom!);
    }

    if (state.filters.dateTo) {
      filtered = filtered.filter(r => new Date(r.Created) <= state.filters.dateTo!);
    }

    setState(prev => ({ ...prev, filteredRequests: filtered }));
  };

  const handleFilterChange = (field: keyof IRequestFilters, value: any): void => {
    setState(prev => ({
      ...prev,
      filters: {
        ...prev.filters,
        [field]: value
      }
    }));
  };

  const handleRequestClick = async (request: IRequest): Promise<void> => {
    try {
      const permissions = await sharePointService.getUserPermissions(request);
      setState(prev => ({
        ...prev,
        selectedRequest: request,
        userPermissions: permissions,
        showDetailsDialog: true
      }));
    } catch (error) {
      console.error('Error getting user permissions:', error);
    }
  };

  const closeDetailsDialog = (): void => {
    setState(prev => ({
      ...prev,
      showDetailsDialog: false,
      selectedRequest: null,
      userPermissions: null
    }));
  };

  const openApprovalDialog = (request: IRequest): void => {
    setState(prev => ({
      ...prev,
      selectedRequest: request,
      showApprovalDialog: true,
      approvalComments: '',
      approvalOutcome: ApprovalOutcome.Approved
    }));
  };

  const closeApprovalDialog = (): void => {
    setState(prev => ({
      ...prev,
      showApprovalDialog: false,
      selectedRequest: null,
      approvalComments: ''
    }));
  };

  const handleApproval = async (): Promise<void> => {
    if (!state.selectedRequest) return;

    try {
      setState(prev => ({ ...prev, isProcessing: true }));

      const newStatus = state.approvalOutcome === ApprovalOutcome.Approved 
        ? RequestStatus.Approved 
        : RequestStatus.Rejected;

      await sharePointService.updateRequest(state.selectedRequest.Id!, {
        Status: newStatus,
        ApprovalOutcome: state.approvalOutcome,
        ApproverComments: state.approvalComments
      });

      // Refresh data
      await loadData();

      setState(prev => ({
        ...prev,
        showApprovalDialog: false,
        selectedRequest: null,
        approvalComments: '',
        isProcessing: false
      }));

    } catch (error) {
      console.error('Error updating request:', error);
      setState(prev => ({ ...prev, isProcessing: false }));
    }
  };

  const openBulkApprovalDialog = (): void => {
    setState(prev => ({
      ...prev,
      showBulkApprovalDialog: true,
      approvalComments: '',
      approvalOutcome: ApprovalOutcome.Approved
    }));
  };

  const closeBulkApprovalDialog = (): void => {
    setState(prev => ({
      ...prev,
      showBulkApprovalDialog: false,
      approvalComments: ''
    }));
  };

  const handleBulkApproval = async (): Promise<void> => {
    try {
      setState(prev => ({ ...prev, isProcessing: true }));

      const newStatus = state.approvalOutcome === ApprovalOutcome.Approved 
        ? RequestStatus.Approved 
        : RequestStatus.Rejected;

      // Process all selected items
      for (const request of state.selectedItems) {
        await sharePointService.updateRequest(request.Id!, {
          Status: newStatus,
          ApprovalOutcome: state.approvalOutcome,
          ApproverComments: state.approvalComments
        });
      }

      // Refresh data
      await loadData();

      setState(prev => ({
        ...prev,
        showBulkApprovalDialog: false,
        approvalComments: '',
        selectedItems: [],
        isProcessing: false
      }));

    } catch (error) {
      console.error('Error processing bulk approval:', error);
      setState(prev => ({ ...prev, isProcessing: false }));
    }
  };

  const getRequestItemUrl = (request: IRequest): string => {
    return `${props.context.pageContext.web.absoluteUrl}/Lists/HR%20Requests/DispForm.aspx?ID=${request.Id}`;
  };

  const getStatusColor = (status: RequestStatus): string => {
    switch (status) {
      case RequestStatus.Draft:
        return '#605e5c';
      case RequestStatus.Submitted:
        return '#0078d4';
      case RequestStatus.PendingApproval:
        return '#ff8c00';
      case RequestStatus.Approved:
        return '#107c10';
      case RequestStatus.Rejected:
        return '#d13438';
      case RequestStatus.Completed:
        return '#881798';
      default:
        return '#605e5c';
    }
  };

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: loadData
    },
    {
      key: 'export',
      text: 'Export',
      iconProps: { iconName: 'Download' },
      onClick: () => {
        // TODO: Implement export functionality
        console.log('Export functionality to be implemented');
      }
    }
  ];

  const commandBarFarItems: ICommandBarItemProps[] = [
    {
      key: 'bulkApproval',
      text: `Bulk Approve (${state.selectedItems.length})`,
      iconProps: { iconName: 'CheckMark' },
      disabled: state.selectedItems.length === 0,
      onClick: openBulkApprovalDialog
    }
  ];

  const columns: IColumn[] = [
    {
      key: 'selection',
      name: '',
      minWidth: 30,
      maxWidth: 30,
      isResizable: false,
      onRender: (item: IRequest) => (
        <Checkbox
          checked={state.selection.isKeySelected(item.Id!.toString())}
          onChange={(_, checked) => {
            if (checked) {
              state.selection.setKeySelected(item.Id!.toString(), true, true);
            } else {
              state.selection.setKeySelected(item.Id!.toString(), false, true);
            }
          }}
        />
      )
    },
    {
      key: 'title',
      name: 'Title',
      fieldName: 'Title',
      minWidth: 200,
      maxWidth: 300,
      isResizable: true,
      onRender: (item: IRequest) => (
        <Link onClick={() => handleRequestClick(item)}>
          {item.Title}
        </Link>
      )
    },
    {
      key: 'requestType',
      name: 'Request Type',
      fieldName: 'RequestType',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true
    },
    {
      key: 'status',
      name: 'Status',
      fieldName: 'Status',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IRequest) => (
        <span style={{ color: getStatusColor(item.Status) }}>
          {item.Status}
        </span>
      )
    },
    {
      key: 'department',
      name: 'Department',
      fieldName: 'Department',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true
    },
    {
      key: 'author',
      name: 'Requestor',
      fieldName: 'Author',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: IRequest) => (
        <Persona
          text={item.Author.DisplayName}
          secondaryText={item.Author.Email}
          size={PersonaSize.size24}
        />
      )
    },
    {
      key: 'created',
      name: 'Created',
      fieldName: 'Created',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: IRequest) => new Date(item.Created).toLocaleDateString()
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: IRequest) => (
        <Stack horizontal tokens={{ childrenGap: 5 }}>
          <IconButton
            iconProps={{ iconName: 'View' }}
            title="View Details"
            onClick={() => handleRequestClick(item)}
          />
          {item.Status === RequestStatus.Submitted && (
            <IconButton
              iconProps={{ iconName: 'CheckMark' }}
              title="Approve/Reject"
              onClick={() => openApprovalDialog(item)}
            />
          )}
        </Stack>
      )
    }
  ];

  const statusOptions: IDropdownOption[] = Object.values(RequestStatus).map(status => ({
    key: status,
    text: status
  }));

  const departmentOptions: IDropdownOption[] = state.departments.map(dept => ({
    key: dept,
    text: dept
  }));

  const requestTypeOptions: IDropdownOption[] = state.requestTypes.map(type => ({
    key: type,
    text: type
  }));

  if (state.isLoading) {
    return (
      <Stack horizontalAlign="center" verticalAlign="center" style={{ minHeight: '200px' }}>
        <Spinner size={SpinnerSize.large} label="Loading requests..." />
      </Stack>
    );
  }

  return (
    <div>
      <Stack tokens={{ childrenGap: 15 }}>
        {/* Header */}
        <Stack>
          <h2>{props.title}</h2>
          {props.description && <p>{props.description}</p>}
        </Stack>

        {/* Command Bar */}
        <CommandBar
          items={commandBarItems}
          farItems={commandBarFarItems}
        />

        {/* Filters */}
        {props.showFilters && (
          <Stack horizontal tokens={{ childrenGap: 10 }} wrap>
            <Dropdown
              label="Status"
              placeholder="Select status"
              options={statusOptions}
              selectedKey={state.filters.status?.[0]}
              onChange={(_, option) => handleFilterChange('status', option ? [option.key as RequestStatus] : undefined)}
              style={{ minWidth: 150 }}
            />
            <Dropdown
              label="Department"
              placeholder="Select department"
              options={departmentOptions}
              selectedKey={state.filters.department}
              onChange={(_, option) => handleFilterChange('department', option?.key as string)}
              style={{ minWidth: 150 }}
            />
            <Dropdown
              label="Request Type"
              placeholder="Select request type"
              options={requestTypeOptions}
              selectedKey={state.filters.requestType}
              onChange={(_, option) => handleFilterChange('requestType', option?.key as string)}
              style={{ minWidth: 150 }}
            />
            <TextField
              label="Date From"
              type="date"
              value={state.filters.dateFrom?.toISOString().split('T')[0] || ''}
              onChange={(_, newValue) => handleFilterChange('dateFrom', newValue ? new Date(newValue) : undefined)}
              style={{ minWidth: 150 }}
            />
            <TextField
              label="Date To"
              type="date"
              value={state.filters.dateTo?.toISOString().split('T')[0] || ''}
              onChange={(_, newValue) => handleFilterChange('dateTo', newValue ? new Date(newValue) : undefined)}
              style={{ minWidth: 150 }}
            />
            <DefaultButton
              text="Clear Filters"
              onClick={() => setState(prev => ({ ...prev, filters: {} }))}
              style={{ alignSelf: 'end' }}
            />
          </Stack>
        )}

        {/* Results Count */}
        <div>
          <Label>
            Showing {state.filteredRequests.length} of {state.requests.length} requests
            {state.selectedItems.length > 0 && ` (${state.selectedItems.length} selected)`}
          </Label>
        </div>

        {/* Requests Grid */}
        <DetailsList
          items={state.filteredRequests}
          columns={columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selection={state.selection}
          isHeaderVisible={true}
          compact={true}
        />

        {/* No Results */}
        {state.filteredRequests.length === 0 && !state.isLoading && (
          <MessageBar messageBarType={MessageBarType.info}>
            No requests found matching your criteria.
          </MessageBar>
        )}
      </Stack>

      {/* Request Details Dialog */}
      <Dialog
        hidden={!state.showDetailsDialog}
        onDismiss={closeDetailsDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: state.selectedRequest?.Title || 'Request Details'
        }}
      >
        {state.selectedRequest && (
          <Stack tokens={{ childrenGap: 10 }}>
            <Stack horizontal tokens={{ childrenGap: 20 }}>
              <Stack.Item grow>
                <Label>Request Type</Label>
                <div>{state.selectedRequest.RequestType}</div>
              </Stack.Item>
              <Stack.Item grow>
                <Label>Status</Label>
                <div style={{ color: getStatusColor(state.selectedRequest.Status) }}>
                  {state.selectedRequest.Status}
                </div>
              </Stack.Item>
            </Stack>

            <div>
              <Label>Description</Label>
              <div style={{ whiteSpace: 'pre-wrap' }}>{state.selectedRequest.Description}</div>
            </div>

            <Stack horizontal tokens={{ childrenGap: 20 }}>
              <Stack.Item grow>
                <Label>Department</Label>
                <div>{state.selectedRequest.Department}</div>
              </Stack.Item>
              <Stack.Item grow>
                <Label>Created</Label>
                <div>{new Date(state.selectedRequest.Created).toLocaleString()}</div>
              </Stack.Item>
            </Stack>

            <div>
              <Label>Requestor</Label>
              <Persona
                text={state.selectedRequest.Author.DisplayName}
                secondaryText={state.selectedRequest.Author.Email}
                size={PersonaSize.size32}
              />
            </div>

            {state.selectedRequest.Manager && (
              <div>
                <Label>Manager</Label>
                <Persona
                  text={state.selectedRequest.Manager.DisplayName}
                  secondaryText={state.selectedRequest.Manager.Email}
                  size={PersonaSize.size32}
                />
              </div>
            )}

            {state.selectedRequest.Attachments && state.selectedRequest.Attachments.length > 0 && (
              <div>
                <Label>Attachments</Label>
                {state.selectedRequest.Attachments.map((attachment, index) => (
                  <div key={index} style={{ marginTop: '5px' }}>
                    <Link
                      href={AttachmentUtils.createDownloadLink(attachment.ServerRelativeUrl)}
                      target="_blank"
                    >
                      {attachment.FileName}
                    </Link>
                    <span style={{ marginLeft: '10px', color: '#666' }}>
                      ({AttachmentUtils.formatFileSize(attachment.Length)})
                    </span>
                  </div>
                ))}
              </div>
            )}

            {state.selectedRequest.ApproverComments && (
              <div>
                <Label>Approver Comments</Label>
                <div style={{ whiteSpace: 'pre-wrap' }}>{state.selectedRequest.ApproverComments}</div>
              </div>
            )}
          </Stack>
        )}

        <DialogFooter>
          <DefaultButton onClick={closeDetailsDialog} text="Close" />
          {state.selectedRequest && (
            <>
              <DefaultButton
                onClick={() => window.open(getRequestItemUrl(state.selectedRequest!), '_blank')}
                text="View Full Details"
                iconProps={{ iconName: 'OpenInNewWindow' }}
              />
              {state.selectedRequest.Status === RequestStatus.Submitted && (
                <PrimaryButton
                  onClick={() => {
                    closeDetailsDialog();
                    openApprovalDialog(state.selectedRequest!);
                  }}
                  text="Approve/Reject"
                  iconProps={{ iconName: 'CheckMark' }}
                />
              )}
            </>
          )}
        </DialogFooter>
      </Dialog>

      {/* Approval Dialog */}
      <Dialog
        hidden={!state.showApprovalDialog}
        onDismiss={closeApprovalDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Approve/Reject Request'
        }}
      >
        <Stack tokens={{ childrenGap: 15 }}>
          <div>
            <Label>Request</Label>
            <div>{state.selectedRequest?.Title}</div>
          </div>

          <Dropdown
            label="Decision"
            options={[
              { key: ApprovalOutcome.Approved, text: 'Approve' },
              { key: ApprovalOutcome.Rejected, text: 'Reject' }
            ]}
            selectedKey={state.approvalOutcome}
            onChange={(_, option) => setState(prev => ({ ...prev, approvalOutcome: option?.key as ApprovalOutcome }))}
          />

          <TextField
            label="Comments"
            value={state.approvalComments}
            onChange={(_, newValue) => setState(prev => ({ ...prev, approvalComments: newValue || '' }))}
            multiline
            rows={4}
            placeholder="Enter approval comments..."
          />
        </Stack>

        <DialogFooter>
          <DefaultButton onClick={closeApprovalDialog} text="Cancel" disabled={state.isProcessing} />
          <PrimaryButton
            onClick={handleApproval}
            text={state.approvalOutcome === ApprovalOutcome.Approved ? 'Approve' : 'Reject'}
            disabled={state.isProcessing}
            iconProps={{ iconName: 'CheckMark' }}
          />
          {state.isProcessing && <Spinner size={SpinnerSize.small} />}
        </DialogFooter>
      </Dialog>

      {/* Bulk Approval Dialog */}
      <Dialog
        hidden={!state.showBulkApprovalDialog}
        onDismiss={closeBulkApprovalDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: `Bulk Approve/Reject (${state.selectedItems.length} items)`
        }}
      >
        <Stack tokens={{ childrenGap: 15 }}>
          <MessageBar messageBarType={MessageBarType.warning}>
            You are about to process {state.selectedItems.length} requests. This action cannot be undone.
          </MessageBar>

          <Dropdown
            label="Decision"
            options={[
              { key: ApprovalOutcome.Approved, text: 'Approve All' },
              { key: ApprovalOutcome.Rejected, text: 'Reject All' }
            ]}
            selectedKey={state.approvalOutcome}
            onChange={(_, option) => setState(prev => ({ ...prev, approvalOutcome: option?.key as ApprovalOutcome }))}
          />

          <TextField
            label="Comments (applied to all items)"
            value={state.approvalComments}
            onChange={(_, newValue) => setState(prev => ({ ...prev, approvalComments: newValue || '' }))}
            multiline
            rows={4}
            placeholder="Enter approval comments..."
          />
        </Stack>

        <DialogFooter>
          <DefaultButton onClick={closeBulkApprovalDialog} text="Cancel" disabled={state.isProcessing} />
          <PrimaryButton
            onClick={handleBulkApproval}
            text={state.approvalOutcome === ApprovalOutcome.Approved ? 'Approve All' : 'Reject All'}
            disabled={state.isProcessing}
            iconProps={{ iconName: 'CheckMark' }}
          />
          {state.isProcessing && <Spinner size={SpinnerSize.small} />}
        </DialogFooter>
      </Dialog>
    </div>
  );
}; 