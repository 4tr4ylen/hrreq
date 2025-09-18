import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  TextField,
  PrimaryButton,
  DefaultButton,
  Stack,
  Label,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  ChoiceGroup,
  IChoiceGroupOption,
  IconButton,
  Dialog,
  DialogType,
  DialogFooter
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { IRequestFormWebPartProps } from '../IRequestFormWebPartProps';
import { SharePointService } from '../../../services/SharePointService';
import { GraphService } from '../../../services/GraphService';
import { IRequestFormData, IUser } from '../../../models/IRequest';
import { ValidationUtils, IValidationResult } from '../../../utils/validation';
import { AttachmentUtils } from '../../../utils/attachments';
// import { IGraphUser } from '../../../services/GraphService';

export interface IRequestFormState {
  formData: IRequestFormData;
  selectedFile: File | null;
  isLoading: boolean;
  isSubmitting: boolean;
  validationErrors: string[];
  showSuccessDialog: boolean;
  createdRequestId: number | null;
  currentUser: IUser | null;
  departments: string[];
  requestTypeOptions: IChoiceGroupOption[];
}

export const RequestForm: React.FC<IRequestFormWebPartProps> = (props) => {
  const [state, setState] = useState<IRequestFormState>({
    formData: {
      title: '',
      requestType: '',
      description: '',
      department: '',
      manager: undefined,
      attachment: undefined
    },
    selectedFile: null,
    isLoading: true,
    isSubmitting: false,
    validationErrors: [],
    showSuccessDialog: false,
    createdRequestId: null,
    currentUser: null,
    departments: [],
    requestTypeOptions: []
  });

  const sharePointService = new SharePointService(props.context);
  const graphService = new GraphService(props.context);

  useEffect(() => {
    initializeForm();
  }, []);

  const initializeForm = async (): Promise<void> => {
    try {
      setState(prev => ({ ...prev, isLoading: true }));

      // Get current user
      const currentUser = await sharePointService.getCurrentUser();
      
      // Get user's department from Graph
      const userDepartment = await graphService.getCurrentUserDepartment();
      
      // Get all departments
      const departments = await graphService.getAllDepartments();
      
      // Parse request types from properties
      const requestTypes = props.requestTypes.split(',').map(type => type.trim()).filter(type => type.length > 0);
      const requestTypeOptions: IChoiceGroupOption[] = requestTypes.map(type => ({
        key: type,
        text: type
      }));

      setState(prev => ({
        ...prev,
        currentUser,
        departments,
        requestTypeOptions,
        formData: {
          ...prev.formData,
          department: userDepartment || currentUser.Department || ''
        },
        isLoading: false
      }));
    } catch (error) {
      console.error('Error initializing form:', error);
      setState(prev => ({
        ...prev,
        validationErrors: ['Failed to initialize form. Please refresh the page.'],
        isLoading: false
      }));
    }
  };

  const handleInputChange = (field: keyof IRequestFormData, value: any): void => {
    setState(prev => ({
      ...prev,
      formData: {
        ...prev.formData,
        [field]: value
      },
      validationErrors: []
    }));
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const file = event.target.files?.[0] || null;
    
    if (file) {
      const maxSizeMB = parseInt(props.maxFileSize) || 10;
      const validationErrors = AttachmentUtils.validateSingleFile(file, maxSizeMB);
      
      setState(prev => ({
        ...prev,
        selectedFile: file,
        formData: {
          ...prev.formData,
          attachment: file
        },
        validationErrors
      }));
    } else {
      setState(prev => ({
        ...prev,
        selectedFile: null,
        formData: {
          ...prev.formData,
          attachment: undefined
        },
        validationErrors: ['Please select a file to upload']
      }));
    }
  };

  const handleManagerChange = (items: any[]): void => {
    if (items && items.length > 0) {
      const manager: IUser = {
        Id: items[0].id,
        Title: items[0].displayName,
        Email: items[0].secondaryText,
        DisplayName: items[0].displayName
      };
      handleInputChange('manager', manager);
    } else {
      handleInputChange('manager', undefined);
    }
  };

  const validateForm = (): IValidationResult => {
    return ValidationUtils.validateRequestForm(state.formData);
  };

  const handleSubmit = async (): Promise<void> => {
    try {
      // Validate form
      const validation = validateForm();
      if (!validation.isValid) {
        setState(prev => ({
          ...prev,
          validationErrors: validation.errors
        }));
        return;
      }

      setState(prev => ({ ...prev, isSubmitting: true, validationErrors: [] }));

      // Create request
      const requestData = {
        Title: state.formData.title,
        RequestType: state.formData.requestType,
        Description: state.formData.description,
        Department: state.formData.department,
        Requestor: state.currentUser,
        Manager: state.formData.manager,
        Status: 'Submitted'
      };

      const createdRequest = await sharePointService.createRequest(requestData as any);

      // Upload attachment if file is selected
      if (state.selectedFile) {
        const fileBuffer = await AttachmentUtils.fileToArrayBuffer(state.selectedFile);
        await sharePointService.uploadAttachment(createdRequest.Id as number, state.selectedFile.name, fileBuffer);
      }

      setState(prev => ({
        ...prev,
        showSuccessDialog: true,
        createdRequestId: (createdRequest.Id as number) ?? null,
        isSubmitting: false
      }));

    } catch (error) {
      console.error('Error submitting request:', error);
      setState(prev => ({
        ...prev,
        validationErrors: ['Failed to submit request. Please try again.'],
        isSubmitting: false
      }));
    }
  };

  const handleReset = (): void => {
    setState(prev => ({
      ...prev,
      formData: {
        title: '',
        requestType: '',
        description: '',
        department: prev.formData.department,
        manager: undefined,
        attachment: undefined
      },
      selectedFile: null,
      validationErrors: []
    }));
  };

  const closeSuccessDialog = (): void => {
    setState(prev => ({
      ...prev,
      showSuccessDialog: false,
      createdRequestId: null
    }));
    handleReset();
  };

  const getRequestItemUrl = (): string => {
    if (state.createdRequestId) {
      return `${props.context.pageContext.web.absoluteUrl}/Lists/HR%20Requests/DispForm.aspx?ID=${state.createdRequestId}`;
    }
    return '';
  };

  if (state.isLoading) {
    return (
      <Stack horizontalAlign="center" verticalAlign="center" style={{ minHeight: '200px' }}>
        <Spinner size={SpinnerSize.large} label="Loading form..." />
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

        {/* Error Messages */}
        {state.validationErrors.length > 0 && (
          <MessageBar messageBarType={MessageBarType.error}>
            <ul style={{ margin: 0, paddingLeft: '20px' }}>
              {state.validationErrors.map((error, index) => (
                <li key={index}>{error}</li>
              ))}
            </ul>
          </MessageBar>
        )}

        {/* Form */}
        <Stack tokens={{ childrenGap: 10 }}>
          {/* Title */}
          <TextField
            label="Request Title *"
            required
            value={state.formData.title}
            onChange={(_, newValue) => handleInputChange('title', newValue)}
            maxLength={255}
          />

          {/* Request Type */}
          <ChoiceGroup
            label="Request Type *"
            required
            options={state.requestTypeOptions}
            selectedKey={state.formData.requestType}
            onChange={(_, option) => handleInputChange('requestType', option?.key || '')}
          />

          {/* Description */}
          <TextField
            label="Description *"
            required
            multiline
            rows={4}
            value={state.formData.description}
            onChange={(_, newValue) => handleInputChange('description', newValue)}
            maxLength={2000}
          />

          {/* Department */}
          <TextField
            label="Department *"
            required
            value={state.formData.department}
            onChange={(_, newValue) => handleInputChange('department', newValue)}
          />

          {/* Manager Field */}
          {props.showManagerField && (
            <div>
              <Label>Manager (Optional)</Label>
              <PeoplePicker
                context={props.context as any}
                titleText=""
                personSelectionLimit={1}
                showtooltip={true}
                required={false}
                disabled={false}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
                onChange={handleManagerChange}
              />
            </div>
          )}

          {/* File Upload */}
          <div>
            <Label>Attachment *</Label>
            <input
              type="file"
              onChange={handleFileChange}
              accept={ValidationUtils.getAllowedFileTypes().map(type => `.${type}`).join(',')}
              style={{ marginTop: '5px' }}
            />
          {state.selectedFile && (
              <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: '10px' }}>
                <IconButton
                  iconProps={{ iconName: AttachmentUtils.getFileIcon(state.selectedFile.name) }}
                  style={{ color: AttachmentUtils.getFileColor(state.selectedFile.name) }}
                />
                <span>{state.selectedFile.name}</span>
                <span>({AttachmentUtils.formatFileSize(state.selectedFile.size)})</span>
              </Stack>
            )}
            <div style={{ fontSize: '12px', color: '#666', marginTop: '5px' }}>
              Maximum file size: {props.maxFileSize || 10}MB. Allowed types: {ValidationUtils.getAllowedFileTypes().join(', ')}
            </div>
          </div>

          {/* Action Buttons */}
          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <PrimaryButton
              text="Submit Request"
              onClick={handleSubmit}
              disabled={state.isSubmitting}
              iconProps={{ iconName: 'Send' }}
            />
            <DefaultButton
              text="Reset"
              onClick={handleReset}
              disabled={state.isSubmitting}
              iconProps={{ iconName: 'Refresh' }}
            />
            {state.isSubmitting && <Spinner size={SpinnerSize.small} />}
          </Stack>
        </Stack>
      </Stack>

      {/* Success Dialog */}
      <Dialog
        hidden={!state.showSuccessDialog}
        onDismiss={closeSuccessDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Request Submitted Successfully',
          subText: 'Your HR request has been submitted and is now being processed.'
        }}
      >
        <DialogFooter>
          <PrimaryButton onClick={closeSuccessDialog} text="Close" />
          {state.createdRequestId && (
            <DefaultButton
              onClick={() => window.open(getRequestItemUrl(), '_blank')}
              text="View Request"
              iconProps={{ iconName: 'OpenInNewWindow' }}
            />
          )}
        </DialogFooter>
      </Dialog>
    </div>
  );
}; 