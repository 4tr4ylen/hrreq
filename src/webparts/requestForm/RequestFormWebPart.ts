import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

const strings = {
  PropertyPaneDescription: 'Configure the HR Request Form settings',
  BasicGroupName: 'Basic Settings',
  TitleFieldLabel: 'Form Title',
  DescriptionFieldLabel: 'Form Description',
  RequestTypesFieldLabel: 'Request Types',
  RequestTypesFieldDescription: 'Enter request types separated by commas (e.g., Leave Request, Equipment Request, Policy Question)',
  MaxFileSizeFieldLabel: 'Maximum File Size (MB)',
  MaxFileSizeFieldDescription: 'Maximum allowed file size in megabytes',
  ShowManagerFieldLabel: 'Show Manager Field',
  RequireManagerApprovalLabel: 'Require Manager Approval'
};
import { RequestForm } from './components/RequestForm';
import { IRequestFormWebPartProps } from './IRequestFormWebPartProps';

export default class RequestFormWebPart extends BaseClientSideWebPart<IRequestFormWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IRequestFormWebPartProps> = React.createElement(
      RequestForm,
      {
        context: this.context,
        title: this.properties.title,
        description: this.properties.description,
        requestTypes: this.properties.requestTypes,
        maxFileSize: this.properties.maxFileSize,
        showManagerField: this.properties.showManagerField,
        requireManagerApproval: this.properties.requireManagerApproval
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  value: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  multiline: true,
                  rows: 3
                }),
                PropertyPaneTextField('requestTypes', {
                  label: strings.RequestTypesFieldLabel,
                  description: strings.RequestTypesFieldDescription,
                  multiline: true,
                  rows: 3
                }),
                PropertyPaneTextField('maxFileSize', {
                  label: strings.MaxFileSizeFieldLabel,
                  description: strings.MaxFileSizeFieldDescription
                }),
                PropertyPaneToggle('showManagerField', {
                  label: strings.ShowManagerFieldLabel
                }),
                PropertyPaneToggle('requireManagerApproval', {
                  label: strings.RequireManagerApprovalLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
} 