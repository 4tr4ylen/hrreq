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
  PropertyPaneDescription: 'Configure the HR Admin Grid settings',
  BasicGroupName: 'Basic Settings',
  TitleFieldLabel: 'Grid Title',
  DescriptionFieldLabel: 'Grid Description',
  ItemsPerPageFieldLabel: 'Items Per Page',
  ItemsPerPageFieldDescription: 'Number of items to display per page',
  ShowFiltersFieldLabel: 'Show Filters',
  EnableBulkActionsFieldLabel: 'Enable Bulk Actions'
};
import { HrAdminGrid } from './components/HrAdminGrid';
import { IHrAdminGridWebPartProps } from './IHrAdminGridWebPartProps';

export default class HrAdminGridWebPart extends BaseClientSideWebPart<IHrAdminGridWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IHrAdminGridWebPartProps> = React.createElement(
      HrAdminGrid,
      {
        context: this.context,
        title: this.properties.title,
        description: this.properties.description,
        showFilters: this.properties.showFilters,
        itemsPerPage: this.properties.itemsPerPage,
        enableBulkActions: this.properties.enableBulkActions
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
                PropertyPaneTextField('itemsPerPage', {
                  label: strings.ItemsPerPageFieldLabel,
                  description: strings.ItemsPerPageFieldDescription
                }),
                PropertyPaneToggle('showFilters', {
                  label: strings.ShowFiltersFieldLabel
                }),
                PropertyPaneToggle('enableBulkActions', {
                  label: strings.EnableBulkActionsFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
} 