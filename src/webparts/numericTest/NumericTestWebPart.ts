import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'NumericTestWebPartStrings';
import NumericTest from './components/NumericTest';
import { INumericTestProps } from './components/INumericTestProps';

import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldTextWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldTextWithCallout';

export interface INumericTestWebPartProps {
  description: string;
  libraryId?: string;
  siteUrl?: string;
}

export default class NumericTestWebPart extends BaseClientSideWebPart<INumericTestWebPartProps> {

  public render(): void {
    const element: React.ReactElement<INumericTestProps> = React.createElement(
      NumericTest,
      {
        description: this.properties.description,
        wpContext: this.context, // DG - 10/09/2021 - Data Service
        libraryId: this.properties.libraryId, // DG - 10/09/2021 - Subscribe to list notifications
        siteUrl: this.properties.siteUrl, // DG - 10/09/2021 - Subscribe to list notifications
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldTextWithCallout('siteUrl', {
                  calloutTrigger: CalloutTriggers.Click,
                  key: 'siteUrlFieldId',
                  label: 'Site URL',
                  calloutContent: React.createElement('span', {}, 'URL of the site where the document library to show documents from is located. Leave empty to connect to a document library from the current site'),
                  calloutWidth: 250,
                  value: this.properties.siteUrl
                }),
                PropertyFieldListPicker('libraryId', {
                  label: 'Select a document library',
                  selectedList: this.properties.libraryId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId',
                  webAbsoluteUrl: this.properties.siteUrl,
                  baseTemplate: 100
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
