import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DemoSpfxWebPartStrings';
import DemoSpfx from './components/DemoSpfx';
import { IDemoSpfxProps } from './components/IDemoSpfxProps';

// DG - 09/09/2021 - Supporting section backgrounds
import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme,
  ISemanticColors
} from '@microsoft/sp-component-base';
import { ListSubscriptionFactory } from '@microsoft/sp-list-subscription';
//////////// DG - 09/09/2021

// DG - 10/09/2021 - PnP Controls and Properties Pane
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldTextWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldTextWithCallout';
//////////// DG - 10/09/2021

export interface IDemoSpfxWebPartProps {
  description: string;
  libraryId?: string;
  siteUrl?: string;
  title: string;
}

export default class DemoSpfxWebPart extends BaseClientSideWebPart<IDemoSpfxWebPartProps> {
  // DG - 09/09/2021 - Supporting section backgrounds
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;  

  protected onInit(): Promise<void> {
    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    return super.onInit();
  }
  //////////// DG - 09/09/2021

  public render(): void {
    const element: React.ReactElement<IDemoSpfxProps> = React.createElement(
      DemoSpfx,
      {
        themeVariant: this._themeVariant, // DG - 09/09/2021 - Supporting section backgrounds
        width: this.width, // DG - 10/09/2021 - Determine the rendered web part size
        description: this.properties.description,
        wpContext: this.context, // DG - 10/09/2021 - Data Service
        displayMode: this.displayMode, // DG - 10/09/2021 - Subscribe to list notifications
        libraryId: this.properties.libraryId, // DG - 10/09/2021 - Subscribe to list notifications
        listSubscriptionFactory: new ListSubscriptionFactory(this), // DG - 10/09/2021 - Subscribe to list notifications
        onConfigure: this._onConfigure, // DG - 10/09/2021 - Subscribe to list notifications
        siteUrl: this.properties.siteUrl, // DG - 10/09/2021 - Subscribe to list notifications
        title: this.properties.title, // DG - 10/09/2021 - Subscribe to list notifications
        updateProperty: value => this.properties.title = value // DG - 10/09/2021 - Subscribe to list notifications
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

  // DG - 10/09/2021 - Determine the rendered web part size
  protected onAfterResize(newWidth: number) {
    console.log(`the new width of the web part is ${newWidth}`);
  }
  //////////// DG - 10/09/2021

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
                  baseTemplate: 101
                })
              ]
            }
          ]
        }
      ]
    };
  }

  // DG - 09/09/2021 - Supporting section backgrounds
  /**
   * Update the current theme variant reference and re-render.
   *
   * @param args The new theme
   */
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }
  //////////// DG - 09/09/2021

  // DG - 10/09/2021 - Subscribe to list notifications
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onPropertyPaneConfigurationComplete(): void {
    // ideally, we'd call a refresh here to update the list of properties
    // but due to a bug in the list picker control, lists are loaded only
    // on component mount, so this wouldn't do anything
    // https://github.com/pnp/sp-dev-fx-property-controls/issues/109
    // this.context.propertyPane.refresh();
  }

  private _onConfigure = () => this.context.propertyPane.open();
  //////////// DG - 10/09/2021
}