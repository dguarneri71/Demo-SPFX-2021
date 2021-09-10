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
//////////// DG - 09/09/2021

export interface IDemoSpfxWebPartProps {
  description: string;
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
        description: this.properties.description
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
}
