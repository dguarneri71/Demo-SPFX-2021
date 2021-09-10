// DG - 09/09/2021 - Supporting section backgrounds
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { DisplayMode } from '@microsoft/sp-core-library';
import { ListSubscriptionFactory } from '@microsoft/sp-list-subscription';
import { WebPartContext } from '@microsoft/sp-webpart-base';
//////////// DG - 09/09/2021

export interface IDemoSpfxProps {
  themeVariant: IReadonlyTheme | undefined; // DG - 09/09/2021 - Supporting section backgrounds
  width: number;
  description: string;
  wpContext: WebPartContext;
  /**
   * Web part display mode
   */
  displayMode: DisplayMode;
  /**
   * ID of the list to retrieve documents from. Undefined, if no library
   * has been selected
   */
  libraryId?: string;
  /**
   * Instance of the ListSubscriptionFactory to use to create a list
   * subscription
   */
  listSubscriptionFactory: ListSubscriptionFactory;
  /**
   * Event handler after clicking the 'Configure' button in the Placeholder
   * component
   */
  onConfigure: () => void;
  /**
   * URL of the site where the selected library is located. Undefined, if the
   * selected library is in the current site
   */
  siteUrl?: string;
  /**
   * Web part title to show in the body
   */
  title: string;
  /**
   * Event handler after updating the web part title
   */
  updateProperty: (value: string) => void;
}
