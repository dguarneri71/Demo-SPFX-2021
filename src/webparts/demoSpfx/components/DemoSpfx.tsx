import * as React from 'react';
import styles from './DemoSpfx.module.scss';
import { IDemoSpfxProps, IDemoSpfxState } from '.';
import { escape } from '@microsoft/sp-lodash-subset';

// DG - 09/09/2021 - Supporting section backgrounds and Subscribe to list notification
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { Guid } from '@microsoft/sp-core-library';
import { IListSubscription } from '@microsoft/sp-list-subscription';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { ListView, SelectionMode } from "@pnp/spfx-controls-react/lib/ListView";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
//////////// DG - 09/09/2021

// DG - 10/09/2021 - Data Service
import { IDataService } from '../../../classes/services/IDataService';
import DataServiceProvider from '../../../classes/services/DataServiceProvider';
//////////// DG - 10/09/2021

export default class DemoSpfx extends React.Component<IDemoSpfxProps, IDemoSpfxState> {
  // DG - 10/09/2021 - Subscribe to list notifications
  private _listSubscription: IListSubscription;
  // DG - 10/09/2021 - Data Service
  private _dataService: IDataService;

  constructor(props: IDemoSpfxProps) {
    super(props);
    this._dataService = DataServiceProvider.get(this.props.wpContext);

    this.state = {
      documents: [],
      error: undefined,
      loading: true
    };
  }

  public componentDidMount(): void {
    if (!this.props.libraryId) {
      return;
    }

    this._configureListSubscription();
    this._loadDocuments();
  }

  public componentDidUpdate(prevProps: Readonly<IDemoSpfxProps>, prevState: Readonly<IDemoSpfxState>, snapshot?: any): void {
    if (this.props.libraryId === prevProps.libraryId) {
      // something has changed but the library id is the same so no need to
      // reload documents or configure the subscription
      return;
    }

    this._configureListSubscription();
    this._loadDocuments();
  }
  //////////// DG - 10/09/2021

  public render(): React.ReactElement<IDemoSpfxProps> {

    const { semanticColors }: IReadonlyTheme = this.props.themeVariant; // DG - 09/09/2021 - Supporting section backgrounds
    var width: number = this.props.width; // DG - 10/09/2021 - Determine the rendered web part size

    // DG - 10/09/2021 - Subscribe to list notifications
    const { onConfigure } = this.props;
    const needsConfiguration: boolean = !this.props.libraryId;
    const { error, documents, loading } = this.state;
    //////////// DG - 10/09/2021

    return (
      <div className={styles.demoSpfx} style={{ backgroundColor: semanticColors.bodyBackground }}>
        <WebPartTitle displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateProperty} />
        <div><p className={styles.description}>{escape(this.props.description)} (Width: {width})</p></div>
        {needsConfiguration &&
          <Placeholder
            iconName='Edit'
            iconText='Configure your web part'
            description='Please configure the web part.'
            buttonLabel='Configure'
            onConfigure={onConfigure} />
        }
        {!needsConfiguration &&
          loading &&
          <div style={{ textAlign: 'center' }}><Spinner size={SpinnerSize.large} label="Loading documents..." /></div>}
        {!needsConfiguration &&
          !loading &&
          error &&
          <div style={{ textAlign: 'center' }}>The following error has occurred while loading documents: <span>{error}</span></div>}
        {!needsConfiguration &&
          !loading &&
          documents.length === 0 &&
          <div style={{ textAlign: 'center' }}>No documents found in the selected list</div>}
        {!needsConfiguration &&
          !loading &&
          documents.length > 0 &&
          <ListView
            items={documents}
            viewFields={[{
              displayName: 'Name',
              name: 'FileLeafRef',
              linkPropertyName: 'FileRef'
            }]}
            iconFieldName="FileRef"
            compact={false}
            selectionMode={SelectionMode.none} />
        }
      </div>
    );
  }

  /**
  * Loads documents from the selected document library
  */
  private _loadDocuments(): void {
    // communicate loading documents to the user
    this.setState({
      documents: [],
      error: undefined,
      loading: true
    });

    this._dataService.loadDocuments(this.props.siteUrl, this.props.libraryId)
      .then(docs => {
        this.setState({
          documents: docs,
          loading: false
        });
      })
      .catch(err => {
        this.setState({
          error: err,
          loading: false
        });
      });
  }

  /**
   * Subscribes to changes in a list documentary using the SharePoint Framework
   * ListSubscriptionFactory.
   * TODO: portare codice in SPDataService
   */
  private _configureListSubscription(): void {
    if (!this.props.libraryId) {
      // no library selected. If there is an existing list subscription, remove it
      if (this._listSubscription) {
        this.props.listSubscriptionFactory.deleteSubscription(this._listSubscription);
      }

      return;
    }

    // if the selected library is located in a different site (collection),
    // we need site collection and site id to setup the list subscription
    let siteCollectionId: string, siteId: string;
    this
      ._getSiteCollectionId(this.props.siteUrl)
      .then((id: string | undefined): Promise<void | string> => {
        siteCollectionId = id;
        return this._getSiteId(this.props.siteUrl);
      })
      .then((id: string | undefined): void => {
        siteId = id;
        // remove existing subscription if any
        if (this._listSubscription) {
          this.props.listSubscriptionFactory.deleteSubscription(this._listSubscription);
        }

        this.props.listSubscriptionFactory.createSubscription({
          siteId: siteCollectionId ? Guid.parse(siteCollectionId) : undefined,
          webId: siteId ? Guid.parse(siteId) : undefined,
          listId: Guid.parse(this.props.libraryId),
          callbacks: {
            notification: this._loadDocuments.bind(this)
          }
        });
      });
  }

  /**
   * Retrieves the ID of the specified site collection
   * 
   * If no URL is specified, returns an empty resolved promise.
   * TODO: portare codice in SPDataService
   * 
   * @param siteUrl URL of the site collection for which to retrieve the ID
   */
  private _getSiteCollectionId(siteUrl?: string): Promise<void | string> {
    if (!siteUrl) {
      return Promise.resolve();
    }

    return this._dataService.getSiteCollectionId(siteUrl);
  }

  /**
   * Retrieves the ID of the specified site
   * 
   * If no URL is specified, returns a empty resolved promise.
   * TODO: portare codice in SPDataService
   * 
   * @param siteUrl URL of the site for which to retrieve the ID
   */
  private _getSiteId(siteUrl?: string): Promise<void | string> {
    if (!siteUrl) {
      return Promise.resolve();
    }

    /* return new Promise<string>((resolve: (siteId: string) => void, reject: (error: any) => void): void => {
      const web: IWeb = Web(siteUrl);
      web.select('Id').get()
        .then(({ Id }): void => {
          resolve(Id);
        })
        .catch(err => reject(err));
    }); */
    return this._dataService.getSiteId(siteUrl);
  }
}