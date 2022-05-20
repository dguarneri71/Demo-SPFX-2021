import * as React from 'react';
import styles from './NumericTest.module.scss';
import { INumericTestProps, INumericTestState } from './INumericTestProps';
import { IDataService } from '../../../classes/services/IDataService';
import DataServiceProvider from '../../../classes/services/DataServiceProvider';

export default class NumericTest extends React.Component<INumericTestProps, INumericTestState> {
  private _dataService: IDataService;

  constructor(props: INumericTestProps) {
    super(props);
    this._dataService = DataServiceProvider.get(this.props.wpContext);

    this.state = {
      items: []
    };
  }

  public componentDidMount(): void {
    if (!this.props.libraryId) {
      return;
    }

    this._loadData();
  }

  public componentDidUpdate(prevProps: Readonly<INumericTestProps>, prevState: Readonly<INumericTestState>, snapshot?: any): void {
    if (this.props.libraryId === prevProps.libraryId) {
      // something has changed but the library id is the same so no need to
      // reload documents or configure the subscription
      return;
    }

    this._loadData();
  }

  public render(): React.ReactElement<INumericTestProps> {
    let count = this.state.items.length;
    return (
      <div className={styles.numericTest}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span>{count}</span>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _loadData(): void {
    // communicate loading documents to the user
    this.setState({
      items: []
    });

    this._dataService.loadItems(this.props.siteUrl, this.props.libraryId)
      .then(docs => {
        console.log("Items: ", docs);
        this.setState({
          items: docs
        });
      });
  }

}
