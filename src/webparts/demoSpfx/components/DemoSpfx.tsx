import * as React from 'react';
import styles from './DemoSpfx.module.scss';
import { IDemoSpfxProps } from './IDemoSpfxProps';
import { escape } from '@microsoft/sp-lodash-subset';

// DG - 09/09/2021 - Supporting section backgrounds
import { IReadonlyTheme } from '@microsoft/sp-component-base';
//////////// DG - 09/09/2021

export default class DemoSpfx extends React.Component<IDemoSpfxProps, {}> {
  public render(): React.ReactElement<IDemoSpfxProps> {

    const { semanticColors }: IReadonlyTheme = this.props.themeVariant; // DG - 09/09/2021 - Supporting section backgrounds
    var width: number = this.props.width; // DG - 10/09/2021 - Determine the rendered web part size

    return (
      <div className={styles.demoSpfx} style={{ backgroundColor: semanticColors.bodyBackground }}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)} - {width}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}