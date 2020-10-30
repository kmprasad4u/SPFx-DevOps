import * as React from 'react';
import styles from './DevOps.module.scss';
import { IDevOpsProps } from './IDevOpsProps';

interface IDevOpsState {
   
}

export class DevOps extends React.Component<IDevOpsProps, IDevOpsState> {
  public constructor(props) {
    super(props);

    this.state = {

    };

    this._refresh1 = this._refresh1.bind(this);
    this._refresh2 = this._refresh2.bind(this);
    this._refresh3 = this._refresh3.bind(this);
  }

  public componentDidMount() {
  }

  public render(): React.ReactElement<IDevOpsProps> {
    return (
      <div className={ styles.devOps }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
              <button onClick={this._refresh1}>Try 1</button>
              <button onClick={this._refresh2}>Try 2</button>
              <button onClick={this._refresh3}>Try 3</button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  public _refresh1() {
    this.props.devOpsService.getProjects1();
  }

  public _refresh2() {
    this.props.devOpsService.getProjects2();
  }

  public _refresh3() {
    this.props.devOpsService.getProjects3();
  }

}
