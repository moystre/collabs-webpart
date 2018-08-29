import * as React from 'react';
import styles from './Collaborators.module.scss';
import { ICollaboratorsProps } from './ICollaboratorsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  IDetailsList
} from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { createRef } from 'office-ui-fabric-react/lib/Utilities';

export default class Collaborators extends React.Component<ICollaboratorsProps, {}> {
  public render(): React.ReactElement<ICollaboratorsProps> {
    return (
      <div className={styles.collaborators}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Collaborators</span>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <p className={styles.description}>Loading from {escape(this.context)}</p>
              <hr></hr>
            </div>
          </div>
          <div>
            {this.props.ispLists.map(ispLists =>
              <ul>
                <li>
                  <span>{ispLists}</span>
                </li>
              </ul>
            )}
          </div>
          <DetailsList 
            items={this.props.ispLists}
            columns={this.props.columns} >
          </DetailsList>
        </div>
      </div>
    );
  }
}
