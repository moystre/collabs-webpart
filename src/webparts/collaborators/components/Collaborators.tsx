import * as React from 'react';
import styles from './Collaborators.module.scss';
import { ICollaboratorsProps } from './ICollaboratorsProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class Collaborators extends React.Component<ICollaboratorsProps, {}> {
  public render(): React.ReactElement<ICollaboratorsProps> {
    return (
      <div className={styles.collaborators}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Collaborators</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}</p>

              {this.props.lists.map(list => <span>{list}</span>)}

              <p className={styles.description}>Loading from {escape(this.context)}</p>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
