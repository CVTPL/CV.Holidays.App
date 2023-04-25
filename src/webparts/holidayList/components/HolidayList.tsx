import * as React from 'react';
import styles from './HolidayList.module.scss';
import { IHolidayListProps } from './IHolidayListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import HolidayDetails from '../../../CommonComponents/HolidayList/HolidayDetails';

export default class HolidayList extends React.Component<IHolidayListProps, {}> {
  public render(): React.ReactElement<IHolidayListProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.holidayList} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
          <HolidayDetails />
        </div>
      </section>
    );
  }
}
