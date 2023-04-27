import * as React from 'react';
import styles from './HolidayList.module.scss';
import { IHolidayListProps } from './IHolidayListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import HolidayDetails from '../../../CommonComponents/HolidayList/HolidayDetails';

export default class HolidayList extends React.Component<IHolidayListProps, {}> {
  public render(): React.ReactElement<IHolidayListProps> {
    const {
      hasTeamsContext,
      webpartTitle
    } = this.props;

    return (
      <section className={`${styles.holidayList} ${hasTeamsContext ? styles.teams : ''}`}>
        {/* <div>
          <h2 className="page-title">{webpartTitle}</h2>
        </div> */}
        <HolidayDetails context={this.props.context} title={webpartTitle}/>
      </section>
    );
  }
}
