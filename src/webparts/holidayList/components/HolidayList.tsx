import * as React from 'react';
import styles from './HolidayList.module.scss';
import { IHolidayListProps } from './IHolidayListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import HolidayDetails from '../../../CommonComponents/HolidayList/HolidayDetails';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { getTheme } from 'office-ui-fabric-react';
import PnpSpCommonServices from '../../../services/PnpSpCommonServices';
import { spfi, SPFx } from "@pnp/sp";
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

export default class HolidayList extends React.Component<IHolidayListProps, any, {}> {
  constructor(props: IHolidayListProps) {
    super(props);

    this.state = {
      isDataAvailable: false,
      loading: true,
      placeholderDisplay: false
    };

  }
  public sp = spfi().using(SPFx(this.props.context));
  componentDidMount(): void {
    this._getCurrentYearHolidays();
  }

  public render(): React.ReactElement<IHolidayListProps> {
    const {
      hasTeamsContext,
      webpartTitle,
    } = this.props;

    return (
      <section className={`${styles.holidayList} ${hasTeamsContext ? styles.teams : ''}`}>

        {this.state.loading &&
          <Spinner label="Loading items..." size={SpinnerSize.large} />
        }

        {this.state.isDataAvailable && this.state.isDataAvailable ?
          <HolidayDetails context={this.props.context} title={webpartTitle} />
          :
          ""}

        {this.state.placeholderDisplay ?
          <Placeholder iconName='Edit'
            iconText='Configure your web part'
            description='Details are not available in list for current year.'
            buttonLabel='Configure'
            onConfigure={this._onConfigure}
            theme={getTheme()} />
          : ""}
      </section>
    );
  }

  private _onConfigure = () => {
    // Context of the web part
    this.props.context.propertyPane.open();
  }

  private async _getCurrentYearHolidays(): Promise<any> {
    let currentYear = new Date().getFullYear();

    let selectString = "Title,CV_Festival_Name,CV_Festival_Date,CV_FestivalDescription,CV_FestivalImage,CV_FestivalInfoLink";
    let expandString = "";
    let filterString = "(CV_Festival_Date eq '" + currentYear + "-01-01T00:00:00Z' or CV_Festival_Date gt '" + currentYear + "-01-01T00:00:00Z') and (CV_Festival_Date eq '" + currentYear + "-12-31T23:59:59Z' or CV_Festival_Date lt '" + currentYear + "-12-31T23:59:59Z')";

    return new Promise((resolve, reject) => {
      PnpSpCommonServices._getListItemsWithExpandStringWithFiltersAndOrderByWithTop(this.sp, "Holiday Details", selectString, expandString, filterString, "Id", true, 4999).then((response) => {
        if (response.length > 0) {
          this.setState({ isDataAvailable: true, loading: false, placeholderDisplay: false });
        }
        else {
          this.setState({ isDataAvailable: false, loading: false, placeholderDisplay: true });
        }
        resolve(response);
      }, (error: any): any => {
        reject(error);
      });
    });
  }
}
