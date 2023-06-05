import * as React from 'react';
import styles from './HolidayList.module.scss';
import { IHolidayListProps } from './IHolidayListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import HolidayDetails from '../../../CommonComponents/HolidayList/HolidayDetails';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { getTheme, initializeIcons } from 'office-ui-fabric-react';
import PnpSpCommonServices from '../../../services/PnpSpCommonServices';
import { spfi, SPFx } from "@pnp/sp";
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import "@pnp/sp/files/web";

export default class HolidayList extends React.Component<IHolidayListProps, any, {}> {
  constructor(props: IHolidayListProps) {
    super(props);

    this.state = {
      isDataAvailable: false,
      loading: true,
      placeholderDisplay: false
    };
    initializeIcons(undefined, { disableWarnings: true });
  }
  public sp = spfi().using(SPFx(this.props.context));
  componentDidMount(): void {
    // console.log("clientHeight : "+ this.props.context.domElement.clientHeight);
    // console.log("clientWidth : "+ this.props.context.domElement.clientWidth);

    //check list is exist or not
    if (Object.keys(this.props.context).length > 0) {
      // let siteUrl = this.props.dataSource && this.props.dataSource == "SiteLevel" ? this.props.context.pageContext.legacyPageContext.webAbsoluteUrl : this.props.context.pageContext.legacyPageContext.appBarParams.portalUrl;
      let siteUrl = this.props.context.pageContext.legacyPageContext.webAbsoluteUrl;

      PnpSpCommonServices._getSiteListByName(this.props.context, "Holiday Details").then((response) => {
        if (response.status == 404) {//list is not available

          //list is not available than check site design available
          PnpSpCommonServices._getSiteDesign(this.sp).then((allSiteDesign) => {
            let checkSiteDesign = allSiteDesign.filter((ele: any) => ele.Title == "HolidayDetailsSiteDesign");
            if (checkSiteDesign.length > 0) {
              //site design is available so apply that site design to site.
              return PnpSpCommonServices._applySiteDesignToSite(this.sp, checkSiteDesign[0].Id, siteUrl).then((response) => {
                return this._commonFlowAfterSideDesignApply();
              });
            }
            else {
              //site design is not available then check site script available
              return PnpSpCommonServices._getSiteScript(this.sp).then((allSiteScripts) => {
                let checkSiteScript = allSiteScripts.filter((ele: any) => ele.Title == "HolidayDetailsSiteScript");
                if (checkSiteScript.length > 0) {
                  //site script is available so create site design and apply to site
                  return PnpSpCommonServices._createSiteDesign(this.sp, checkSiteScript[0].Id).then((response) => {
                    return PnpSpCommonServices._applySiteDesignToSite(this.sp, response.Id, siteUrl);
                  }).then((response) => {
                    return this._commonFlowAfterSideDesignApply();
                  });
                }
                else {
                  // site script is not available so create site script and site design and apply to site
                  PnpSpCommonServices._createSiteScript(this.props.context, this.sp).then((response) => {
                    return PnpSpCommonServices._createSiteDesign(this.sp, response.Id);
                  }).then((response) => {
                    return PnpSpCommonServices._applySiteDesignToSite(this.sp, response.Id, siteUrl);
                  }).then((response) => {
                    return this._commonFlowAfterSideDesignApply();
                  });
                }
              });
            }
          });
        }
        else {
          return this._getCurrentYearHolidays();
        }
      });
    }
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
  private async _addDefaultItemsInList(items: any): Promise<any> {
    return new Promise((resolve, reject) => {
      PnpSpCommonServices._addItemsUsingBatch(this.sp, "Holiday Details", items).then((response) => {
        resolve(response);
      },
        (error: any): any => {
          reject(error);
        });
    });
  }
  private _setupImageObject = async () => {

    // USA
    let LaborDayImg: any = await fetch(require("../../../assets/img/defaultImages/US/Labor-Day.jpg"));
    LaborDayImg = await LaborDayImg.blob();

    let MemorialDayImg: any = await fetch(require("../../../assets/img/defaultImages/US/Memorial-Day.jpg"));
    MemorialDayImg = await MemorialDayImg.blob();

    let USIndependenceDayImg: any = await fetch(require("../../../assets/img/defaultImages/US/US-Independence-Day.jpg"));
    USIndependenceDayImg = await USIndependenceDayImg.blob();

    // INDIA
    let RepublicDayImg: any = await fetch(require("../../../assets/img/defaultImages/INDIA/republic-day.jpg"));
    RepublicDayImg = await RepublicDayImg.blob();

    let INDIAIndependenceDayImg: any = await fetch(require("../../../assets/img/defaultImages/INDIA/INDIA-Independence-Day.jpg"));
    INDIAIndependenceDayImg = await INDIAIndependenceDayImg.blob();

    let GandhiJayantiImg: any = await fetch(require("../../../assets/img/defaultImages/INDIA/Gandhi-Jayanti.jpg"));
    GandhiJayantiImg = await GandhiJayantiImg.blob();

    // Common
    let ChristmasDayImg: any = await fetch(require("../../../assets/img/defaultImages/Common/Christmas-Day.jpg"));
    ChristmasDayImg = await ChristmasDayImg.blob();

    let NewYearImg: any = await fetch(require("../../../assets/img/defaultImages/Common/New-Year.jpg"));
    NewYearImg = await NewYearImg.blob();

    let imagesArray = [
      {
        fileName: "Labor-Day.jpg",
        fileContent: LaborDayImg
      },
      {
        fileName: "Memorial-Day.jpg",
        fileContent: MemorialDayImg
      },
      {
        fileName: "US-Independence-Day.jpg",
        fileContent: USIndependenceDayImg
      },
      {
        fileName: "republic-day.jpg",
        fileContent: RepublicDayImg
      },
      {
        fileName: "INDIA-Independence-Day.jpg",
        fileContent: INDIAIndependenceDayImg
      },
      {
        fileName: "Gandhi-Jayanti.jpg",
        fileContent: GandhiJayantiImg
      },
      {
        fileName: "Christmas-Day.jpg",
        fileContent: ChristmasDayImg
      },
      {
        fileName: "New-Year.jpg",
        fileContent: NewYearImg
      }
    ]

    return imagesArray;
  }
  private _commonFlowAfterSideDesignApply = async () => {
    let siteUrl = this.props.context.pageContext.legacyPageContext.webAbsoluteUrl;
    let listId = "";

    PnpSpCommonServices._ensureSiteAssetsLibraryexist(this.sp).then((response) => {
      return PnpSpCommonServices._getFolderByPath(this.props.context, "SiteAssets/Lists")
    }).then((response) => {
      //check Lists folder in Site Assets already exists if no then create.
      if (response.status == 200) {
        return;
      }
      else {
        return PnpSpCommonServices._createFolder(this.sp, "SiteAssets/Lists");
      }
    }).then((response) => {
      return PnpSpCommonServices._getSiteListByName(this.props.context, "Holiday Details");
    }).then(async (response) => {
      return await response.json();
    }).then((response) => {
      listId = response.d.Id;
      return PnpSpCommonServices._createFolder(this.sp, "SiteAssets/Lists/" + listId + "");
    }).then(async (response) => {
      return this._setupImageObject();
    }).then(async (response) => {
      response.forEach(async image => {
        await PnpSpCommonServices._addImage(this.sp, "SiteAssets/Lists/" + listId + "", image);
      });
    }).then(async (response) => {
      return await this.sp.web.regionalSettings.timeZone();
    }).then((response) => {
      let currentYear = new Date().getFullYear();
      let nextYear = new Date().getFullYear() + 1;

      let defaultDataCurrentYear: any = [];
      let defaultDataNextYear: any = [];

      switch (response.Id) {
        case 13://US and Canada
          defaultDataCurrentYear = [
            {
              Title: "Holiday-" + currentYear + "",
              CV_Festival_Name: "New Year's Day",
              CV_Festival_Date: new Date("" + currentYear + "-01-01T00:00:00Z"),
              CV_FestivalDescription: "New Year's Day on January 1 in the Gregorian calendar is celebrated in many countries",
              CV_FestivalInfoLink: {
                Description: "New Year's Day",
                Url: "https://en.wikipedia.org/wiki/New_Year%27s_Day",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/New-Year.jpg"
              })
            },
            {
              Title: "Holiday-" + currentYear + "",
              CV_Festival_Name: "Memorial Day",
              CV_Festival_Date: this.getLastMondayOfMay(currentYear),
              CV_FestivalDescription: "Memorial Day is a federal holiday in the United States for honoring and mourning the U.S. ",
              CV_FestivalInfoLink: {
                Description: "Memorial Day",
                Url: "https://en.wikipedia.org/wiki/Memorial_Day",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/Memorial-Day.jpg"
              })
            },
            {
              Title: "Holiday-" + currentYear + "",
              CV_Festival_Name: "Independence Day (United States)",
              CV_Festival_Date: new Date("" + currentYear + "-07-04T00:00:00Z"),
              CV_FestivalDescription: "Independence Day, also called Fourth of July or July 4th, in the United States, the annual celebration of nationhood.",
              CV_FestivalInfoLink: {
                Description: "Independence Day (United States)",
                Url: "https://en.wikipedia.org/wiki/Independence_Day_(United_States)",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/US-Independence-Day.jpg"
              })
            },
            {
              Title: "Holiday-" + currentYear + "",
              CV_Festival_Name: "Labor Day",
              CV_Festival_Date: this.getFirstMondayOfSeptember(currentYear),
              CV_FestivalDescription: "Labor Day is a federal holiday in the United States celebrated on the first Monday in September to honor and recognize the American labor movement and the works and contributions of laborers to the development and achievements of the United States.",
              CV_FestivalInfoLink: {
                Description: "Labor Day",
                Url: "https://en.wikipedia.org/wiki/Labor_Day",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/Labor-Day.jpg"
              })
            },
            {
              Title: "Holiday-" + currentYear + "",
              CV_Festival_Name: "Christmas Day",
              CV_Festival_Date: new Date("" + currentYear + "-12-25T00:00:00Z"),
              CV_FestivalDescription: "Christmas is celebrated by many Christians on December 25 in the Gregorian calendar.",
              CV_FestivalInfoLink: {
                Description: "Christmas Day",
                Url: "https://en.wikipedia.org/wiki/Christmas",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/Christmas-Day.jpg"
              })
            }
          ];
          defaultDataNextYear = [
            {
              Title: "Holiday-" + nextYear + "",
              CV_Festival_Name: "New Year's Day",
              CV_Festival_Date: new Date("" + nextYear + "-01-01T00:00:00Z"),
              CV_FestivalDescription: "New Year's Day on January 1 in the Gregorian calendar is celebrated in many countries",
              CV_FestivalInfoLink: {
                Description: "New Year's Day",
                Url: "https://en.wikipedia.org/wiki/New_Year%27s_Day",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/New-Year.jpg"
              })
            },
            {
              Title: "Holiday-" + nextYear + "",
              CV_Festival_Name: "Memorial Day",
              CV_Festival_Date: this.getLastMondayOfMay(nextYear),
              CV_FestivalDescription: "Memorial Day is a federal holiday in the United States for honoring and mourning the U.S. ",
              CV_FestivalInfoLink: {
                Description: "Memorial Day",
                Url: "https://en.wikipedia.org/wiki/Memorial_Day",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/Memorial-Day.jpg"
              })
            },
            {
              Title: "Holiday-" + nextYear + "",
              CV_Festival_Name: "Independence Day (United States)",
              CV_Festival_Date: new Date("" + nextYear + "-07-04T00:00:00Z"),
              CV_FestivalDescription: "Independence Day, also called Fourth of July or July 4th, in the United States, the annual celebration of nationhood.",
              CV_FestivalInfoLink: {
                Description: "Independence Day (United States)",
                Url: "https://en.wikipedia.org/wiki/Independence_Day_(United_States)",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/US-Independence-Day.jpg"
              })
            },
            {
              Title: "Holiday-" + nextYear + "",
              CV_Festival_Name: "Labor Day",
              CV_Festival_Date: this.getFirstMondayOfSeptember(nextYear),
              CV_FestivalDescription: "Labor Day is a federal holiday in the United States celebrated on the first Monday in September to honor and recognize the American labor movement and the works and contributions of laborers to the development and achievements of the United States.",
              CV_FestivalInfoLink: {
                Description: "Labor Day",
                Url: "https://en.wikipedia.org/wiki/Labor_Day",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/Labor-Day.jpg"
              })
            },
            {
              Title: "Holiday-" + nextYear + "",
              CV_Festival_Name: "Christmas Day",
              CV_Festival_Date: new Date("" + nextYear + "-12-25T00:00:00Z"),
              CV_FestivalDescription: "Christmas is celebrated by many Christians on December 25 in the Gregorian calendar.",
              CV_FestivalInfoLink: {
                Description: "Christmas Day",
                Url: "https://en.wikipedia.org/wiki/Christmas",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/Christmas-Day.jpg"
              })
            }
          ];
          break;
        case 23://INDIA
          defaultDataCurrentYear = [
            {
              Title: "Holiday-" + currentYear + "",
              CV_Festival_Name: "New Year's Day",
              CV_Festival_Date: new Date("" + currentYear + "-01-01T00:00:00Z"),
              CV_FestivalDescription: "New Year's Day on January 1 in the Gregorian calendar is celebrated in many countries",
              CV_FestivalInfoLink: {
                Description: "New Year's Day",
                Url: "https://en.wikipedia.org/wiki/New_Year%27s_Day",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/New-Year.jpg"
              })
            },
            {
              Title: "Holiday-" + currentYear + "",
              CV_Festival_Name: "Republic Day",
              CV_Festival_Date: new Date("" + currentYear + "-01-26T00:00:00Z"),
              CV_FestivalDescription: "Republic Day is the day when India marks and celebrates the date on which the Constitution of India came into effect on 26 January 1950.",
              CV_FestivalInfoLink: {
                Description: "Republic Day",
                Url: "https://en.wikipedia.org/wiki/Republic_Day_(India)",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/republic-day.jpg"
              })
            },
            {
              Title: "Holiday-" + currentYear + "",
              CV_Festival_Name: "Indian Independence Day",
              CV_Festival_Date: new Date("" + currentYear + "-08-15T00:00:00Z"),
              CV_FestivalDescription: "Independence Day is celebrated annually on 15 August as a public holiday in India commemorating the nation's independence from the United Kingdom on 15 August 1947.",
              CV_FestivalInfoLink: {
                Description: "Indian Independence Day",
                Url: "https://en.wikipedia.org/wiki/Independence_Day_(India)",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/INDIA-Independence-Day.jpg"
              })
            },
            {
              Title: "Holiday-" + currentYear + "",
              CV_Festival_Name: "Gandhi Jayanti",
              CV_Festival_Date: new Date("" + currentYear + "-10-02T00:00:00Z"),
              CV_FestivalDescription: "Gandhi Jayanti is an event celebrated in India to mark the birthday of Mahatma Gandhi. It is celebrated annually on 2 October, and is one of the three national holidays of India.",
              CV_FestivalInfoLink: {
                Description: "Gandhi Jayanti",
                Url: "https://en.wikipedia.org/wiki/Gandhi_Jayanti",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/Gandhi-Jayanti.jpg"
              })
            },
            {
              Title: "Holiday-" + currentYear + "",
              CV_Festival_Name: "Christmas Day",
              CV_Festival_Date: new Date("" + currentYear + "-12-25T00:00:00Z"),
              CV_FestivalDescription: "Christmas is celebrated by many Christians on December 25 in the Gregorian calendar.",
              CV_FestivalInfoLink: {
                Description: "Christmas Day",
                Url: "https://en.wikipedia.org/wiki/Christmas",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/Christmas-Day.jpg"
              })
            }
          ];

          defaultDataNextYear = [
            {
              Title: "Holiday-" + nextYear + "",
              CV_Festival_Name: "New Year's Day",
              CV_Festival_Date: new Date("" + nextYear + "-01-01T00:00:00Z"),
              CV_FestivalDescription: "New Year's Day on January 1 in the Gregorian calendar is celebrated in many countries",
              CV_FestivalInfoLink: {
                Description: "New Year's Day",
                Url: "https://en.wikipedia.org/wiki/New_Year%27s_Day",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/New-Year.jpg"
              })
            },
            {
              Title: "Holiday-" + nextYear + "",
              CV_Festival_Name: "Republic Day",
              CV_Festival_Date: new Date("" + nextYear + "-01-26T00:00:00Z"),
              CV_FestivalDescription: "Republic Day is the day when India marks and celebrates the date on which the Constitution of India came into effect on 26 January 1950.",
              CV_FestivalInfoLink: {
                Description: "Republic Day",
                Url: "https://en.wikipedia.org/wiki/Republic_Day_(India)",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/republic-day.jpg"
              })
            },
            {
              Title: "Holiday-" + nextYear + "",
              CV_Festival_Name: "Indian Independence Day",
              CV_Festival_Date: new Date("" + nextYear + "-08-15T00:00:00Z"),
              CV_FestivalDescription: "Independence Day is celebrated annually on 15 August as a public holiday in India commemorating the nation's independence from the United Kingdom on 15 August 1947.",
              CV_FestivalInfoLink: {
                Description: "Indian Independence Day",
                Url: "https://en.wikipedia.org/wiki/Independence_Day_(India)",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/INDIA-Independence-Day.jpg"
              })
            },
            {
              Title: "Holiday-" + nextYear + "",
              CV_Festival_Name: "Gandhi Jayanti",
              CV_Festival_Date: new Date("" + nextYear + "-10-02T00:00:00Z"),
              CV_FestivalDescription: "Gandhi Jayanti is an event celebrated in India to mark the birthday of Mahatma Gandhi. It is celebrated annually on 2 October, and is one of the three national holidays of India.",
              CV_FestivalInfoLink: {
                Description: "Gandhi Jayanti",
                Url: "https://en.wikipedia.org/wiki/Gandhi_Jayanti",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/Gandhi-Jayanti.jpg"
              })
            },
            {
              Title: "Holiday-" + nextYear + "",
              CV_Festival_Name: "Christmas Day",
              CV_Festival_Date: new Date("" + nextYear + "-12-25T00:00:00Z"),
              CV_FestivalDescription: "Christmas is celebrated by many Christians on December 25 in the Gregorian calendar.",
              CV_FestivalInfoLink: {
                Description: "Christmas Day",
                Url: "https://en.wikipedia.org/wiki/Christmas",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/Christmas-Day.jpg"
              })
            }
          ];
          break;
        default:
          defaultDataCurrentYear = [
            {
              Title: "Holiday-" + currentYear + "",
              CV_Festival_Name: "New Year's Day",
              CV_Festival_Date: new Date("" + currentYear + "-01-01T00:00:00Z"),
              CV_FestivalDescription: "New Year's Day on January 1 in the Gregorian calendar is celebrated in many countries",
              CV_FestivalInfoLink: {
                Description: "New Year's Day",
                Url: "https://en.wikipedia.org/wiki/New_Year%27s_Day",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/New-Year.jpg"
              })
            },
            {
              Title: "Holiday-" + currentYear + "",
              CV_Festival_Name: "Christmas Day",
              CV_Festival_Date: new Date("" + currentYear + "-12-25T00:00:00Z"),
              CV_FestivalDescription: "Christmas is celebrated by many Christians on December 25 in the Gregorian calendar.",
              CV_FestivalInfoLink: {
                Description: "Christmas Day",
                Url: "https://en.wikipedia.org/wiki/Christmas",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/Christmas-Day.jpg"
              })
            }
          ];
          defaultDataNextYear = [
            {
              Title: "Holiday-" + nextYear + "",
              CV_Festival_Name: "New Year's Day",
              CV_Festival_Date: new Date("" + nextYear + "-01-01T00:00:00Z"),
              CV_FestivalDescription: "New Year's Day on January 1 in the Gregorian calendar is celebrated in many countries",
              CV_FestivalInfoLink: {
                Description: "New Year's Day",
                Url: "https://en.wikipedia.org/wiki/New_Year%27s_Day",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/New-Year.jpg"
              })
            },
            {
              Title: "Holiday-" + nextYear + "",
              CV_Festival_Name: "Christmas Day",
              CV_Festival_Date: new Date("" + nextYear + "-12-25T00:00:00Z"),
              CV_FestivalDescription: "Christmas is celebrated by many Christians on December 25 in the Gregorian calendar.",
              CV_FestivalInfoLink: {
                Description: "Christmas Day",
                Url: "https://en.wikipedia.org/wiki/Christmas",
              },
              CV_FestivalImage: JSON.stringify({
                type: 'thumbnail',
                serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/Christmas-Day.jpg"
              })
            }
          ];
          break;
      }
      return this._addDefaultItemsInList(defaultDataCurrentYear.concat(defaultDataNextYear));
    }).then((response) => {
      return this._getCurrentYearHolidays();
    });
  }
  //function for get date of Labor Day holiday
  private getFirstMondayOfSeptember = (year: any) => {
    var date: any = new Date(year, 8, 1); // first day of September
    var firstMonday = null;
    while (date.getMonth() === 8) { // while still in September
      if (date.getDay() === 1) { // if Monday
        firstMonday = new Date(date); // update first Monday
        break; // exit loop
      }
      date.setDate(date.getDate() + 1); // increment date by one day
    }
    return firstMonday;
  }
  // function for get date of Memorial Day holiday
  private getLastMondayOfMay = (year: any) => {
    var date: any = new Date(year, 4, 1); // first day of May
    var lastMonday = null;
    while (date.getMonth() === 4) { // while still in May
      if (date.getDay() === 1) { // if Monday
        lastMonday = new Date(date); // update last Monday
      }
      date.setDate(date.getDate() + 1); // increment date by one day
    }
    return lastMonday;
  }
}
