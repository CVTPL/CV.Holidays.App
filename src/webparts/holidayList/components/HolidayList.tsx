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

    //check list is exist or not
    if (Object.keys(this.props.context).length > 0) {
      console.log(this.props.context.domElement.clientWidth);
      
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
              }).then((response) => {
                console.log("Done");
              })
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
          // return this._checkDefaultData();
          return this._getCurrentYearHolidays();
          // return this._commonFlowAfterSideDesignApply();
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
  private async _checkDefaultData(): Promise<any> {
    let currentYear = new Date().getFullYear();
    let filterString = "CV_Festival_Name eq 'New Year's Day' and CV_Festival_Date eq " + new Date("" + currentYear + "-01-01T00:00:00Z") + "";
    return new Promise((resolve, reject) => {
      PnpSpCommonServices._getListItemsWithExpandStringWithFiltersAndOrderByWithTop(this.sp, "Holiday Details", "*", "", "filter", "Id", false, 4999).then((response) => {
        resolve(response);
      },
        (error: any) => {
          reject(error);
          console.log(error);
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

    let ChristmasDayImg: any = await fetch(require("../../../assets/img/defaultImages/Christmas-Day.jpg"));
    ChristmasDayImg = await ChristmasDayImg.blob();

    let INDIAIndependenceDayImg: any = await fetch(require("../../../assets/img/defaultImages/INDIA-Independence-Day.jpg"));
    INDIAIndependenceDayImg = await INDIAIndependenceDayImg.blob();

    let NewYearImg: any = await fetch(require("../../../assets/img/defaultImages/New-Year.jpg"));
    NewYearImg = await NewYearImg.blob();

    let USIndependenceDayImg: any = await fetch(require("../../../assets/img/defaultImages/US-Independence-Day.jpg"));
    USIndependenceDayImg = await USIndependenceDayImg.blob();

    let imagesArray = [
      {
        fileName: "Christmas-Day.jpg",
        fileContent: ChristmasDayImg
      },
      {
        fileName: "INDIA-Independence-Day.jpg",
        fileContent: INDIAIndependenceDayImg
      },
      {
        fileName: "New-Year.jpg",
        fileContent: NewYearImg
      },
      {
        fileName: "US-Independence-Day.jpg",
        fileContent: USIndependenceDayImg
      }
    ]

    return imagesArray;
  }

  private _commonFlowAfterSideDesignApply = async () => {
    let siteUrl = this.props.context.pageContext.legacyPageContext.webAbsoluteUrl;
    let listId = "";
    PnpSpCommonServices._getFolderByPath(this.props.context, "SiteAssets/Lists")
      .then((response) => {
        //check Lists folder in Site Assets already exists if no then create.
        if (response.status == 200) {
          return;
        }
        else {
          return PnpSpCommonServices._createFolder(this.sp, siteUrl + "/SiteAssets/Lists");
        }
      }).then((response) => {
        //get list object for get ID
        return PnpSpCommonServices._getSiteListByName(this.props.context, "Holiday Details");
      }).then(async (response) => {
        return await response.json();
      }).then((response) => {
        listId = response.d.Id;
        //create folder using list ID
        return PnpSpCommonServices._createFolder(this.sp, "" + siteUrl + "/SiteAssets/Lists/" + listId + "");
      }).then(async (response) => {
        return this._setupImageObject();
      }).then(async (response) => {
        response.forEach(async image => {
          await PnpSpCommonServices._addImage(this.sp, siteUrl + "/SiteAssets/Lists/" + listId + "", image);
        });
      }).then((response) => {
        let currentYear = new Date().getFullYear();
        let defaultData: any = [
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
            CV_Festival_Name: "Indian Independence Day",
            CV_Festival_Date: new Date("" + currentYear + "-08-15T00:00:00Z"),
            CV_FestivalDescription: "Independence Day is celebrated annually on 15 August as a public holiday in India commemorating the nation's independence from the United Kingdom on 15 August 1947",
            CV_FestivalInfoLink: {
              Description: "INDIA-Independence-Day",
              Url: "https://en.wikipedia.org/wiki/Independence_Day_(India)",
            },
            CV_FestivalImage: JSON.stringify({
              type: 'thumbnail',
              serverRelativeUrl: siteUrl + '/SiteAssets/Lists/' + listId + "/INDIA-Independence-Day.jpg"
            })
          },
          {
            Title: "Holiday-" + currentYear + "",
            CV_Festival_Name: "Christmas Day",
            CV_Festival_Date: new Date("" + currentYear + "-12-25T00:00:00Z"),
            CV_FestivalDescription: "Christmas is an annual festival commemorating the birth of Jesus Christ, observed primarily on December 25 as a religious and cultural celebration among billions of people around the world.",
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
        return this._addDefaultItemsInList(defaultData);
      }).then((response) => {
        return this._getCurrentYearHolidays();
      });
  }
}