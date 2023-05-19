import * as React from 'react';
import { IHolidayDetailsProps } from './IHolidayDetailsProps';
import { spfi, SPFx } from "@pnp/sp";
import PnpSpCommonServices from '../../services/PnpSpCommonServices';
import { ChoiceGroup, IChoiceGroupOption, FocusZone, List, mergeStyleSets, ITheme, getTheme, IRectangle, TooltipHost, FontIcon, Dropdown, IDropdownOption, DefaultButton } from 'office-ui-fabric-react';
require("../../assets/stylesheets/base/global.scss");

const HolidayDetails: React.FunctionComponent<IHolidayDetailsProps> = (props) => {
    const sp = spfi().using(SPFx(props.context));
    const [currentYearHolidays, setCurrentYearHolidays] = React.useState([]);
    const [festivalDetailsListView, setFestivalDetailsListView] = React.useState([]);
    const [selectedView, setSelectedView] = React.useState("list");
    const [selectedYear, setSelectedYear]: any = React.useState("");
    const [dateFormat, setDateFormat]: any = React.useState({ day: 'numeric', month: 'long', year: 'numeric' });
    // weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'

    const options: any[] = [
        {
            key: 'list',
            iconProps: { iconName: 'BulletedList2' }
        },
        {
            key: 'card',
            iconProps: { iconName: 'SnapToGrid' }
        }
    ];

    const dropDownOptions = [
        { key: '2022', text: '2022' },
        { key: '2023', text: '2023' }
    ];
    const theme: ITheme = getTheme();
    const calloutProps = { gapSpace: 0 };

    React.useEffect(() => {
        initialFunction(new Date().getFullYear().toString());
    }, []);

    return (
        <>
            <div className="ms-Grid">
                <div className="ms-Grid-row">
                    <div className="title-with-choicegroup-box">
                        <div className="left-content">
                            <h2 className="page-title">{props.title}</h2>
                        </div>
                        <div className="right-content">
                            <Dropdown
                                label="Year"
                                selectedKey={selectedYear ? selectedYear : undefined}
                                onChange={(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) => { setSelectedYear(item.key); initialFunction(item.key); }}
                                placeholder="Select an option"
                                options={dropDownOptions}
                            />
                            <ChoiceGroup defaultSelectedKey={selectedView} className="switch-button-container" options={options} onChange={_onChangeChoiceGroup} />
                        </div>
                    </div>
                </div>
            </div>

            {/* DOM for list view */}
            {selectedView == "list" ?
                <div className="holiday-card-list">
                    {festivalDetailsListView && festivalDetailsListView.length > 0 ?
                        festivalDetailsListView.map((ele: any) => {
                            return (
                                <>
                                    <div className="holiday-card-list-item">
                                        <div className="cmn-card-shadow-box holiday-card-list-container center-content">
                                            <div className="card-header">
                                                <h2>{Object.keys(ele)[0] + " " + selectedYear}</h2>
                                            </div>
                                            <div className="card-body">
                                                <ul className="month-card-list">
                                                    {ele[Object.keys(ele)[0]].length > 0 ?
                                                        ele[Object.keys(ele)[0]].map((festival: any) => {
                                                            let tooltipId = festival.CV_Fest;
                                                            return (
                                                                <li className="month-card-list-item">
                                                                    <div className="left-content">
                                                                        <span>{new Date(festival.CV_Festival_Date).getDate()}</span>
                                                                        <span>{new Date(festival.CV_Festival_Date).toLocaleDateString('en-US', { weekday: 'short' })}</span>
                                                                    </div>
                                                                    <div className="right-content">
                                                                        <TooltipHost
                                                                            content={festival.CV_FestivalDescription}
                                                                            id={tooltipId}
                                                                            calloutProps={calloutProps}
                                                                        >
                                                                            <h3 aria-describedby={tooltipId}>{festival.CV_Festival_Name}</h3>
                                                                        </TooltipHost>

                                                                    </div>
                                                                </li>
                                                            );
                                                        })
                                                        :
                                                        <li className="top-left-center-alignment">No Holiday</li>
                                                    }
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </>
                            );
                        })
                        : ""}
                </div>
                : ""}

            {/* DOM for card view */}
            {selectedView == "card" ?
                <div className="holiday-card-list">
                    {currentYearHolidays && currentYearHolidays.length > 0 ?
                        currentYearHolidays.map((festival: any) => {
                            const imageJSON = JSON.parse(festival.CV_FestivalImage);
                            let tooltipId = festival.CV_Fest;
                            return (
                                <div className="holiday-card-list-item">
                                    <div className="cmn-card-shadow-box holiday-card-view-container">
                                        <div className="card-header">
                                            <div className="card-image">
                                                <img src={imageJSON ? imageJSON.serverRelativeUrl : ""} alt={festival.CV_Festival_Name} />
                                            </div>
                                        </div>
                                        <div className="card-body">
                                            <ul className="month-card-list">
                                                <li className="month-card-list-item">
                                                    <div className="left-content">
                                                        <span>{new Date(festival.CV_Festival_Date).getDate()}</span>
                                                        <span>{new Date(festival.CV_Festival_Date).toLocaleDateString('en-US', { weekday: 'short' })}</span>
                                                    </div>
                                                    <div className="right-content">
                                                        <span className="month">{new Date(festival.CV_Festival_Date).toLocaleString('default', { month: 'long' })}</span>
                                                        <span className="year">{new Date(festival.CV_Festival_Date).getFullYear()}</span>
                                                    </div>
                                                </li>
                                            </ul>
                                            <div className="content-box">
                                                <div className="detail-box">
                                                    <h3>{festival.CV_Festival_Name}</h3>
                                                    {festival.CV_FestivalInfoLink ?
                                                        <a className="link-icon" target='_blank' href={festival.CV_FestivalInfoLink.Url}>
                                                            <TooltipHost
                                                                content={"Click to view more"}
                                                                id={tooltipId}
                                                                calloutProps={calloutProps}
                                                            >
                                                                <FontIcon aria-label="Info" iconName="OpenInNewWindow" aria-describedby={tooltipId} />
                                                            </TooltipHost>
                                                        </a>

                                                        : ""}
                                                </div>
                                                <p>{festival.CV_FestivalDescription ? festival.CV_FestivalDescription : ""}</p>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            );
                        })
                        :
                        ""}
                </div>
                : ""}

            {/* DOM for data not available */}
            {(currentYearHolidays && currentYearHolidays.length == 0) && (currentYearHolidays && currentYearHolidays.length == 0) ?
                <div className="not-found-message-content-box">
                    <div className="content-box">
                        <img src={require("../../assets/img/icons/DataNotFound.jpg")} alt="Not available now" />
                        <h3>No data found!</h3>
                        <p>You can click on below button to add data.</p>
                        <div className="btn-container btn-center">
                            <DefaultButton text='Add data' className="btn-primary-1" onClick={() => { location.href = props.context.pageContext.web.absoluteUrl + "/Lists/CV_HolidayDetails/AllItems.aspx" }} />
                        </div>
                    </div>
                </div>
                : ""}
        </>
    );

    /**
     * Function for get current year holidays.
     * @returns 
     */
    async function _getCurrentYearHolidays(year: any): Promise<any> {

        let selectString = "Title,CV_Festival_Name,CV_Festival_Date,CV_FestivalDescription,CV_FestivalImage,CV_FestivalInfoLink";
        let expandString = "";
        let filterString = "(CV_Festival_Date eq '" + year + "-01-01T00:00:00Z' or CV_Festival_Date gt '" + year + "-01-01T00:00:00Z') and (CV_Festival_Date eq '" + year + "-12-31T23:59:59Z' or CV_Festival_Date lt '" + year + "-12-31T23:59:59Z')";

        return new Promise((resolve, reject) => {
            PnpSpCommonServices._getListItemsWithExpandStringWithFiltersAndOrderByWithTop(sp, "Holiday Details", selectString, expandString, filterString, "Id", true, 4999).then((response) => {
                resolve(response);
            },
                (error: any) => {
                    console.log(error);
                    reject(error);
                });
        });
    }
    /**
     * Function for get choice button value
     * @param ev 
     * @param option 
     */
    function _onChangeChoiceGroup(ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): void {
        setSelectedView(option.key)
    }

    function initialFunction(selectedYear: any) {
        setSelectedYear(selectedYear)
        _getCurrentYearHolidays(selectedYear).then((currentYearHolidayResponse) => {
            setCurrentYearHolidays(currentYearHolidayResponse);
            let festivalSeparateMonth: any = [];
            if (currentYearHolidayResponse.length > 0) {
                //separate all the holidays by month
                festivalSeparateMonth = [
                    { January: currentYearHolidayResponse.filter((ele: any) => new Date(ele.CV_Festival_Date).getMonth() == 0) },
                    { February: currentYearHolidayResponse.filter((ele: any) => new Date(ele.CV_Festival_Date).getMonth() == 1) },
                    { March: currentYearHolidayResponse.filter((ele: any) => new Date(ele.CV_Festival_Date).getMonth() == 2) },
                    { April: currentYearHolidayResponse.filter((ele: any) => new Date(ele.CV_Festival_Date).getMonth() == 3) },
                    { May: currentYearHolidayResponse.filter((ele: any) => new Date(ele.CV_Festival_Date).getMonth() == 4) },
                    { June: currentYearHolidayResponse.filter((ele: any) => new Date(ele.CV_Festival_Date).getMonth() == 5) },
                    { July: currentYearHolidayResponse.filter((ele: any) => new Date(ele.CV_Festival_Date).getMonth() == 6) },
                    { August: currentYearHolidayResponse.filter((ele: any) => new Date(ele.CV_Festival_Date).getMonth() == 7) },
                    { September: currentYearHolidayResponse.filter((ele: any) => new Date(ele.CV_Festival_Date).getMonth() == 8) },
                    { October: currentYearHolidayResponse.filter((ele: any) => new Date(ele.CV_Festival_Date).getMonth() == 9) },
                    { November: currentYearHolidayResponse.filter((ele: any) => new Date(ele.CV_Festival_Date).getMonth() == 10) },
                    { December: currentYearHolidayResponse.filter((ele: any) => new Date(ele.CV_Festival_Date).getMonth() == 11) }
                ];
            }
            setFestivalDetailsListView(festivalSeparateMonth);
        });
    }
};

export default HolidayDetails;
