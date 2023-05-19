import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/site-scripts";
import "@pnp/sp/site-designs";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/batching";

const PnpSpCommonServices = {
    _getSiteListByName: async (context: any, listName: string) => {
        var myHeaders = new Headers({
            'Accept': 'application/json; odata=verbose'
        });

        var myInit = {
            method: 'GET',
            headers: myHeaders,
        }

        return await fetch(context.pageContext.legacyPageContext.webAbsoluteUrl + "/_api/web/lists/getByTitle('" + listName + "')", myInit).then((response) => {
            return response;
        });
    },
    _getFolderByPath: async (context: any, folderPath: string) => {
        var myHeaders = new Headers({
            'Accept': 'application/json; odata=verbose'
        });

        var myInit = {
            method: 'GET',
            headers: myHeaders,
        }

        return await fetch(context.pageContext.legacyPageContext.webAbsoluteUrl + "/_api/web/getFolderByServerRelativeUrl('" + folderPath + "')", myInit).then((response) => {
            return response;
        });
    },
    _getSiteScript: async (sp: any) => {
        return await sp.siteScripts.getSiteScripts();
    },
    _getSiteDesign: async (sp: any) => {
        return await sp.siteDesigns.getSiteDesigns();
    },
    _createSiteScript: async (context: any, sp: any) => {

        const holidayDetailssiteScript = {
            "$schema": "https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json",
            "actions": [
                {
                    "verb": "createSiteColumnXml",
                    "schemaXml": "<Field Type=\"Text\" ID=\"{c6f2163a-f7d5-4574-9836-3a421292293f}\" Name=\"CV_Festival_Name\" DisplayName=\"Festival Name\" Required=\"TRUE\" Group=\"_CV\" StaticName=\"CV_Festival_Name\" Customization=\"\" />"
                },
                {
                    "verb": "createSiteColumnXml",
                    "schemaXml": "<Field Type=\"DateTime\" ID=\"{6e80def4-250e-4faa-ba84-ba23c4068beb}\" Name=\"CV_Festival_Date\" DisplayName=\"Festival Date\" Required=\"TRUE\" StaticName=\"CV_PGA_DateJoinGroup\" Format=\"DateOnly\" Group=\"_CV\" Customization=\"\" />"
                },
                {
                    "verb": "createSiteColumnXml",
                    "schemaXml": "<Field Type=\"Note\" ID=\"{20918695-dd3c-435e-a5f7-c459a4202655}\" Name=\"CV_FestivalDescription\" DisplayName=\"Description\" Required=\"FALSE\" NumLines=\"6\" IsolateStyles=\"TRUE\" StaticName=\"CV_FestivalDescription\" Group=\"_CV\" Customization=\"\" />"
                },
                {
                    "verb": "createSiteColumnXml",
                    "schemaXml": "<Field Type=\"Thumbnail\" ID=\"{29107248-b00a-4a4b-8144-d503beb5b697}\" Name=\"CV_FestivalImage\" DisplayName=\"Festival Image\" Required=\"TRUE\" StaticName=\"CV_FestivalImage\" Group=\"_CV\" Customization=\"\" />"
                },
                {
                    "verb": "createSiteColumnXml",
                    "schemaXml": "<Field Type=\"URL\" ID=\"{500f1aa3-d039-4c06-986d-1509d27a6166}\" Name=\"CV_FestivalInfoLink\" DisplayName=\"Festival Info Link\" Required=\"False\" Format=\"Hyperlink\" StaticName=\"CV_FestivalInfoLink\" Group=\"_CV\" Customization=\"\" />"
                },
                {
                    "verb": "createContentType",
                    "name": "CV_HolidayDetails_CT",
                    "description": "Holiday Details Content Type",
                    "id": "0x01002620f6de3ca948b28679918fe3601b4c",
                    "hidden": false,
                    "group": "_CV",
                    "subactions":
                        [
                            {
                                "verb": "addSiteColumn",
                                "internalName": "CV_Festival_Name"
                            },
                            {
                                "verb": "addSiteColumn",
                                "internalName": "CV_Festival_Date"
                            },
                            {
                                "verb": "addSiteColumn",
                                "internalName": "CV_FestivalDescription"
                            },
                            {
                                "verb": "addSiteColumn",
                                "internalName": "CV_FestivalImage"
                            },
                            {
                                "verb": "addSiteColumn",
                                "internalName": "CV_FestivalInfoLink"
                            }
                        ]
                },
                {
                    "verb": "createSPList",
                    "listName": "CV_HolidayDetails",
                    "templateType": 100,
                    "subactions": [
                        {
                            "verb": "addContentType",
                            "name": "CV_HolidayDetails_CT"
                        },
                        {
                            "verb": "setDescription",
                            "description": "This list contains holiday details."
                        },
                        {
                            "verb": "setTitle",
                            "title": "Holiday Details"
                        },
                        {
                            "verb": "addSPView",
                            "name": "All Items",
                            "viewFields": [
                                "LinkTitle",
                                "CV_Festival_Name",
                                "CV_Festival_Date",
                                "CV_FestivalDescription",
                                "CV_FestivalImage",
                                "CV_FestivalInfoLink"
                            ],
                            "query": "",
                            "rowLimit": 100,
                            "isPaged": true,
                            "makeDefault": true,
                            "replaceViewFields": true
                        }
                    ]

                }
            ],
            "bindata": {},
            "version": "1"
        }
        return await sp.siteScripts.createSiteScript("HolidayDetailsSiteScript", "HolidayDetailsSiteScript", holidayDetailssiteScript);
    },
    _createSiteDesign: async (sp: any, siteScriptId: any) => {
        return await sp.siteDesigns.createSiteDesign({
            SiteScriptIds: [siteScriptId],
            Title: "HolidayDetailsSiteDesign",
            WebTemplate: "64",
        });
    },
    _applySiteDesignToSite: async (sp: any, siteDesignId: string, siteUrl: string) => {
        return await sp.siteDesigns.applySiteDesign(siteDesignId, siteUrl);
    },
    _getListItemsWithExpandStringWithFiltersAndOrderByWithTop: async (sp: any, listName: string, selectString: string, expandString: string, filterString: string, orderByColumn: string, ascending: boolean, topCount: number) => {
        return await sp.web.lists.getByTitle(listName).items.select(selectString).expand(expandString).filter(filterString).orderBy(orderByColumn, ascending).top(topCount)();
    },
    _createFolder: async (sp: any, folderUrl: string) => {
        return await sp.web.folders.addUsingPath(folderUrl);
    },
    _addItemsUsingBatch: async (sp: any, listName: string, dataArray: any) => {
        const [batchedSP, execute] = sp.batched();
        const list = batchedSP.web.lists.getByTitle(listName);
        let res: any = [];

        dataArray.forEach((dataItems: any) => {
            list.items.add(dataItems).then((r: any) => res.push(r));
        });
        await execute();
        return res;
    },
    _addImage: async (sp: any, folderPath: string, file: any) => {
        // const [batchedSP, execute] = sp.batched();
        // // const folder = batchedSP.web.getFolderByServerRelativePath(folderPath);
        // let res: any = [];

        // files.forEach((fileItems: any) => {
        //     // folder.files.addUsingPath(fileItems.fileName, fileItems.fileContent, { Overwrite: true }).then((r: any) => res.push(r));
        //     batchedSP.web.getFolderByServerRelativePath(folderPath).files.addUsingPath(files.fileName, files.fileContent, { Overwrite: true }).then((r: any) => res.push(r));
        // });
        // await execute();
        // return res;

        return await sp.web.getFolderByServerRelativePath(folderPath).files.addUsingPath(file.fileName, file.fileContent, { Overwrite: true });
    }
}
export default PnpSpCommonServices;