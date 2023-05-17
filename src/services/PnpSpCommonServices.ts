import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";

const PnpSpCommonServices = {
    _getSiteLists: async (sp: any) => {
        return await sp.web.lists();
    },
    _getListItemsWithExpandStringWithFiltersAndOrderByWithTop: async (sp: any, listName: string, selectString: string, expandString: string, filterString: string, orderByColumn: string, ascending: boolean, topCount: number) => {
        return await sp.web.lists.getByTitle(listName).items.select(selectString).expand(expandString).filter(filterString).orderBy(orderByColumn, ascending).top(topCount)();
    }
}
export default PnpSpCommonServices;