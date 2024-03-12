import { ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/files";
import "@pnp/sp/folders";


import { IFileItem } from "../model/IFileItem";
import { IItem } from "@pnp/sp/items";

export class SPService {
  private _listName: string;
  private _fieldName: string;
  private _sp: SPFI;

  public constructor (serviceScope: ServiceScope, listname: string, fieldname: string) {
    this._listName = listname;
    this._fieldName = fieldname;
    serviceScope.whenFinished(() => {
      const pageContext: PageContext = serviceScope.consume(PageContext.serviceKey);
      this._sp = spfi().using(SPFx({ pageContext }));
    });
  }

  

}