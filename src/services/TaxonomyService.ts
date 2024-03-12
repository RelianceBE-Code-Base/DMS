import { ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { SPFI, spfi, SPFx } from '@pnp/sp/presets/all';
import "@pnp/sp/taxonomy";
import { IOrderedTermInfo, ITermInfo } from "@pnp/sp/taxonomy";
import "@pnp/sp/fields";
import { IFileItem } from "../model/IFileItem";
import { ITermNode } from "../model/ITermNode";

export class TaxonomyService {
    private _sp: SPFI;

    public constructor (serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
          const pageContext: PageContext = serviceScope.consume(PageContext.serviceKey);
          this._sp = spfi().using(SPFx({ pageContext }));
        });
      }
    
}