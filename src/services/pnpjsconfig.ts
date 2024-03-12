import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp/presets/all";
import { LogLevel, PnPLogging } from "@pnp/logging";

//import { IDmsProps } from "../webparts/dms/components/IDmsProps";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";



let _sp: SPFI = null;

export const getSP: any = (context?: WebPartContext): SPFI => {

    
    if (_sp === null && context != null) {
      //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
      // The LogLevel set's at what level a message will be written to the console
      _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
    }
    return _sp;
  };