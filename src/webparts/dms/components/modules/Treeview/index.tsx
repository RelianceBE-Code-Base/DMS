import * as React from "react";
import { SPFI } from '@pnp/sp/presets/all';

import { IDmsProps } from "../../IDmsProps";


import { escape } from '@microsoft/sp-lodash-subset';

import * as jQuery from "jquery";


import { IDocumentLibraryInformation } from "@pnp/sp/sites";
import "@pnp/sp/sites";
import { useEffect, useState } from 'react';

//import { getSP } from "../../pnpjscon";

import { getSP } from "../../../../../services/pnpjsconfig";



import "@pnp/sp/webs";
import "@pnp/sp/folders";
import { IFileInfo, IFolderInfo, Web } from "@pnp/sp/presets/all";
import { folderFromAbsolutePath } from "@pnp/sp/folders";
import { folderFromPath } from "@pnp/sp/folders";
import { folderFromServerRelativePath } from "@pnp/sp/folders";
import "@pnp/sp/webs";
import "@pnp/sp/presets/all";
import { DetailsList } from 'office-ui-fabric-react';
import { WebPartContext } from "@microsoft/sp-webpart-base";






const TreeView  = (props:IDmsProps) =>{

    let _sp:SPFI = getSP(props.context);

  // const [docItems, setdocItems] = useState<IEDMS[]>([])
  
  console.log('context',_sp)
  const url = `${props.siteUrl}`;
  //let subFolderArray = [];

console.log (_sp);
console.log(url)

 
      
  return(
    <>
        <h4>hello</h4> {props.userDisplayName}
    </>
    )

  
}



export default TreeView;