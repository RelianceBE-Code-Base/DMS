import * as React from 'react';
import styles from './Dms.module.scss';
import { IDmsProps } from './IDmsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as jQuery from "jquery";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { useEffect, useState } from 'react';
import { IStackProps, Stack } from 'office-ui-fabric-react/lib/Stack';
import { Spinner, SpinnerSize, SpinnerLabelPosition, TextField, Dropdown, IDropdownStyles, IDropdownOption, DropdownMenuItemType } from 'office-ui-fabric-react';
import { Panel } from '@fluentui/react/lib/Panel';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IDocumentLibraryInformation, Item, SPFI, SPFx, spfi, ICamlQuery } from '@pnp/sp/presets/all';

import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import { getTheme, mergeStyleSets } from "office-ui-fabric-react/lib/Styling";
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
//import Link from '@material-ui/core/Link';
import { getSP } from '../../../services/pnpjsconfig';

import {
  FontWeights,
  ContextualMenu,
  Toggle,
  IDragOptions,
} from '@fluentui/react';

import { Nav, INavLink } from '@fluentui/react/lib/Nav'
import { initializeIcons } from "@fluentui/react";
import { IconButton, PrimaryButton, DefaultButton, IButtonStyles } from '@fluentui/react/lib/Button';

import { Icon, IIconProps } from 'office-ui-fabric-react/lib/Icon';

import { IColumn, DetailsList,DetailsListLayoutMode, Selection, SelectionMode} from '@fluentui/react/lib/DetailsList';
import { ShimmeredDetailsList } from 'office-ui-fabric-react/lib/ShimmeredDetailsList';

import { CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react/lib/CommandBar';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react';



import { useBoolean } from '@fluentui/react-hooks';

//import folderIcon from '../components/icons/folder.svg';

//const folderIcon = require('../components/icons/folder.svg') 


//"data:image/png;base64,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";


const theme = getTheme();

const contentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    flexFlow: 'column nowrap',
    alignItems: 'stretch',
    
  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      borderTop: `4px solid ${theme.palette.themePrimary}`,
      color: theme.palette.neutralPrimary,
      display: 'flex',
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      padding: '12px 12px 14px 24px',
    },
  ],
  heading: {
    color: theme.palette.neutralPrimary,
    fontWeight: FontWeights.semibold,
    fontSize: 'inherit',
    margin: '0',
  },
  body: {
    flex: '4 4 auto',
    padding: '0 24px 24px 24px',
    overflowY: 'hidden',
    selectors: {
      p: { margin: '14px 0' },
      'p:first-child': { marginTop: 0 },
      'p:last-child': { marginBottom: 0 },
    },
  },
});
const stackProps: Partial<IStackProps> = {
  horizontal: true,
  tokens: { childrenGap: 40 },
  styles: { root: { marginBottom: 20 } },
};
const iconButtonStyles: Partial<IButtonStyles> = {
  root: {
    color: theme.palette.neutralPrimary,
    marginLeft: 'auto',
    marginTop: '4px',
    marginRight: '2px',
  },
  rootHovered: {
    color: theme.palette.neutralDark,
  },
};


const cancelIcon: IIconProps = { iconName: 'Cancel' };

// Details list style
const margin = '0 30px 20px 0';

const classNames = mergeStyleSets({
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: "16px"
  },
  fileIconCell: {
    textAlign: "center",
    selectors: {
      "&:before": {
        content: ".",
        display: "inline-block",
        verticalAlign: "middle",
        height: "100%",
        width: "0px",
        visibility: "hidden"
      }
    }
  },
  fileIconImg: {
    verticalAlign: "middle",
    maxHeight: "27px",
    maxWidth: "16px"
  },
  controlWrapper: {
    display: "flex",
    flexWrap: "wrap"
  },
  
});
const controlStyles = {
  root: {
    margin: "0 30px 20px 0",
    maxWidth: "300px"
  }
};



const detailsStyles = {
  root: {
    margin: "1% 1%",
    textDecoration: 'none',
    width: '78%',
    
    
  }
};



const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },
};


interface IModalExampleState {
  isModalHidden: boolean;
  selectedItem: any; // Adjust the type according to your requirements
}


// export default class Dms extends React.Component<IDmsProps, {}> {
//   public render(): React.ReactElement<IDmsProps> {

    // const {
    //   userDisplayName,
    //   context,
    //   siteUrl,
    //   folderUrl  
    // } = this.props;

    


  interface DocArrayObj {
      file?: {mimeType:string};
      folder: boolean;
      id: string;
      name: string;
      webUrl: string;
      parentId?: string;
      // folder?: {childCount: number};
      parentReference?:{driveId:string, driveType: string, id: string, path:string, siteId:string} ////filter by drivetype where value is "documentLibrary"
      //children?: DocArrayObj[]; 
      Approver: string
      //fileType?: string; // Assuming you have a fileType property in your data
      dateModifiedValue?: string; // Assuming you have a dateModifiedValue property in your data
      modifiedBy?: string; // Assuming you have a modifiedBy property in your data
      fileSizeRaw?: number; // Assuming you have a fileSizeRaw property in your data
      lastModifiedDateTime? : string;
      lastModifiedBy?: {user?:{displayName:string, email:string, id:string}};
      size:number;
      listItem?: {fields?:{Approver:string, Narration:string, DocIcon:string, Tag: string, description: string}, id:string}
    }

  export interface IDocItems {
    odata: string;
    value: DocArrayObj[];
    id: string;
    name: string;
}
    
// spinner 
const rowProps: IStackProps = { horizontal: true, verticalAlign: 'center' };
const tokens = {
    sectionStack: {
        childrenGap: 10,
    },
    spinnerStack: {
        childrenGap: 10,
    },
};

//nav style
const navigationStyles = {
  root: {
    height: "100vh",
    width: "100%",
    boxSizing: "border-box",
    border: "1px solid #eee",
    overflowY: "auto",
    float: "left",
    margin: "0",
    fontSize: "5px"
    
    //paddingTop: '10vh',
  },
};


// export default class Dms extends React.Component<IDmsProps, {}> {
// public render(): React.ReactElement<IDmsProps> {
const Dms = (props:IDmsProps) => {


  //command bar Items
const _commandBarItems: ICommandBarItemProps[] = [
  // {
  //   key: 'newItem',
  //   text: "New",
  //   iconProps: { iconName: 'Add' },
  //   onClick: () => console.log('newItem')
  // },

  // {
  //   key: 'newItem',
  //   text: "Refresh",
  //   iconProps: { iconName: 'refresh' },
  //   onClick: () => console.log('newItem')
  // },

  {
    key: 'uploadItem',
    text: "Upload",
    iconProps: { iconName: 'Upload' },

    onClick: () => {showModal()}
    

    //onClick: () => console.log('uploadItem')
  },
  // {
  //   key: 'downloadItem',
  //   text: "download",
  //   iconProps: { iconName: 'download' },
  //   onClick: () => {}
  // },
  // {
  //   key: 'delete',
  //   text: "Delete",
  //   iconProps: { iconName: 'Delete' },
  //   onClick: () => {_deleteBodyModal()}
  // },
  // {
  //   key: 'submit',
  //   text: "Submit",
  //   iconProps: { iconName: 'Send' },
  //   onClick: () => console.log('submit'),
  // }
];

// const _farItems: ICommandBarItemProps[] = [
//   {
//     key: 'info',
//     text: 'Info',
//     ariaLabel: 'Info',
//     iconOnly: true,
//     iconProps: { iconName: 'Info' },
//     onClick: () => openPanel(),
//   },
// ];


/*********************************/

const trimFileNameByExtension = (fileName: string) => {
  return fileName.replace(/\.[^/.]+$/, '');
};

//Detail List Items
const _detailsListColumns: IColumn[] = [
  {
    key: "column1",
    name: "File Type",
    className: classNames.fileIconCell,
    iconClassName: classNames.fileIconHeaderIcon,
    ariaLabel:"Column operations for File type, Press to sort on File type",
    iconName: "Page",
    isIconOnly: true,
    fieldName: "name",
    minWidth: 16,
    maxWidth: 16,
    // onColumnClick: this._onColumnClick,
    onRender: (item: DocArrayObj) => {
      

      let iconurl:string ;
      if(item.file?.mimeType.toLowerCase() === "video/mp4")
{
 iconurl ='https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/video.png';
}
else if(item.file?.mimeType.toLowerCase() === "image/png")
{
 iconurl ='https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/photo.png';
}
else if(item.file?.mimeType.toLowerCase() === "image/jpg")
{
 iconurl ='https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/photo.png';
}
else if(item.file?.mimeType.toLowerCase() === "image/jpeg")
{
 iconurl ='https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/photo.png';
}

else if(item.file?.mimeType.toLowerCase() === "application/pdf")
{
 iconurl ='https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/pdf.png';
}
else if(item.file?.mimeType.toLowerCase() === "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
{
 iconurl ='https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/docx.png';
}

else if(item.file?.mimeType.toLowerCase() === "text/plain")
{
 iconurl ='https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/txt.png';
}

else if(item.file?.mimeType.toLowerCase() === "application/vnd.ms-excel")
{
 iconurl ='https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/xlsx.png';
}

else if(item.file?.mimeType.toLowerCase() === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
{
 iconurl ='https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/xlsx.png';
}

else{
 iconurl =`https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/${item.listItem?.fields?.DocIcon}_16x1.svg`;
}

      return (
        <img
          src={iconurl}
          className={classNames.fileIconImg}
          //alt={item.name + " file icon"}
        />
      );
    },
  },
  {
    key: "column2",
    name: "Name",
    fieldName: "name",
    minWidth: 210,
    maxWidth: 350,
    isRowHeader: true,
    isResizable: true,
    isSorted: true,
    isSortedDescending: false,
    sortAscendingAriaLabel: "Sorted A to Z",
    sortDescendingAriaLabel: "Sorted Z to A",
    //onColumnClick: this._onColumnClick,
    onRender: (item: DocArrayObj) => (
      <a href={item.webUrl} target="_blank" rel="noopener noreferrer" style={{ textDecoration:'none'}}>
        {trimFileNameByExtension(item.name)} 
      </a>
    ),
    data: "string",
    isPadded: true
  },
  {
    key: "column3",
    name: "Date Modified",
    fieldName: "dateModifiedValue",
    minWidth: 70,
    maxWidth: 90,
    isResizable: true,
    // onColumnClick: this._onColumnClick,
    data: "number",
    onRender: (item: DocArrayObj) => {
      return <span>{item.lastModifiedDateTime}</span>;
    },
    isPadded: true
  },
  // {
  //   key: "column3",
  //   name: "Narration",
  //   fieldName: "Narration",
  //   minWidth: 70,
  //   maxWidth: 90,
  //   isResizable: true,
  //   // onColumnClick: this._onColumnClick,
  //   data: "string",
  //   onRender: (item: DocArrayObj) => {
  //     return <span>{item.listItem.fields.Narration}</span>;
  //   },
  //   isPadded: true
  // },
  {
    key: "column4",
    name: "Modified By",
    fieldName: "modifiedBy",
    minWidth: 70,
    maxWidth: 90,
    isResizable: true,
    isCollapsible: true,
    data: "string",
    // onColumnClick: this._onColumnClick,
    onRender: (item: DocArrayObj) => {
      return <span>{item.lastModifiedBy?.user?.displayName}</span>;
    },
    isPadded: true
  },
  {
    key: "column5",
    name: "File Size",
    fieldName: "fileSizeRaw",
    minWidth: 70,
    maxWidth: 90,
    isResizable: true,
    isCollapsible: true,
    data: "number",
    // onColumnClick: this._onColumnClick,
    onRender: (item: DocArrayObj) => {
      return <span>{item.size}</span>;
    }
  }
];

/************************************* */
    
    jQuery("#workbenchPageContent").prop("style", "max-width: none");
    jQuery(".SPCanvas-canvas").prop("style", "max-width: none");
    jQuery(".CanvasZone").prop("style", "max-width: none");
    
  
    let _sp:SPFI = getSP(props.context);

     
    


    
// Extracting root URL
    let webAddress = props.siteUrl;
    let rootUrl = webAddress.split('/sites')[0].replace(/^https?:\/\//, ''); //replace https:// and get only items before sites
  //Rekative path   
    let relativePath = props.folderUrl;
    let newRelativePath = relativePath.substring(1);

    // alert("Root")
    // alert(rootUrl);
    // alert("Server")
    // alert(newRelativePath)


    const [reload, setReload] = useState<boolean>(false);
    const [docfolderItem, setdocfolderItem] = useState<Array<DocArrayObj>>([]);
    const [filesItem, setfilesItem] = useState<DocArrayObj[]>([]);

    let [newfilesItem, setnewfilesItem] = useState([]);

    const [newfilesID, setnewfilesID] = useState("");
    const [refectch, setRefetch] = useState("");

    const [isLoadingMenu, setIsLoadingMenu] = useState(true);
    const [isLoadingUpload, setIsLoadingUpload] = useState(true);
    const [isDialogVisible, setIsDialogVisible] = useState(false);
    const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] = useBoolean(false);
    const [isModalOpen, setIsModalOpen] = useState(false);

    const [dropdownOptions, setDropdownOptions] = useState<IDropdownOption[]>([]);
    const [selectedOption, setSelectedOption] = useState({text: ""} as any);

    const [selection, setSelection] = useState(new Selection());
    
    const [filesUpload, setfilesUpload] = useState({} as any)

    const [Approver, setApprover] = useState("")
    const [Narration, setNarration] = useState("")
    const _deleteBodyModal = () => {
      // Implement the body content for your modal
      return (
        <div>
          <p>Are you sure you want to delete this file</p>
        </div>
      );
    };


    //const [deltaItems, setDeltaItems] = useState<Array<any>>([]); 

    //https://graph.microsoft.com/v1.0/sites/mysite.sharepoint.com:/sites/MyDocumentSite:/drives

  //  https://graph.microsoft.com/v1.0/sites/3zv7tj.sharepoint.com:/sites/DynamicCCTV:/drives

    //let resTrimmedID: string;

    //get site root folders
    
//rootUrl

//newRelativePath


    const getSiteItem = async () => {
    props.context.msGraphClientFactory  
    .getClient('3')
      .then((client: MSGraphClientV3) => {  
        client
          .api(newRelativePath == "" ? `sites/root/drives` : `sites/${rootUrl}:/${newRelativePath}:/drives`) //sites  //sites/${folderSite}:/   //sites/{host-name}:/{server-relative-path}
          .version("v1.0")  //id, name, webUrl
          .get(async (err, res:IDocItems) => {
            if (err) {
              console.error("MSGraphAPI Error")
              console.error(err);
              return;
            }
            //console.log("Response:",res)
            setdocfolderItem(res.value);
            setIsLoadingMenu(false);

            // Call the new function for each drive
            res.value.forEach((drive) => {
             getSubRootDocumentItems(client, drive.id);                         
           });       
          });
      });
    }

    //get all items(folders within root folder)
    const getSubRootDocumentItems  = async (client: MSGraphClientV3, driveId: string,) => {  //parentId?: string
      try {
        const deltaResponse = await client.api(`sites/root/drives/${driveId}/root/delta`).version('v1.0').select('*').get();  //
        //console.log(`deltaResponse for ${driveId}`, deltaResponse)
        // Filter out root and non-folder items
        const foldersOnly = deltaResponse.value.filter((item: any) =>  item.folder && item.name.toLowerCase() !== 'root');
        //console.log(foldersOnly)

        //filter out folder items
        const filesOnly = deltaResponse.value.filter((item: any) => item.file);

        
        //console.log(filesOnly)
      // const filterFiles = (items: DocArrayObj[]) => {return items.filter(item => item.file && item.file.toLowerCase() !== 'folder');};


       //console.log(`Delta Response for Drive ${driveId}:`, foldersOnly);

       //Document folder listItems : to get custom columns
       const listItemResponse = await client.api(`drives/${driveId}/root/children?expand=listItem`)
      .version('v1.0')
      .select('*')
      .get();
      console.log(`Child Response for ${driveId} ` ,listItemResponse)


      


       //update folder item states
       setdocfolderItem((prevDeltaItems) => [
          ...prevDeltaItems,
          ...foldersOnly.map((item: any) => ({
            ...item,
            parentId: item.parentReference ? item.parentReference.driveId : driveId,
          })),
        ]);

        


        // Update fileItems state
        setfilesItem((prevFileItems) => [
        ...prevFileItems,
        ...filesOnly?.map((item: any) => ({
          ...item,
          parentId: item.parentReference ? item.parentReference.driveId : driveId,
        })),
      ]);
        

        // Process deltaResponse as needed
      } catch (error) {
        console.error(`Error fetching delta items for Drive ${driveId}:`, error);
      }
    };
   

    // const getUpdatedSiteItem = async () => {
    //   props.context.msGraphClientFactory  
    //   .getClient('3')
    //     .then((client: MSGraphClientV3) => {
    //       client
    //         .api(`sites/root/drives`) //sites/${folderSite}:/ 
    //         .version("v1.0")  //id, name, webUrl
    //         .get(async (err, res:IDocItems) => {
    //           if (err) {
    //             console.error("MSGraphAPI Error")
    //             console.error(err);
    //             return;
    //           }
        
    //           setdocfolderItem(res.value);
    //           setIsLoadingMenu(false);
  
    //                   res.value.forEach((drive) => {
    //            refetchSubRootDocumentItems(client, drive.id);            
          
               
    //          });       
    //         });
    //     });
    //   }
  
    //   // console.log(docfolderItem, "logged")
    //   console.log(filesItem, "filesOnly")

    //   const refetchSubRootDocumentItems  = async (client: MSGraphClientV3, driveId: string,) => {  //parentId?: string
    //       try {
    //         const deltaResponse = await client.api(`sites/root/drives/${driveId}/root/delta`).version('v1.0').select('*').get();  //
    //         const foldersOnly = deltaResponse.value.filter((item: any) =>  item.folder && item.name.toLowerCase() !== 'root');

    //         const filesOnly = deltaResponse.value.filter((item: any) => item.file);
    //        const listItemResponse = await client.api(`drives/${driveId}/root/children?expand=listItem`)
    //       .version('v1.0')
    //       .select('*')
    //       .get();

    //       console.log(filesOnly, "filesOnly")
          
    //                   setfilesItem((prevFileItems) => [
    //         ...filesOnly?.map((item: any) => ({
    //           ...item,
    //           parentId: item.parentReference ? item.parentReference.driveId : driveId,
    //         })),
    //       ]);
    //       setRefetch(newfilesID)
    //       setnewfilesID(newfilesID)

    //         } catch (error) {
    //         console.error(`Error fetching delta items for Drive ${driveId}:`, error);
    //       }
    //     };
  

 



    //generate treeview
    const generateTreeviewData = (items: DocArrayObj[]) => {
      const treeData: any[] = [];
      items.forEach((item) => {
        if (!item.parentId) {
          treeData.push({
            name: item.name,
            //url: item.webUrl,
            key: item.id,
            links: generateChildLinks(item.id),
            // iconProps: {
            //   iconName: 'FolderHorizontal',
            //   // styles: {
            //   //   root: {
            //   //     fontSize: 15,
            //   //     color: '#106ebe',
            //   //     marginLeft: '20px'
            //   //   },
            //   // }
            // }
          });
        }
      });
  
      return treeData;
    };



    //generate child links
    const generateChildLinks = (parentId: string) => {
      const childLinks: any[] = [];
  
      docfolderItem.forEach((item) => {
        if (item.folder && item.parentId === parentId && item?.parentReference?.path &&
          item.parentReference.path.endsWith("/root:") &&
          item.parentReference.path.split("/").length === 4) { //parentReference.id
          childLinks.push({
            name: item.name,
            //url: item.webUrl,
            key: item.id,
            links: generateSubChildLinks(item.id),
            // iconProps: {
            //   iconName: 'FolderHorizontal',
            //   // styles: {
            //   //   root: {
            //   //     fontSize: 15,
            //   //     color: '#106ebe',
            //   //     marginLeft: '20px'
            //   //   },
            //   // }
            // }
          });
        }
      });
  
      return childLinks;
    };


        //generate child links
        const generateSubChildLinks = (parentId: string) => {
          const childLinks: any[] = [];
      
          docfolderItem.forEach((item) => {
            if (item.folder && item.parentReference.id === parentId) { //
              childLinks.push({
                name: item.name,
                //url: item.webUrl,
                key: item.id,
                links: generateChildLinks(item.id),
                // iconProps: {
                //   iconName: 'FolderHorizontal',
                //   // styles: {
                //   //   root: {
                //   //     fontSize: 15,
                //   //     color: '#106ebe',
                //   //     marginLeft: '20px'
                //   //   },
                //   // }
                // }    
              });
            }
          });
      
          return childLinks;
        };
    
    const navLinks = generateTreeviewData(docfolderItem);
    //console.log("Treeview Items", docfolderItem);

    newfilesItem = React.useMemo(() => {
      if(newfilesID || refectch){
        const filteredFiles = filesItem.filter((file) => file.parentId === newfilesID || file.parentReference.id === newfilesID);
        // console.log("Filtered Files", filteredFiles)
        // console.log("Files ID", newfilesID)
        return filteredFiles
      }
      else{
        return []
      }

    },[newfilesID, refectch])


    // Handler for folder link click
  const onFolderLinkClick = (_ev?: React.MouseEvent<HTMLElement>, item?: INavLink) => {
    if (item && item.key) {
      const selectedFolderId = item.key;
      setnewfilesID(selectedFolderId)
      
      // Filter files based on the selected folder
      // const filteredFiles = filesItem.filter((file) => file.parentId === selectedFolderId || file.parentReference.id === selectedFolderId);
      
      // Update the DetailsList with the filtered files
      // setfilesItem(filteredFiles);
    }
  };



  //console.log("FileItems",filesItem)


//Show Dialog
    const showDialog = () => {
      setIsDialogVisible(true);
    };
  
//Hide Dialog
    const hideDialog = () => {
      setIsDialogVisible(false);
    };
  
//Show Dialog
const showModal = () => {
  setIsModalOpen(true);
  setIsLoadingUpload(false);
};

//Hide Dialog
const hideModal = () => {
  setIsModalOpen(false);
};

    

    useEffect(() => {
      initializeIcons();
      getSiteItem()
    }, [reload, props.context]);  

    // Update the dropdown options whenever docfolderItem changes
  useEffect(() => {
    // Update the dropdown options
    const driveOptions: IDropdownOption[] = docfolderItem.map((drive: DocArrayObj) => ({
      key: drive.id,
      text: drive.name,
      //Approver: drive.listItem.fields.Approver,
      
    }));

    setDropdownOptions(driveOptions);
    console.log("Drive Option",driveOptions);
    setIsLoadingMenu(false);
  }, [docfolderItem]);





  //Brings out other fields when on selected documents library is selected
  const handleDropdownChange = (_event: any, option: any) => {
      // Update the selected option
    setSelectedOption(option);
    console.log("Selected",selectedOption)

  };


  const handleApprover = (event: any) => {
  // Update the selected option
  setApprover(event.target.value)
  };

const handleNarration = (event: any) => {
setNarration(event.target.value)
};


const handleUploadChange = (event: any) => {
    // Update the selected option
  setfilesUpload(event.target.files[0])

};

// const testUpload = ()=>{
//   console.log("Test")
// }


  const uploadFile = async ()=>{    
    setIsLoadingUpload(true);
    try {
      const response = await _sp.web
        .getFolderByServerRelativePath(selectedOption?.text)
        .files.addUsingPath(`${filesUpload.name}`, filesUpload, {Overwrite: true});

    const fileItem = await response.file.getItem();
    await fileItem.update({
      //Approver: Approver,
      Narration: Narration
      
    });
    console.log('Approver Column Updated',Approver)
    
    
      const data = await response;
      
      return data;
    } catch (err) {
      console.error("Error uploading photo:", err);
      return { status: "rejected", reason: err };
    }
    finally {
      
         // Fetch the updated list of items after upload and update state
    // const updatedFiles = await getUpdatedFilesList(selectedOption?.text); // Pass selected library title as a parameter
    // setnewfilesItem(newfilesItem);
    
      //getUpdatedSiteItem()
    // console.log("Updated fields",updatedFiles);
      setIsLoadingUpload(false);
      setReload(!reload);
      hideModal()
      
    }
    
  }
    
  const deleteDocItem = async (id: string) => {
    // Get a reference to the SharePoint list named "Quotes"
    const list = _sp.web.lists.getByTitle(selectedOption?.text);
    try {
      // Delete the list item with the specified ID
      await list.items.getItemByStringId(id).delete();
      
      setReload(!reload);
      // Log a message to indicate that the list item has been successfully deleted
      console.log('List item deleted');
    } catch (e) {
      // Log any errors that occur during the deletion process
      console.log(e);
    }
  }



  // Function to fetch the updated list of items
const getUpdatedFilesList = async (libraryTitle: string) => {
  try {
    // Implement the logic to fetch the updated list of items from SharePoint
    // For example, you can use _sp.web.lists.getByTitle('YourListTitle').items.getAll()
    // Make sure to get the necessary fields for your application
    const updatedList = await _sp.web.lists.getByTitle(libraryTitle).items();
    
    return updatedList;
  } catch (error) {
    console.error("Error fetching updated list:", error);
    throw error;
  }
};
    


/**
 *   // Function to fetch the updated list of items
  const getUpdatedFilesList = async (libraryTitle: string,) => { 
    try {
      
      const client = await props.context.msGraphClientFactory.getClient('3');
      const deltaResponse = await client.api(`sites/root/drives/${libraryTitle}/root/delta`).version('v1.0').select('*').get();
      
      const updatedList = deltaResponse.value.filter((item: any) => item.file);
  
      //const updatedList = await _sp.web.lists.getByTitle(libraryTitle).items();
      
      return updatedList;
    } catch (error) {
      console.error("Error fetching updated list:", error);
      throw error;
    }
  };
      
  


 */



  return (
      <div className={styles.dms}>
      {
        isLoadingMenu ? (
          <div style={{ height: '6rem', display: 'flex', justifyContent: 'center', alignContent: 'center', width: '100%' }}>
          <Spinner  size={SpinnerSize.large} label='Loading...' />
          </div>
        ) : (
          <>
          {/* Treview */}
      <div className={styles.navigation}>
      
        <Nav
          onLinkClick={onFolderLinkClick}//_onLinkClick
          styles={navigationStyles}
          groups={[{ links: navLinks }]}
          onRenderLink={(props, defaultRender) => {
            if (props.name !== "" ) {
          return <><span className={styles.folderIcon}></span>
            {/* <img src={folderIcon} width={20} height={20} /> */}
          <span className={styles.folderName}>{props.name}</span></>
            } 
return defaultRender(props)
}}
        />  
      </div>

      {/* Detailed List */}
    <div className={styles.details}>
          <div>
            <CommandBar
                items={_commandBarItems}
                ariaLabel="Use left and right arrow keys to navigate between commands"
                //farItems={_farItems}
              />
              <hr/>
          </div>
          <DetailsList
            items={newfilesItem}  //filterFiles(files) 
            columns={_detailsListColumns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}  
            isHeaderVisible={true}
            selectionPreservedOnEmptyClick={true}
            selection={selection}
            enterModalSelectionOnTouch={true}
            onItemInvoked={null}
            styles={detailsStyles}
            
            //onRenderItemColumn={_detailsListColumns.}
            selectionMode={SelectionMode.multiple}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          />
          
          <Panel
        headerText="Details"
        isOpen={isOpen}
        onDismiss={dismissPanel}
        // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
        closeButtonAriaLabel="Close"
        >
        {/* <p>Details</p> */}
      </Panel>

      <Modal
        titleAriaId="id"
        isOpen={isModalOpen}
        onDismiss={hideModal}
        isBlocking={true}
      >

        <div className={contentStyles.header}>
          <h2 className={contentStyles.heading}>
            Upload Documents
          </h2>
          <IconButton
            styles={iconButtonStyles}
            iconProps = {cancelIcon}
            ariaLabel="Close popup modal"
            onClick={hideModal}
          />
        </div>
        
        <div className={contentStyles.body}>
        <p>
        <Dropdown
        placeholder="Select a document library"
        label="Document libraries"
        options={dropdownOptions}
        styles={dropdownStyles}
        onChange={handleDropdownChange}
        disabled={isLoadingUpload}
      />
      {selectedOption && (
      <>
      <TextField label="Narration" onChange={handleNarration} disabled={isLoadingUpload}/>
      <TextField label="Upload" type="file"  onChange={handleUploadChange} disabled={isLoadingUpload}/>
      </>
      )}
        <br />
       <PrimaryButton text={isLoadingUpload ? 'Saving...' : 'Save'} onClick={uploadFile} disabled={isLoadingUpload}/>
       {
        isLoadingUpload ? (<Spinner size={SpinnerSize.small} label='Uploading...'/>):
        (<div></div>)
         }
          </p>
        </div>
      </Modal>
    </div>
        </>
      )
    }
</div>   

            
    );
  };



  
export default Dms;



/**
 * export enum BrandIcons {
  Word = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/docx.png",
  PowerPoint = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/pptx.png",
  Excel = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/xlsx.png",
  Pdf = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/pdf.png",
  OneNote = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/onetoc.png",
  OneNotePage = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/one.png",
  InfoPath = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/xsn.png",
  Visio = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/vsdx.png",
  Publisher = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/pub.png",
  Project = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/mpp.png",
  Access = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/accdb.png",
  Mail = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/email.png",
  Csv = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/xlsx.png",
  Archive = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/zip.png",
  Xps = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/genericfile.png",
  Audio = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/audio.png",
  Video = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/video.png",
  Image = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/photo.png",
  Text = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/txt.png",
  Xml = "https://spoprod-a.akamaihd.net/files/fabric/assets/item-types/32/xml.png"
}
 */








































//let itemSite: string = `${props.siteUrl}`
    //console.log ('site', itemSite)  //https://3zv7tj.sharepoint.com/

    //let folderSite: string = itemSite.replace("https://", "")
    //console.log ('folderSite:', folderSite) 













   //   <HashRouter>
    //   <Switch>
    //     <Route component={HomeScreen} path="/" exact/>
    //   </Switch>
    // </HashRouter>



    //https://graph.microsoft.com/v1.0/sites/root/drive/root/search(q='')
    //sites/{site-id}/lists/{list-id}/items?$expand=fields,driveItem&$filter=fields/ContentType eq 'Document'

    //{site-id}/drive/root
    //{site-id}/drive/items/{drive-root-id}/children



    

