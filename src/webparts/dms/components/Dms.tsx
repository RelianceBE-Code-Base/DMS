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
      listItem?: {fields?:{Approver:string, DocIcon:string, Tag: string}}
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

  {
    key: 'newItem',
    text: "Refresh",
    iconProps: { iconName: 'refresh' },
    onClick: () => console.log('newItem')
  },

  {
    key: 'uploadItem',
    text: "Upload",
    iconProps: { iconName: 'Upload' },

    onClick: () => {showModal()}
    

    //onClick: () => console.log('uploadItem')
  },
  {
    key: 'downloadItem',
    text: "download",
    iconProps: { iconName: 'download' },
    onClick: () => {}
  },
  {
    key: 'delete',
    text: "Delete",
    iconProps: { iconName: 'Delete' },
    onClick: () => console.log('delete')
  },
  // {
  //   key: 'submit',
  //   text: "Submit",
  //   iconProps: { iconName: 'Send' },
  //   onClick: () => console.log('submit'),
  // }
];

const _farItems: ICommandBarItemProps[] = [
  {
    key: 'info',
    text: 'Info',
    ariaLabel: 'Info',
    iconOnly: true,
    iconProps: { iconName: 'Info' },
    onClick: () => openPanel(),
  },
];


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
      <a href={item.webUrl} target="_blank" rel="noopener noreferrer">
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
//    const [reload, setReload] = useState<boolean>(false);
    const [docfolderItem, setdocfolderItem] = useState<Array<DocArrayObj>>([]);
    const [filesItem, setfilesItem] = useState<DocArrayObj[]>([]);

    let [newfilesItem, setnewfilesItem] = useState([]);

    const [newfilesID, setnewfilesID] = useState("");
    const [refectch, setRefetch] = useState("");

    const [isLoadingMenu, setIsLoadingMenu] = useState(true);
    const [isLoadingUpload, setIsLoadingUpload] = useState(true);
    //const [isDialogVisible, setIsDialogVisible] = useState(false);
    const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] = useBoolean(false);
    const [isModalOpen, setIsModalOpen] = useState(false);

    const [dropdownOptions, setDropdownOptions] = useState<IDropdownOption[]>([]);
    const [selectedOption, setSelectedOption] = useState({text: ""} as any);

    const [selection, setSelection] = useState(new Selection());
    
    const [filesUpload, setfilesUpload] = useState({} as any)

    const [Approver, setApprover] = useState("")

    // const _getBodyModal = () => {
    //   // Implement the body content for your modal
    //   return (
    //     <div>
    //       <p>This is the modal body content</p>
    //     </div>
    //   );
    // };


    //const [deltaItems, setDeltaItems] = useState<Array<any>>([]); 

    //let resTrimmedID: string;

    //get site root folders
    const getSiteItem = async () => {
    props.context.msGraphClientFactory  
    .getClient('3')
      .then((client: MSGraphClientV3) => {
        client
          .api(`sites/root/drives`) //sites/${folderSite}:/ 
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
   

    const getUpdatedSiteItem = async () => {
      props.context.msGraphClientFactory  
      .getClient('3')
        .then((client: MSGraphClientV3) => {
          client
            .api(`sites/root/drives`) //sites/${folderSite}:/ 
            .version("v1.0")  //id, name, webUrl
            .get(async (err, res:IDocItems) => {
              if (err) {
                console.error("MSGraphAPI Error")
                console.error(err);
                return;
              }
        
              setdocfolderItem(res.value);
              setIsLoadingMenu(false);
  
                      res.value.forEach((drive) => {
               refetchSubRootDocumentItems(client, drive.id);            
          
               
             });       
            });
        });
      }
  
      // console.log(docfolderItem, "logged")
      console.log(filesItem, "filesOnly")

      const refetchSubRootDocumentItems  = async (client: MSGraphClientV3, driveId: string,) => {  //parentId?: string
          try {
            const deltaResponse = await client.api(`sites/root/drives/${driveId}/root/delta`).version('v1.0').select('*').get();  //
            const foldersOnly = deltaResponse.value.filter((item: any) =>  item.folder && item.name.toLowerCase() !== 'root');

            const filesOnly = deltaResponse.value.filter((item: any) => item.file);
           const listItemResponse = await client.api(`drives/${driveId}/root/children?expand=listItem`)
          .version('v1.0')
          .select('*')
          .get();

          console.log(filesOnly, "filesOnly")
          
                      setfilesItem((prevFileItems) => [
            ...filesOnly?.map((item: any) => ({
              ...item,
              parentId: item.parentReference ? item.parentReference.driveId : driveId,
            })),
          ]);
          setRefetch(newfilesID)
          setnewfilesID(newfilesID)

            } catch (error) {
            console.error(`Error fetching delta items for Drive ${driveId}:`, error);
          }
        };
  

 



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
    // const showDialog = () => {
    //   setIsDialogVisible(true);
    // };
  
//Hide Dialog
    // const hideDialog = () => {
    //   setIsDialogVisible(false);
    // };
  
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
    }, [props.context]);  

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
      Approver: Approver,
      
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
    getUpdatedSiteItem()
    // console.log("Updated fields",updatedFiles);
    
      setIsLoadingUpload(false);
      hideModal()
      
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
      />  
      </div>

      {/* Detailed List */}
          <div className={styles.details}>
          <CommandBar
                items={_commandBarItems}
                ariaLabel="Use left and right arrow keys to navigate between commands"
                farItems={_farItems}
              />
        
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
      <TextField label="Approver" onChange={handleApprover} disabled={isLoadingUpload}/>
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



    // const fileItem = await response.file.getItem();

    // await fileItem.update({
    //   Approval: 'Approved' /
    // });
    
    // const endpoint: string = props.context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle(${selectedOption?.text})/RootFolder/Files/add(url='${filesUpload.name}',overwrite=true)`;
    
    //     const requestOptions: ISPHttpClientOptions = {
    //       body: filesUpload,
    //       headers: {
    //         'Content-Type': 'application/octet-stream',
    //         'Accept': 'application/json'
    //       }
    //     };
    
    //     props.context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, requestOptions)
    //       .then((response: SPHttpClientResponse): Promise<any> => {
    //         if (response.ok) {
    //           return response.json();
    //         } else {
    //           console.log(`Error uploading file: ${response.statusText}`);
    //           return Promise.reject(response.statusText);
    //         }
    //       })
    //       .then((responseJSON: any): void => {
    //         console.log(`File uploaded successfully: ${responseJSON.ServerRelativeUrl}`);
    //       })
    //       .catch((error: any): void => {
    //         console.log('Error occurred:', error);
    //       });



// const FILE_ICONS: { name: string }[] = [
//   { name: "accdb" },
//   { name: "audio" },
//   { name: "code" },
//   { name: "csv" },
//   { name: "docx" },
//   { name: "dotx" },
//   { name: "mpp" },
//   { name: "mpt" },
//   { name: "model" },
//   { name: "one" },
//   { name: "onetoc" },
//   { name: "potx" },
//   { name: "ppsx" },
//   { name: "pdf" },
//   { name: "photo" },
//   { name: "pptx" },
//   { name: "presentation" },
//   { name: "potx" },
//   { name: "pub" },
//   { name: "rtf" },
//   { name: "spreadsheet" },
//   { name: "txt" },
//   { name: "vector" },
//   { name: "vsdx" },
//   { name: "vssx" },
//   { name: "vstx" },
//   { name: "xlsx" },
//   { name: "xltx" },
//   { name: "xsn" },
// ];

// function _randomFileIcon(): { docType: string; url: string } {
//   const docType: string =
//     FILE_ICONS[Math.floor(Math.random() * FILE_ICONS.length)].name;
//   return {
//     docType,
//     url: `https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/assets/item-types/16/${docType}.svg`,
//   };
// }







































//let itemSite: string = `${props.siteUrl}`
    //console.log ('site', itemSite)  //https://3zv7tj.sharepoint.com/

    //let folderSite: string = itemSite.replace("https://", "")
    //console.log ('folderSite:', folderSite) 















  //let userResult : IDocItems[] = []




    {/* <Dialog
          hidden={!isDialogVisible}
          onDismiss={hideDialog}
          dialogContentProps={null}
          isBlocking={true}>
            <div>
            <Dropdown
        placeholder="Select a Library"
        label="Libraries"
        options={options}
        styles={dropdownStyles}
      />
            </div>
          <DialogFooter>
            <PrimaryButton text="Save" onClick={null} />
            <DefaultButton text="Cancel" onClick={hideDialog} />
          </DialogFooter>
        </Dialog> */}
























// const convertToNavLinks = (items: DocArrayObj[]): INavLink[] => {
    //   return items.map((item) => ({
    //     name: item.name,
    //     url: item.webUrl,
    //     links: item.children ? convertToNavLinks(item.children) : undefined,
    //   }));
    // };

    // const navLinks: INavLink[] = convertToNavLinks(docItem);


   //   <HashRouter>
    //   <Switch>
    //     <Route component={HomeScreen} path="/" exact/>
    //   </Switch>
    // </HashRouter>

   //   <div>
    //     <ul>
    //     {docItem.map((item) => (
    //       <li key={item.id}>
    //       <a href={item.webUrl} target="_blank" rel="noopener noreferrer">{item.name}</a>
    //       </li>
    //     ))}
    //   </ul>
    // </div>


       //<div>Hello {JSON.stringify(docItem)}</div>


    // const useResTrimmedID = (trimmedID: string) => {
    //   // Now you can use trimmedID in this function

    //   //sites/3cecca9c-84ed-4cbd-a605-1fa4122d79ad,163e7858-e037-4433-8727-8587374af252/drives
    //   console.log("Using resTrimmedID in another function:", trimmedID);
  
    //   // Perform other actions with trimmedID
    // };


    

    //https://graph.microsoft.com/v1.0/sites/root/drive/root/search(q='')
    //sites/{site-id}/lists/{list-id}/items?$expand=fields,driveItem&$filter=fields/ContentType eq 'Document'

    //{site-id}/drive/root
    //{site-id}/drive/items/{drive-root-id}/children


//  const docLibs : IDocumentLibraryInformation[] = await _sp.site.getDocumentLibraries(url);
  // const subfolder = await Promise.all(docLibs.map(async(docLibItems: IDocumentLibraryInformation) => {
  //     const listRootFolder = await _sp.web.lists.getByTitle(docLibItems.Title).rootFolder();
  //     console.log(listRootFolder);
  //   }))

  //   return subfolder




    // const getSubRootDocumentItems = async (driveId: string) => {      
    //   props.context.msGraphClientFactory
    //   .getClient('3')
    //     .then((client: MSGraphClientV3) => {
    //       client
    //         .api(`sites/root/drives/${driveId}/root/delta`)    
    //         .version("v1.0")
    //         .select('id,name')
    //         .get((err, res:IDocItems) => {
    //           if (err) {
    //             console.error("MSGraphAPI Error")
    //             console.error(err);
    //             return;
    //           }
    //           console.log('Response for drive', driveId, ':', res);             
    //         });
    //     });
    //   }


        // //get files 
        // const getDriveFiles = async (driveId: string) => {
        //   try {
        //     const client = await props.context.msGraphClientFactory.getClient('3');
        //     const deltaResponse = await client.api(`sites/root/drives/${driveId}/root/delta`).version('v1.0').get();
      
        //     const filesOnly = deltaResponse.value.filter((item: any) => item.file);
        //     setDocItem(filesOnly);
        //   } catch (error) {
        //     console.error(`Error fetching files for Drive ${driveId}:`, error);
        //   }
        // };
    
    
    

    // const getRootDocumentItems = async (trimmedID: string) => {
      
    //   props.context.msGraphClientFactory
    //   .getClient('3')
    //     .then((client: MSGraphClientV3) => {
    //       client
    //         .api(`sites/${trimmedID}/drives`)    //https://graph.microsoft.com/v1.0/sites/3cecca9c-84ed-4cbd-a605-1fa4122d79ad,163e7858-e037-4433-8727-8587374af252/drives/b!nMrsPO2EvUymBR-kEi15rVh4PhY34DNEhyeFhzdK8lJ7fMD6tDJcQIifIUQ1jOkY/root/delta
    //         .version("v1.0")
    //         .select('id,name')
    //         .get((err, res:IDocItems) => {
    //           if (err) {
    //             console.error("MSGraphAPI Error")
    //             console.error(err);
    //             return;
    //           }
    //           console.log("Response:",res)
    //           setDocItem(res.value);
    //           //getRootDocumentItems(resTrimmedID);
    //         });
    //     });
    //   }

      // const getAllSubDocumentItems = async (mappedID: string) => {
      
      //   props.context.msGraphClientFactory
      //   .getClient('3')
      //     .then((client: MSGraphClientV3) => {
      //       client
      //         .api(`sites/${trimmedID}/drives/b!nMrsPO2EvUymBR-kEi15rVh4PhY34DNEhyeFhzdK8lJ7fMD6tDJcQIifIUQ1jOkY/root/delta`)    
      //         .version("v1.0")
      //         .select('id,name')
      //         .get((err, res:IDocItems) => {
      //           if (err) {
      //             console.error("MSGraphAPI Error")
      //             console.error(err);
      //             return;
      //           }
      //           console.log("Response:",res)
      //           setDocItem(res.value);
                
      //         });
      //     });
      //   }
        


