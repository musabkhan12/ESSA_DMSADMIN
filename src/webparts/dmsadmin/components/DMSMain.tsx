declare global {
    interface Window {
      managePermission:(DocumentLibraryName:string,SiteTilte:string , SiteID:string) => void;
      manageWorkflow:(DocumentLibraryName:string,SiteTilte:string , SiteID:string) => void;
      view:(message:string) => void;
      PreviewFile: (path: string, siteID: string, docLibName:any,  filemasterlist:any , filepreview:any) => void;
      deleteFile:(fileId: string , siteID:string, listToUpdate:any ) => void   
      shareFile:(fileID:String,siteId:String,currentFolderPathForFile:string,fileName:string)=> void;
      editFile:(siteName: string, documentLibraryName:string )=> void;
    }

  }

      // props for Manage work flow
// props for Manage work flow
const propsForManageWorkFlow={
  SiteTitle:"",
  DocumentLibraryName:"",
  SiteID:""
}
// props for managePermission
const managePermissionProps={
  SiteTitle:"",
  DocumentLibraryName:"",
  SiteID:""
}
  interface UploadFileProps {
    currentfolderpath: {
      CurrentEntity: string;
      currentEntityURL: string;
      currentsiteID: string;
      // ... other properties
    };
  }
  // export interface IDmsMusaibProps {
  //   description: string;
  //   isDarkTheme: boolean;
  //   environmentMessage: string;
  //   hasTeamsContext: boolean;
  //   userDisplayName: string;
  //   context: any;
  //   siteUrl: string;
  // }
  
  // @ts-ignore
  import * as React from "react";
  import { getSP } from "../loc/pnpjsConfig";
  import { SPFI } from "@pnp/sp";
  import "bootstrap/dist/css/bootstrap.min.css";
  // import "bootstrap//dist/"
  import "../../../CustomCss/mainCustom.scss";
  // import "../../verticalSideBar/components/VerticalSidebar.scss";
  import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";
  import UserContext from "../../../GlobalContext/context";
  import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
  import {  
    faUser, 
    faShareAlt, 
    faListSquares,
    faTableCells,
    faList
    // faTrash, 
    // faEdit, 
    // faEye  
  } from "@fortawesome/free-solid-svg-icons";
  import {
    faStar as faStarRegular,
    faFolder as faFolderRegular,
  
  } from "@fortawesome/free-regular-svg-icons";
  // import { useState , useEffect } from "react";
  import Provider from "../../../GlobalContext/provider";
  import { useMediaQuery } from "react-responsive";
  import "@pnp/sp/webs";
  import "@pnp/sp/folders";
  import "@pnp/sp/files";
  import "@pnp/sp/sites"
  import "@pnp/sp/presets/all"
  import "@pnp/sp/webs";
  import "@pnp/sp/sites";
  import "@pnp/sp/site-users/web";
  import { PermissionKind } from "@pnp/sp/security";
  import "bootstrap/dist/css/bootstrap.min.css";
  import "../../../CustomCss/mainCustom.scss";
   import "../../verticalSideBar/components/VerticalSidebar.scss";
  import "./dmscss.css";
  import { useState , useRef , useEffect} from "react";
  import UploadFile from "./UploadFile";
  import CreateFolder from "./CreateFolder";
  import Table from "./Table";
  import { IFileInfo } from "@pnp/sp/files";
  import { Popup } from "@fluentui/react";
  
  
  import {IDmsMusaibProps} from './IDmsMusaibProps'
  import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
import ManageWorkFlow from "./ManageWorkFlow";
import ManageFolderPermission from "./ManageFolderPermission";
  let sharewithmeicon = require('../assets/nodes.png')
  let sharewithothericon = require('../assets/share.png')
  let starticon = require('../assets/star.png')
  let listicon = require('../assets/list.png')
   let foldericon = require('../assets/foldericon.png')
  let Docicon = require("../assets/DOC.png");
  let Txticon = require("../assets/TXT.png");
  let Pdficon = require("../assets/PDF.png");
  let Xlsicon = require("../assets/XLS.png");
  let Zipicon = require("../assets/ZIP.png");
  let MainRounteVariable = 'MyRequest'
  
  let managePermissionIcon =  require('../assets/ManagePermission.svg') 
  // import managePermissionIcon from '../assets/ManagePermission.svg';
  let manageWorkFlowIcon =  require('../assets/ManageWorkflow.svg')
  // import manageWorkFlowIcon from '../assets/ManageWorkflow.svg';
  let viewIcon =  require('../assets/View.svg')
  // import viewIcon from '../assets/View.svg';
  let editIcon =  require('../assets/Edit.svg')
  // import editIcon from '../assets/Edit.svg';
  let deleteIcon =  require('../assets/Delete.svg')
  // import deleteIcon from '../assets/Delete.svg';
  let FillFavouriteFile = require('../assets/FillFavourite.svg')
  let ShareFile = require('../assets/Edit.svg')
  let UnFillFavouriteFile = require('../assets/UnFillFavourite.svg')
  let myfolderdata:any = []
  
  let currentDocumentLibrary = "";
  let currentFolder           = ""
  let currentfolderpath = "";
  // @ts-ignore
   let parentfolder            = ""
  let currentDevision = "";
    // @ts-ignore
  let currentDepartment       = ""
  let currentEntityURL = "";
    // @ts-ignore
  let currentEntity = ""
  let currentsiteID = ""
  let mydata: string[] = [];
  
  // start
  // let searchArray:any=[];
  let routeToDiffSideBar="";
  // end
  
  
  
  const ArgPoc = ({ props }: any) => {
    const sp: SPFI = getSP();
    // console.log(sp, "sp");
    const [showDeletepopup, setShowDeletepopup] = useState(false);
   const [activeButton] = React.useState<string>("");
    const { useHide }: any = React.useContext(UserContext);
    const elementRef = React.useRef<HTMLDivElement>(null);
    const [showFirstDiv, setShowFirstDiv] = useState(true);
    const [showworkflowdiv, setshowworkflowdiv] = useState('');
    const [showWorkflow, setShowWorkflow] = useState(false);
    const [showfolderpermission, setShowfolderpermission] = useState(false);
  
    // const handleButtonClickShow = () => {
    //   setShowFirstDiv(false);
    // };
  
  
    React.useEffect(() => {
      // console.log("This function is called only once", useHide);
  
      const showNavbar = (
        toggleId: string,
        navId: string,
        bodyId: string,
        headerId: string
      ) => {
        const toggle = document.getElementById(toggleId);
        const nav = document.getElementById(navId);
        const bodypd = document.getElementById(bodyId);
        const headerpd = document.getElementById(headerId);
  
        if (toggle && nav && bodypd && headerpd) {
          toggle.addEventListener("click", () => {
            nav.classList.toggle("show");
            toggle.classList.toggle("bx-x");
            bodypd.classList.toggle("body-pd");
            headerpd.classList.toggle("body-pd");
          });
        }
      };
  
      showNavbar("header-toggle", "nav-bar", "body-pd", "header");
  
      const linkColor = document.querySelectorAll(".nav_link");
  
      function colorLink(this: HTMLElement) {
        if (linkColor) {
          linkColor.forEach((l) => l.classList.remove("active"));
          this.classList.add("active");
        }
      }
  
      linkColor.forEach((l) => l.addEventListener("click", colorLink));
    }, [useHide]);
    // Media query to check if the screen width is less than 768px
    const isMobile = useMediaQuery({ query: "(max-width: 768px)" });
  
    React.useEffect(() => {
      // console.log("This function is called only once", useHide);
  
      const showNavbar = (
        toggleId: string,
        navId: string,
        bodyId: string,
        headerId: string
      ) => {
        const toggle = document.getElementById(toggleId);
        const nav = document.getElementById(navId);
        const bodypd = document.getElementById(bodyId);
        const headerpd = document.getElementById(headerId);
  
        if (toggle && nav && bodypd && headerpd) {
          toggle.addEventListener("click", () => {
            nav.classList.toggle("show");
            toggle.classList.toggle("bx-x");
            bodypd.classList.toggle("body-pd");
            headerpd.classList.toggle("body-pd");
          });
        }
      };
  
      showNavbar("header-toggle", "nav-bar", "body-pd", "header");
  
      const linkColor = document.querySelectorAll(".nav_link");
  
      function colorLink(this: HTMLElement) {
        if (linkColor) {
          linkColor.forEach((l) => l.classList.remove("active"));
          this.classList.add("active");
        }
      }
  
      linkColor.forEach((l) => l.addEventListener("click", colorLink));
    }, [useHide]);
    React.useEffect(() => {
      const handleEscape = (e: KeyboardEvent) => {
        if (e.key === "Escape") {
          if (document.fullscreenElement) {
            document.exitFullscreen();
          }
        }
      };
  
      window.addEventListener("keydown", handleEscape);
      return () => window.removeEventListener("keydown", handleEscape);
    }, []);
  /////////////////// DMS Code start / ////////////////////////////////////
  const buttonDivRef = useRef<HTMLDivElement>(null); 
  const [showMyrequButtons, setShowMyrequButtons] = useState(true); // Initially hidden
  const [showMyfavButtons, setShowMyfavButtons] = useState(false); // Initially hidden
  const [Myreqormyfav, setMyreqormyfav] = useState(''); // Initially hidden
  // console.log(Myreqormyfav , "Myreqormyfav")
    // console.log("This is current side ID",currentsiteID)
    const currentUserEmailRef = useRef('');
    useEffect(() => {
       getcurrentuseremail()

       
  }, []);
  const myrequestbuttonclick =()=>{
    const musa = document.getElementById('Myrequestbutton')
      if(musa){
        // alert("enter")
        musa.click();
        // alert("click")
      }
  
   }
  
   const getcurrentuseremail = async()=>{
    const userdata = await sp.web.currentUser();
    currentUserEmailRef.current = userdata.Email;
    myrequestbuttonclick()
    // console.log(currentUserEmailRef.current, "currentuser")
   }
  
  
    const fetchAndBuildTree2 = async () => {
      event.preventDefault()
      event.stopImmediatePropagation()
      event.stopPropagation()
      try {
        //Old working code
      //  Fetch data from EntityDivisionDepartmentMappingMasterList
        const entityItems = await sp.web.lists
          .getByTitle("EntityDivisionDepartmentMappingMasterList")
          .items.select(
            "Entitylookup/Title, Entitylookup/SiteURL", "Entitylookup/SiteID" ,
            "Devisionlookup/Title",
            "Departmentlookup/Title",
            "Devisionlookup/Active",
            "Departmentlookup/Active"
          )
          .expand("Entitylookup", "Devisionlookup", "Departmentlookup")
          .filter("Entitylookup/Active eq 'Yes'")();
           console.log(entityItems, "entityItems 1")
          const uniqueEntityMap = new Map();
          const uniqueEntitiesWithAccess: any = [];
          
          // Loop through each item and check permissions
          for (const item of entityItems) {
            const entityTitle = item.Entitylookup.Title;
            try {
              const subsiteWeb = await sp.site.openWebById(item.Entitylookup.SiteID);
              const hasAccess = await subsiteWeb.web.currentUserHasPermissions(PermissionKind.ViewListItems);
          
              if (hasAccess) {
                // Add to uniqueEntitiesWithAccess only if user has access
                uniqueEntityMap.set(entityTitle, item); // Store the item or any required data
                uniqueEntitiesWithAccess.push(item);  // Add the item to the list of entities with access
                console.log(`User has access to site: ${entityTitle}`, item);
              } else {
                console.log(`User does not have access to site: ${entityTitle}`);
              }
            } catch (error) {
              console.error(`Error while checking access for site: ${entityTitle}`, error);
            }
          }
    console.log(uniqueEntityMap , "uniqueEntityMap ......")
    console.log(uniqueEntitiesWithAccess , "uniqueEntitiesWithAccess");
        /// New Code 
  
  
        // Fetch data from DMSFolderMaster
        const folderItems = await sp.web.lists
          .getByTitle("DMSFolderMaster")
          .items.getAll();
         console.log("folderItems", folderItems);
  
        const myButton = document.getElementById("mybutton");
             const createFileButton=document.getElementById("createFileButton");
             const createFileButton2=document.getElementById("createFileButton2");
                   const createFolderButton=document.getElementById("createFolderButton");
        // Create a map to hold folder data by SiteTitle, Devision, Department
        const folderMap = new Map();
        folderItems.forEach((folderItem) => {
          const {
            SiteTitle,
            Devision,
            Department,
            DocumentLibraryName,
            FolderName,
            ParentFolderId,
            FolderPath,
          } = folderItem;
          if (SiteTitle) {
            const key = `${SiteTitle.trim()}::${Devision?.trim() || ""}::${
              Department?.trim() || ""
            }`;
            if (!folderMap.has(key)) {
              folderMap.set(key, []);
            }
            if (DocumentLibraryName) {
              folderMap
                .get(key)
                .push({
                  FolderPath,
                  ParentFolderId,
                  DocumentLibraryName,
                  FolderName: Array.isArray(FolderName)
                    ? FolderName
                    : [FolderName],
                });
            }
          }
        });
        // console.log(folderMap, "folderMap");
        // const entitiesMap = new Map();
        const entitiesMap: any = new Map();
  
        uniqueEntitiesWithAccess.forEach((item:any) => {
          const entityTitle = item.Entitylookup.Title;
          const siteURL = item.Entitylookup.SiteURL;
          const siteID = item.Entitylookup.SiteID;
        
          if (!entitiesMap.has(entityTitle)) {
            entitiesMap.set(entityTitle, {
              siteURL: siteURL,
              entityTitle: entityTitle,
              siteID: siteID,
              devisions: new Map(),
            });
          }
        
          const entry = entitiesMap.get(entityTitle);
          const devisionTitle = item.Devisionlookup?.Title;
          const departmentTitle = item.Departmentlookup?.Title;
          const isDevisionActive = item.Devisionlookup?.Active === "Yes";
          const isDepartmentActive = item.Departmentlookup?.Active === "Yes";
        
          if (devisionTitle && isDevisionActive) {
            if (!entry.devisions.has(devisionTitle)) {
              entry.devisions.set(devisionTitle, {
                departments: new Set(),
                docLibs: new Set(),
              });
            }
            const devisionEntry = entry.devisions.get(devisionTitle);
            if (departmentTitle && isDepartmentActive) {
              devisionEntry.departments.add(departmentTitle);
            } else if (!departmentTitle || !isDepartmentActive) {
              const nullDeptKey = `${entityTitle.trim()}::${devisionTitle.trim()}::`;
              // Handle case where department is null or inactive
            }
          }
        });
        const buildFolderStructure = (
          folderList: HTMLElement,
          folders: any[],
          parentFolderId: string | null
        ) => {
          const filteredFolders = folders.filter(
            (folder) => folder.ParentFolderId === parentFolderId
          );
          filteredFolders.forEach((folder) => {
            const folderElement = document.createElement("li");
            folderElement.textContent = folder.FolderName;
            folderList.appendChild(folderElement);
  
            const childFolderList = document.createElement("ul");
            childFolderList.style.display = "none";
            folderElement.appendChild(childFolderList);
  
            folderElement.addEventListener("click", (event) => {
              event.stopPropagation();
              // currentFolder = folder.FolderName;
              toggleVisibility(childFolderList);
            });
  
            // Recursively build the structure for subfolders
            buildFolderStructure(childFolderList, folders, folder.FolderName);
          });
        };
        // Build the folder tree structure in the DOM
        const container = document.getElementById("folderContainer2");
  
        if (container) {
          container.innerHTML = ""; // Clear previous contents
        } else {
          console.error("Container element not found");
        }
        // container.innerHTML = ''; // Clear previous contents
  
        const toggleVisibility = (element: any, forceShow = false) => {
          const isVisible = element.style.display === "block";
          element.style.display = isVisible && !forceShow ? "none" : "block";
        };
        const createImageElement = (src: string, alt: string) => {
          const img = document.createElement("img");
          img.src = require("../assets/add-folder.png");
          img.alt = alt;
          img.style.float = "left";
          img.style.width = "20px"; // Adjust the size as needed
          img.style.height = "20px"; // Adjust the size as needed
          img.style.marginRight = "5px"; // Space between image and text
          return img;
        };
          const createToggleButton = () => {
          const link = document.createElement("a");
          link.textContent = "+"; // Initial text
          link.className="toggle-button"
          link.style.cursor = "pointer";
          link.style.textDecoration = "none";
          
          link.addEventListener("click", (e) => {
              e.preventDefault()
              console.log("Button clicked +/-");
              if (link.textContent === "+") {
                  link.textContent = "-"; // Change to minus when content is visible
              } else if(link.textContent){
                  link.textContent = "+"; // Change to plus when content is hidden
              }
          });
      
          return link;
      };
      
        entitiesMap.forEach((value:any, entityTitle:any) => {
          const titleElement = document.createElement("p");
  
          // titleElement.textContent = entityTitle;
          titleElement.classList.add("folder", "icon");
          titleElement.style.cursor = "pointer";
          // const entityImage = createImageElement(
          //   "icons/entity-icon.png",
          //   "Entity Icon"
          // );
          const toggleButton=createToggleButton();
          titleElement.appendChild(toggleButton);
          titleElement.appendChild(document.createTextNode(entityTitle));
  
          if (container) {
            container.appendChild(titleElement);
          } else {
            console.error("Container element not found");
          }
  
          const documentList = document.createElement("ul");
          titleElement.appendChild(documentList);
          documentList.style.display = "none";
          /////start: Display Document library with recursive folder under Enitiy directly when Devision and Department Null /////
          const nullKey = `${entityTitle.trim()}::::`;
          if (folderMap.has(nullKey)) {
            const documentLibraries = folderMap.get(nullKey) || [];
  
            // Create a map to store unique DocumentLibraryNames and their details
            const uniqueDocLibs = new Map();
  
            // Iterate over document libraries and populate the map with unique DocumentLibraryNames
            documentLibraries.forEach((item: any) => {
              if (!uniqueDocLibs.has(item.DocumentLibraryName)) {
                uniqueDocLibs.set(item.DocumentLibraryName, {
                  folders: [],
                  folderPath: item.FolderPath, // Store FolderPath with other details
                });
              }
              uniqueDocLibs.get(item.DocumentLibraryName).folders.push(item);
            });
  
            // Now render each unique DocumentLibraryName and its associated folders
            uniqueDocLibs.forEach((data, docLibName) => {
              const docLibElement = document.createElement("li");
              docLibElement.textContent = docLibName;
  
              // Optionally display the FolderPath in the docLibElement
              const pathText = document.createElement("span");
              // pathText.textContent = ` (${data.folderPath})`; // Display FolderPath
              docLibElement.appendChild(pathText);
  
              documentList.appendChild(docLibElement);
  
              const folderList = document.createElement("ul");
              folderList.style.display = "none";
              folderList.style.width = "240px";
              const entityImage = createImageElement(
                "icons/entity-icon.png",
                "Entity Icon"
              );
              docLibElement.appendChild(entityImage);
              docLibElement.appendChild(folderList);
  
              // Handle click to toggle the visibility of the folder list
              docLibElement.addEventListener("click", (event:any) => {
                event.preventDefault()
                event.stopPropagation();
                // setlistorgriddata('')
                // setShowMyrequButtons(false)
                // setShowMyfavButtons(false)
                handleNavigation(value.entityTitle, null , null , docLibName , null )
                toggleVisibility(folderList);
                getdoclibdata(data.folderPath , value.siteID , docLibName);
                currentfolderpath = data.folderPath
                currentDocumentLibrary = docLibName;
                currentEntityURL = value.siteURL;
                currentEntity = value.entityTitle
                currentsiteID = value.siteID
                currentDevision = ''
                currentDepartment = ''
                currentFolder = ''
       
                console.log(currentEntityURL , "currentEntityURL")
                console.log(currentsiteID , "currentsiteID")
                console.log(currentEntity , "currentEntity")
                console.log(currentDocumentLibrary , "currentFolder")
                console.log(currentfolderpath , "currentfolderpath")
                console.log(currentDevision , "currentfolderpath")
                console.log(currentDepartment , "currentfolderpath")
                     createFileButton.style.display = "block";
                     createFileButton2.style.display = "block";
                      if(createFolderButton){
                  createFolderButton.style.display="block"
                }
                
                if(createFileButton){
                  createFileButton.style.display = "block";
                }
                if(createFileButton2){
                  createFileButton2.style.display = "block";
                }
                      
                if (myButton) {
                  myButton.textContent = `Create Folder under ${docLibName}`;
                } else {
                  console.error();
                }
              });
  
              // Handle double-click to hide the folder list
              docLibElement.addEventListener("dblclick", (event) => {
                event.stopPropagation();
                toggleVisibility(folderList, false);
              });
  
              // Function to build the folder structure recursively
              const buildFolderStructure = (
                parentFolderId: any,
                parentElement: any
              ) => {
                data.folders.forEach((item: any) => {
                  const folderNamesArray = Array.isArray(item.FolderName)
                    ? item.FolderName
                    : [item.FolderName];
  
                  folderNamesArray.forEach((folderName: any) => {
                    if (folderName && item.ParentFolderId === parentFolderId) {
                      // Only display non-null folder names
                      const folderElement = document.createElement("li");
                      folderElement.textContent = folderName;
                      parentElement.appendChild(folderElement);
                      const entityImage = createImageElement(
                        "icons/entity-icon.png",
                        "Entity Icon"
                      );
                      folderElement.appendChild(entityImage);
                      const subFolderList = document.createElement("ul");
                      subFolderList.style.display = "none";
                      subFolderList.style.width = "240px";
                      folderElement.appendChild(subFolderList);
  
                      folderElement.addEventListener("click", (event:any) => {
                         event.preventDefault();  // Prevent default action
                         event.stopPropagation();  // Stop event bubbling
                         console.log("Event listener triggered");
                        currentEntityURL = value.siteURL;
                        currentsiteID = value.siteID
                        currentEntity = value.entityTitle
                        currentDocumentLibrary = docLibName;
                        currentFolder  = folderName;
                        parentfolder = item.ParentFolderId;
                        currentfolderpath = item.FolderPath;
                        currentDevision = ''
                        currentDepartment = ''
                        console.log(currentEntityURL , "currentEntityURL")
                        console.log(currentsiteID , "currentsiteID")
                        console.log(currentEntity , "currentEntity")
                        console.log(currentDocumentLibrary , "currentDocumentLibrary")
                        console.log(currentFolder , "currentFolder")
                        console.log(parentfolder , "parentfolder")
                        console.log(currentfolderpath , "currentfolderpath");
                        handleNavigation(value.entityTitle, null , null , docLibName , folderName )
                        event.stopPropagation();
                        getdoclibdata(item.FolderPath,currentsiteID ,docLibName )
                        if (myButton) {
                          myButton.textContent = `Create Folder under ${folderName}`;
                        } else {
                          console.error();
                        }
  
          
                        toggleVisibility(subFolderList);
  
                        // Clear existing sub-folder list to avoid duplications
                        subFolderList.innerHTML = "";
  
                        // Recursively build the sub-folder structure
                        buildFolderStructure(folderName, subFolderList);
                      });
                    }
                  });
                });
              };
  
              // Start building the folder structure from the root level (null ParentFolderId)
              buildFolderStructure(null, folderList);
            });
          }
          /////End Display Document library with recursive folder under Enitiy directly when Devision and Department Null /////
          const devisionList = document.createElement("ul");
          devisionList.style.display = "none";
          titleElement.appendChild(devisionList);
  
          value.devisions.forEach((devisionValue: any, devisionTitle: any) => {
            const devisionElement = document.createElement("li");
            devisionElement.textContent = devisionTitle;
            devisionElement.classList.add("folder", "icon");
            devisionElement.style.cursor = "pointer";
            devisionList.appendChild(devisionElement);
  
            const docLibList = document.createElement("ul");
            docLibList.style.display = "none";
            const entityImage = createImageElement(
              "icons/entity-icon.png",
              "Entity Icon"
            );
            devisionElement.appendChild(entityImage);
            devisionElement.appendChild(docLibList);
  
            // Display unique DocumentLibraryName under Devision
            devisionValue.docLibs.forEach((docLibName: any) => {
              const docLibElement = document.createElement("li");
              docLibElement.textContent = docLibName;
              docLibElement.classList.add("file-icon", "icon");
              docLibList.appendChild(docLibElement);
  
              const folderList = document.createElement("ul");
              folderList.style.display = "none";
  
              docLibElement.appendChild(folderList);
  
              const docLibKey = `${entityTitle.trim()}::${devisionTitle.trim()}::`;
              const docLibFolders = folderMap.get(docLibKey) || [];
              docLibFolders.forEach((folderItem: any) => {
                const folderElement = document.createElement("li");
                folderElement.textContent = folderItem.FolderName;
  
                folderList.appendChild(folderElement);
              });
  
              docLibElement.addEventListener("click", (event) => {
                console.log(devisionValue, "devisionValue");
                event.stopPropagation();
                currentDocumentLibrary = docLibName;
                // currentFolder = '';
                currentDevision = devisionTitle;
                // currentDepartment = '';
                currentEntityURL = value.siteURL;
                currentEntity = value.entityTitle
                currentsiteID = value.siteID
         
                console.log("currentEntityURL", currentEntityURL);
                console.log("currentEntity", currentEntity);
                console.log("currentsiteID", currentsiteID);
                console.log("currentDevision", currentDevision);
                console.log("currentDocumentLibrary", currentDocumentLibrary);
                if (myButton) {
                  myButton.textContent = `Create Library under ${docLibName}`;
                } else {
                  console.error();
                }
  
                toggleVisibility(folderList);
              });
  
              docLibElement.addEventListener("dblclick", (event) => {
                event.stopPropagation();
                toggleVisibility(folderList, false);
              });
            });
  
            const departmentList = document.createElement("ul");
  
            departmentList.style.display = "none";
            devisionElement.appendChild(departmentList);
  
            devisionValue.departments.forEach((departmentTitle: any) => {
              const departmentElement = document.createElement("li");
              departmentElement.textContent = departmentTitle;
              departmentElement.classList.add("folder");
              departmentElement.style.cursor = "pointer";
              departmentList.appendChild(departmentElement);
  
              const documentList = document.createElement("ul");
              documentList.style.display = "none";
              documentList.style.width = "300px";
              const entityImage = createImageElement(
                "icons/entity-icon.png",
                "Entity Icon"
              );
              departmentElement.appendChild(entityImage);
              departmentElement.appendChild(documentList);
  
              departmentElement.addEventListener("click", (event) => {
                currentEntityURL = value.siteURL;
                      currentsiteID = value.siteID
                      currentEntity = value.entityTitle;
                      currentDevision = devisionTitle;
                      currentDepartment = departmentTitle;
                      currentDocumentLibrary = ''
                      currentFolder = ''
                      currentfolderpath = ''
                    console.log("currentEntityURL", currentEntityURL);
                    console.log("currentsiteID", currentsiteID);
                    console.log("currentEntity", currentEntity);
                    console.log("currentDevision", currentDevision);
                    console.log("currentDepartment", currentDepartment);
                    handleNavigation(value.entityTitle, devisionTitle , departmentTitle , null , null )
                event.stopPropagation();
                if (myButton) {
                  myButton.textContent = `Create Library under ${departmentTitle}`;
                } else {
                  console.error();
                }
  
                // Prevent toggling visibility before the list is populated
                if (documentList.innerHTML === "") {
                  const key = `${entityTitle.trim()}::${devisionTitle.trim()}::${departmentTitle.trim()}`;
                  const documentLibraries = folderMap.get(key) || [];
                  documentList.innerHTML = ""; 
                  const uniqueDocLibs = new Map();
  
                  documentLibraries.forEach((item: any) => {
                    if (!uniqueDocLibs.has(item.DocumentLibraryName)) {
                      uniqueDocLibs.set(item.DocumentLibraryName, {
                        folders: [],
                        folderPath: item.FolderPath, // Store FolderPath
                      });
                    }
                    uniqueDocLibs
                      .get(item.DocumentLibraryName)
                      .folders.push(item);
                  });
  
                  uniqueDocLibs.forEach((data, docLibName) => {
                    console.log(uniqueDocLibs , "uniqueDocLibs")
                    const docLibElement = document.createElement("li");
                    docLibElement.textContent = docLibName;
  
                    // Optionally display the FolderPath in the docLibElement
                    // const pathText = document.createElement("span");
                    // pathText.textContent = ` (${data.folderPath})`; // Display FolderPath
                    // docLibElement.appendChild(pathText);
  
                    documentList.appendChild(docLibElement);
  
                    const folderList = document.createElement("ul");
                    folderList.style.display = "none";
                    folderList.style.width = "351px";
                    const entityImage = createImageElement(
                      "icons/entity-icon.png",
                      "Entity Icon"
                    );
                    docLibElement.appendChild(entityImage);
                    docLibElement.appendChild(folderList);
  
                    docLibElement.addEventListener("click", (event) => {
                      event.stopPropagation();
                      currentEntityURL = value.siteURL;
                      currentsiteID = value.siteID
                      currentEntity = value.entityTitle;
                      currentDevision = devisionTitle;
                      currentDocumentLibrary = docLibName;
                      currentDepartment = departmentTitle;
                      currentfolderpath = data.folderPath,
                      currentFolder =''
                      console.log(data, data  ,"data")
                    console.log("currentEntityURL", currentEntityURL);
                    console.log("currentsiteID", currentsiteID);
                    console.log("currentEntity", currentEntity);
                    console.log("currentDevision", currentDevision);
                    console.log("currentDepartment", currentDepartment);
                    console.log("currentDocumentLibrary", currentDocumentLibrary);
                    console.log("currentfolderpath", currentfolderpath);
                    console.log("parentfolder", parentfolder);
                    getdoclibdata(data.folderPath , value.siteID , docLibName)
                    handleNavigation(value.entityTitle, devisionTitle , departmentTitle , docLibName , null )
                      console.log(
                        "FolderPath for document library:",
                        data.folderPath
                      );
                      toggleVisibility(folderList);
                          const createFileButton=document.getElementById("createFileButton")
                          const createFileButton2=document.getElementById("createFileButton")
                      createFileButton.style.display="block";
                      createFileButton2.style.display="block";
                      if (myButton) {
                        myButton.textContent = `Create Folder under ${docLibName}`;
                      } else {
                        console.error();
                      }
                    });
  
                    docLibElement.addEventListener("dblclick", (event) => {
                      event.stopPropagation();
                      toggleVisibility(folderList, false);
                    });
                    const buildFolderStructure = (
                      parentFolderId: any,
                      parentElement: any
                    ) => {
                      data.folders.forEach((item: any) => {
                    
                        const folderNamesArray = Array.isArray(item.FolderName)
                          ? item.FolderName
                          : [item.FolderName];
  
                        folderNamesArray.forEach((folderName: any) => {
              
                          if (
                            folderName &&
                            item.ParentFolderId === parentFolderId
                          ) {
                            const folderElement = document.createElement("li");
                            folderElement.textContent = folderName;
                            parentElement.appendChild(folderElement);
                            const entityImage = createImageElement(
                              "icons/entity-icon.png",
                              "Entity Icon"
                            );
                            folderElement.appendChild(entityImage);
                            const subFolderList = document.createElement("ul");
                            subFolderList.style.display = "none";
                            folderElement.appendChild(subFolderList);
  
                            folderElement.addEventListener("click", (event) => {
                              currentEntityURL = value.siteURL;
                              currentEntity = value.entityTitle;
                              currentsiteID = value.siteID
                              currentDevision = devisionTitle;
                              currentDepartment = departmentTitle;
                              currentDocumentLibrary = docLibName;
                              currentFolder = folderName
                    
                            console.log("currentEntityURL", currentEntityURL);
                            console.log("currentEntity", currentEntity);
                            console.log("currentsiteID", currentsiteID);
                            console.log("currentDevision", currentDevision);
                            console.log("currentDepartment", currentDepartment);
                            console.log("currentDocumentLibrary", currentDocumentLibrary);
                            console.log("currentfolderpath", item.FolderPath);
                            getdoclibdata(item.FolderPath,currentsiteID , docLibName)
                            handleNavigation(value.entityTitle, devisionTitle , departmentTitle , docLibName , folderName )
                                 const createFileButton=document.getElementById("createFileButton")
                            createFileButton.style.display="block";
                                 const createFileButton2=document.getElementById("createFileButton")
                            createFileButton2.style.display="block";
                              if (myButton) {
                                myButton.textContent = `Create Folder under ${folderName}`;
                              } else {
                                console.error();
                              }
                              event.stopPropagation();
                              toggleVisibility(subFolderList);
                              subFolderList.innerHTML = "";
                              buildFolderStructure(folderName, subFolderList);
                            });
                          }
                        });
                      });
                    };
                    buildFolderStructure(null, folderList);
                  });
                }
  
                toggleVisibility(documentList);
              });
  
              departmentElement.addEventListener("dblclick", (event) => {
                event.stopPropagation();
                toggleVisibility(documentList, false);
              });
            });
  
            ///Start: display all Document libraries under Devision directly if Department null with nested folder //////
            const keyForDevisionOnly = `${entityTitle.trim()}::${devisionTitle.trim()}::`;
  
            if (folderMap.has(keyForDevisionOnly)) {
              const documentLibraries = folderMap.get(keyForDevisionOnly) || [];
              // console.log(documentLibraries, "documentLibraries");
              const uniqueDocLibNames = new Set();
  
              documentLibraries.forEach((item: any) => {
                const normalizedDocLibName =
                  item.DocumentLibraryName.trim().toLowerCase();
  
                if (!uniqueDocLibNames.has(normalizedDocLibName)) {
                  uniqueDocLibNames.add(normalizedDocLibName);
  
                  const docLibElement = document.createElement("li");
                  docLibElement.textContent = item.DocumentLibraryName;
                  departmentList.appendChild(docLibElement);
  
                  const folderList = document.createElement("ul");
                  folderList.style.display = "none";
                  const entityImage = createImageElement(
                    "icons/entity-icon.png",
                    "Entity Icon"
                  );
                  docLibElement.appendChild(entityImage);
  
                  docLibElement.appendChild(folderList);
  
                  docLibElement.addEventListener("click", (event) => {
                    event.stopPropagation();
                    currentEntityURL = value.siteURL; // Use the SiteURL from entitiesMap
                    currentsiteID = value.siteID
                    currentEntity = value.entityTitle
                    currentDevision = devisionTitle;
                    currentDepartment = ''
                    currentFolder=''
                    currentDocumentLibrary = item.DocumentLibraryName;
                    currentfolderpath = item.FolderPath;
                    console.log("currentEntityURL", currentEntityURL);
                    console.log("currentsiteID", currentsiteID);
                    console.log("currentEntity", currentEntity);
                    console.log("currentDevision", currentDevision);
                    console.log("currentDepartment", currentDepartment);
                    console.log("currentDocumentLibrary", currentDocumentLibrary);
                    console.log("currentfolderpath", currentfolderpath);
                    getdoclibdata(item.FolderPath , value.siteID , item.DocumentLibraryName)
                    handleNavigation(value.entityTitle , devisionTitle, null , item.DocumentLibraryName )
                    const createFileButton=document.getElementById("createFileButton")
                    createFileButton.style.display="block";
                    const createFileButton2=document.getElementById("createFileButton")
                    createFileButton2.style.display="block";
                    if (myButton) {
                      myButton.textContent = `Create Folder under ${item.DocumentLibraryName}`;
                    } else {
                      console.error();
                    }
                    toggleVisibility(folderList);
                    folderList.innerHTML = "";
                    const buildFolderStructure = (
                      parentFolderId: any,
                      parentElement: any
                    ) => {
                      const createImageElement = (src: string, alt: string) => {
                        const img = document.createElement("img");
                        img.src = require("../assets/add-folder.png");
                        img.alt = alt;
                        img.style.float = "left";
                        img.style.width = "20px"; // Adjust the size as needed
                        img.style.height = "20px"; // Adjust the size as needed
                        img.style.marginRight = "5px"; // Space between image and text
                        return img;
                      };
                      documentLibraries.forEach((libItem: any) => {
                      
                        if (
                          libItem.DocumentLibraryName.trim().toLowerCase() ===
                          normalizedDocLibName
                        ) {
                          const folderNamesArray = Array.isArray(
                            libItem.FolderName
                          )
                            ? libItem.FolderName
                            : [libItem.FolderName];
  
                          folderNamesArray.forEach((folderName: any) => {
                           
                            if (
                              folderName &&
                              libItem.ParentFolderId === parentFolderId
                            ) {
                              // Only display non-null folder names
                              const folderElement2 = document.createElement("li");
                              folderElement2.textContent = folderName;
                              parentElement.appendChild(folderElement2);
                              const folderPath = libItem.FolderPath; 
                              const entityImage = createImageElement(
                                "icons/entity-icon.png",
                                "Entity Icon"
                              );
                              folderElement2.appendChild(entityImage);
                              const subFolderList2 = document.createElement("ul");
                              subFolderList2.style.display = "none";
  
                              // const entityImage = createImageElement('icons/entity-icon.png', 'Entity Icon')
                              folderElement2.appendChild(entityImage);
                              subFolderList2.appendChild(entityImage);
                              folderElement2.appendChild(subFolderList2);
  
                              folderElement2.addEventListener(
                                
                                "click",
                                (event) => {
                                  currentEntityURL = value.siteURL; // Use the SiteURL from entitiesMap
                                  currentsiteID = value.siteID
                                  currentEntity = value.entityTitle
                                  currentDevision = devisionTitle;
                                  currentDepartment = null
                                  currentDocumentLibrary = item.DocumentLibraryName;
                                  currentFolder = folderName
                                  currentDepartment=''
                                  currentFolder = folderPath
                                  // currentfolderpath = item.FolderPath;
                                  parentfolder = parentFolderId
                                  console.log("currentEntityURL", currentEntityURL);
                                  console.log("currentsiteID", currentsiteID);
                                  
                                  console.log("currentEntity", currentEntity);
                                  console.log("currentDevision", currentDevision);
                                  console.log("currentDepartment", currentDepartment);
                                  console.log("currentDocumentLibrary", currentDocumentLibrary);
                                  console.log("currentFolder", currentFolder);
                                  console.log("currentfolderpath", folderPath);
                                  console.log("parentfolder", parentfolder);
                                  handleNavigation(value.entityTitle , devisionTitle ,null , item.DocumentLibraryName , folderName)
                                  event.stopPropagation();
                                  toggleVisibility(subFolderList2);
                                  console.log("enter ee");
                                  getdoclibdata(folderPath,currentsiteID, item.DocumentLibraryName)
                                    const createFileButton=document.getElementById("createFileButton")
                                  createFileButton.style.display="block";
                                    const createFileButton2=document.getElementById("createFileButton")
                                  createFileButton2.style.display="block";
                                  if (myButton) {
                                    myButton.textContent = `Create Folder under ${folderName}`;
                                  } else {
                                    console.error();
                                  }
  
                                  // Clear existing sub-folder list to avoid duplications
                                  subFolderList2.innerHTML = "";
  
                                  // Recursively build the sub-folder structure
                                  buildFolderStructure(
                                    folderName,
                                    subFolderList2
                                  );
                                }
                              );
                            }
                          });
                        }
                      });
                    };
  
                    // Start building the folder structure from the root level (null ParentFolderId)
                    buildFolderStructure(null, folderList);
                  });
  
                  // Optionally, expand the folder structure by default
                  // buildFolderStructure(folderList, documentLibraries, null);
                }
              });
            }
  
            ///End: display all Document libraries under Devision directly if Department null with nested folder //////
  
            devisionElement.addEventListener("click", (event) => {
              event.stopPropagation();
              currentDevision = devisionTitle;
              currentEntityURL = value.siteURL;
              currentEntity = value.entityTitle
              currentsiteID = value.siteID
              currentDepartment = ''
              currentDocumentLibrary = ''
              currentFolder =''
              currentfolderpath = ''
              console.log("currentEntityURL", currentEntityURL);
              console.log("currentsiteID", currentsiteID);
              console.log("currentEntity", currentEntity);
              console.log("currentDevision", currentDevision);
              handleNavigation(value.entityTitle , devisionTitle , null , null , null)
              toggleVisibility(departmentList);
              // Toggle plus/minus icon
              devisionElement.classList.remove("expanded");
               // const //createFileButton=document.getElementById("createFileButton")
             // createFileButton.style.display="block";
              if (myButton) {
                myButton.textContent = `Create Library under ${devisionTitle}`;
              } else {
                console.error();
              }
            });
  
            devisionElement.addEventListener("dblclick", (event) => {
              event.stopPropagation();
              toggleVisibility(departmentList, false);
              // Toggle plus/minus icon
              devisionElement.classList.remove("expanded");
            });
          });
  
          let clickTimer:any;
  
          titleElement.addEventListener("click", (event) => {
              event.stopPropagation();
              const createFileButton2 = document.getElementById("createFileButton2");
              // Clear any existing timer
              clearTimeout(clickTimer);
          
              // Set a new timer
              clickTimer = setTimeout(() => {
                  setlistorgriddata('');
                  currentEntity= value.entityTitle
                  currentEntityURL = value.siteURL;
                  currentsiteID = value.siteID;
                  currentDevision=""
              currentDepartment =''
                  currentDocumentLibrary=""
                  currentFolder=""
                  currentfolderpath=""
                  console.log(value.entityTitle, "value");
                  console.log(currentsiteID, "currentsiteID");
                  console.log("currentEntityURL", currentEntityURL);
                  mydata.push(value.siteURL);
                  console.log(mydata, "my mydata");
                  toggleVisibility(devisionList);
                  toggleVisibility(documentList);
                  const hidegidvewlistviewbutton = document.getElementById("hidegidvewlistviewbutton");
                  const hidegidvewlistviewbutton2 = document.getElementById("hidegidvewlistviewbutton2");
                  if (hidegidvewlistviewbutton) {
                      console.log("enter here .....................");
                      hidegidvewlistviewbutton.style.display = 'none';
                  }
                  if (hidegidvewlistviewbutton2) {
                      console.log("enter here .....................");
                      hidegidvewlistviewbutton2.style.display = 'none';
                  }
                  handleNavigation(value.entityTitle, null, null, null, null);
                  // Toggle plus/minus icon
                  titleElement.classList.toggle("expanded");
                  console.log(value, "value");
                  const createFileButton = document.getElementById("createFileButton");
                  const createFileButton2 = document.getElementById("createFileButton2");
                  // if (createFolderButton) {
                  //     createFolderButton.style.display = "block";
                  // }
                  if (createFileButton) {
                      createFileButton.style.display = "none";
                  }
                  if (createFileButton2) {
                      createFileButton2.style.display = "block";
                  }
                  if (myButton) {
                      myButton.textContent = `Create Library under ${entityTitle}`;
                  } else {
                      console.error();
                  }
                  // fetchData(currentEntityURL);
              }, 300); // Adjust the delay as needed
          });
          
          titleElement.addEventListener("dblclick", (event) => {
              event.stopPropagation();
          
              // Clear the single click timer
              clearTimeout(clickTimer);
          
              setlistorgriddata('');
              toggleVisibility(devisionList, false);
              toggleVisibility(documentList, false);
              // Toggle plus/minus icon
              titleElement.classList.remove("expanded");
          });
        });
      } catch (error) {
        console.error("Error fetching or building folder tree:", error);
      }
    };
    fetchAndBuildTree2();
    // Call the function to fetch data and build the tree
    // thi is working new function for getting files from documnet library with pagination batching
    const getdoclibdata = async (FolderPath: any , siteID:any , docLibName:any) => {
      const wait = document.getElementById('files-container')
      wait.classList.remove('hidemydatacards')
      // const subsite = await sp.web.getByTitle('Central Trading Company').select("Id")();
      // console.log("subsite ID" , subsite)
      // event.preventDefault()
      // event.stopPropagation()
      // setlistorgriddata('')
      // setShowMyrequButtons(false)
      // setShowMyfavButtons(false)
      console.log('path   ', FolderPath)
      console.log('SiteID :    ', siteID)
      
      // start
      // Empty the routeToDiffSideBar
      routeToDiffSideBar="";
      // end  
  
      const testidsub = await sp.site.openWebById(siteID);
      let files:any = [];
      let batchSize = 5000;
      let nextLink = null;
      let hasMoreItems = true;
      currentsiteID=siteID;
      currentfolderpath=FolderPath;
      const container = document.getElementById("files-container");
      container.innerHTML = "";
      console.log("folderpath:", FolderPath);
      try {
        while (hasMoreItems) {
          let response;
          if (nextLink) {
            response = await sp.web(nextLink);
          } else {
            response = await testidsub.web
              .getFolderByServerRelativePath(FolderPath)
              .files.select("Name", "Length", "ServerRelativeUrl", "UniqueId","MajorVersion","ListItemAllFields/Status","ListItemAllFields/IsDeleted").expand("ListItemAllFields")
              .top(batchSize)();
              myfolderdata = response
              console.log(response , "response")


              // const items = await sp.web.lists.getByTitle('FINTESTDOC').items
              // .filter(`Status eq 'Pending'`)();

              // console.log(items , "items ,,,,,")
          }
          // Add the current batch of files to the files array
          files = [...files, ...response as IFileInfo[]];
          // Check if there is a nextLink for more items
          if ("@odata.nextLink" in response) {
            nextLink = response["@odata.nextLink"];
          } else {
            hasMoreItems = false; // No more items, exit loop
          }
        }
        console.log("All files fetched:", files);
        // Now process the files
        // const container = document.getElementById("files-container");
        // container.innerHTML = "";
     
        const DMSEntityFileMasterList=`DMS${currentEntity}FileMaster`;
        console.log(DMSEntityFileMasterList);
        const filesData = await sp.web.lists
        .getByTitle(`${DMSEntityFileMasterList}`)
        .items.select("FileUID","IsFavourite")
        .filter(
          `IsFavourite eq 1 and CurrentUser eq '${currentUserEmailRef.current}'`
        )();
     
        // Create a map for quick lookup of IsFavourite status by FileUID
        const favouriteMap = new Map(
          filesData.map((item: any) => [item.FileUID, item.IsFavourite])
        );
     
        console.log("FavouriteMap",favouriteMap)
        console.log("Files", filesData);
     
        files.forEach((file:any) => {
          if(file.ListItemAllFields.IsDeleted
            ===          null){
                      if(file.ListItemAllFields.Status !== "Pending"){
          const truncateText = (text: string, maxLength: number) => {
            return text.length > maxLength ? text.substring(0, maxLength) + "..." : text;
          };
     
          const isFavourite = favouriteMap.get(file.UniqueId) || 0;
          const favouriteText = isFavourite ? "Unmark as Favourite" : "Mark as Favourite";
     
          // Set display properties based on favorite status
          const displayPropertyforFillFavourite = isFavourite ? "block" : "none";
          const displayPropertyforUnFillFavourite = isFavourite ? "none" : "block";
     
          const card = document.createElement("div");
          const {fileIcon} = getFileIcon(file.Name); // Function to get file icon
          card.className = "card";
          card.dataset.fileId = file.UniqueId; // Store file ID in the card element
          card.innerHTML = `   
              <div class="IMGContainer">     
              <img class="filextension" src=${fileIcon} alt="File icon"/>
                     </div>     
                        <div class="CardTextContainer">
              <p class="p1st">${truncateText(file.Name, 10)}</p>
              <p class="p3rd">${((file.Length as unknown as number) / (1024 * 1024)).toFixed(2)} MB</p>
              <div id="three-dots" class="three-dots" onclick="toggleMenu2('${file.UniqueId}', '${siteID}')">
              <span>...</span>
              </div>
                      </div>
            `;
     
          const menu = document.createElement("div");
          menu.id = `menu-${file.UniqueId}`;
          menu.className = "popup-menu";
          menu.innerHTML = `
            <ul>
              <li onclick="confirmDeleteFile('${file.UniqueId}', '${siteID}')">
              <img src=${deleteIcon} alt="Delete"/>
                          Delete
              </li>
              <li onclick="editFile('${file.UniqueId}', '${siteID}')">
              <img src=${editIcon} alt="Edit"/>
                          Audit History
              </li>
              <li onclick="PreviewFile('${file.ServerRelativeUrl}', '${siteID}' , '${docLibName}')">
              <img src=${editIcon} alt="Preview"/>
                          Preview File
              </li>
              <li id="favouriteToggle-${file.UniqueId}" onclick="toggleFavourite('${file.UniqueId}', '${siteID}')">
              <img src=${UnFillFavouriteFile} alt="Mark as Favourite" class="mark-as-favourite" style="display:${displayPropertyforUnFillFavourite};"/>
              <img src=${FillFavouriteFile} alt="Unmark as Favourite" class="unmark-as-favourite" style="display:${displayPropertyforFillFavourite};"/>
              <span class="favourite-text">${favouriteText}</span>
              </li>  
            </ul>
          `;
          card.appendChild(menu);
          container.appendChild(card);


        }}
        });
      } catch (error) {
        console.error("Error fetching Doclib data:", error);
      }
    };
    // Helper function to determine the file icon based on file extension
    // const getFileIcon = (fileName: string) => {
    //   console.log(fileName , "filenmae")
    //   const fileExtension = fileName.split(".").pop().toLowerCase();
    //   switch (fileExtension) {
    //     case "doc":
    //     case "docx":
    //       return require("../assets/DOC.png");
    //     case "txt":
    //       return require("../assets/TXT.png");
    //     case "pdf":
    //       return require("../assets/PDF.png");
    //     case "xls":
    //     case "xlsx":
    //       return require("../assets/XLS.png");
    //     case "zip":
    //       return require("../assets/ZIP.png");
    //     default:
    //       return require("../assets/DOC.png");
    //   }
    // };
  
   // This function give the File Icon
   const getFileIcon = (fileName:any) => {
         
     
    const fileExtension = fileName.split(".").pop().toLowerCase();
    let fileIcon;
    switch (fileExtension) {
      case "doc":
      case "docx":
        fileIcon = require("../assets/DOC.png");
        break;
      case "txt":
        fileIcon = require("../assets/TXT.png");
        break;
      case "pdf":
        fileIcon = require("../assets/PDF.png");
        break;
      case "xls":
      case "xlsx":
        fileIcon = require("../assets/XLS.png");
        break;
      case "zip":
        fileIcon = require("../assets/ZIP.png");
        break;
      default:
        fileIcon = require("../assets/DOC.png"); // Default icon if no match
        break;
    }
    return {fileIcon,fileExtension};
  };
  window.PreviewFile = function(path :any , SiteID:any , docLibName:any , filemasterlist:any , filepreview:any){
    
    if(filepreview !== undefined || null ){
      alert(filepreview)
      const createpreviewdiv = document.createElement('div')
    createpreviewdiv.style.display = 'grid'
    const previewfileframe = document.createElement('iframe') 
    previewfileframe.id = 'filePreview'
    previewfileframe.style.width = '930px'
    previewfileframe.style.height = '500px'
    const librarydiv= document.getElementById('files-container')
    const createbutton = document.createElement('button')
    createbutton.textContent = 'Back To DMS';
   console.log("enter here in preview : ",path)
   const encodedFilePath = encodeURIComponent(path);
  console.log(encodedFilePath, "encodedFilePath");
  
  // Extract the parent folder correctly
  const parentFolder = path.substring(0, path.lastIndexOf('/'));
  console.log(parentFolder, "parentFolder");
  
  // Correctly encode the parent folder
  const encodedParentFolder = encodeURIComponent(parentFolder);
  
  // Get the base site URL
  const siteUrl = window.location.origin;
  console.log(siteUrl, "siteUrl");
  
  console.log(path , ".....path")
  // Generate the correct preview URL
  // const previewUrl = `${siteUrl}/sites/AlRostmani/${currentEntity}/${myactualdoclib}/Forms/AllItems.aspx?id=${path}&parent=${encodedParentFolder}`;
  
  // console.log(previewUrl, "Generated preview URL");
   
    // console.log("Generated Preview URL:", previewUrl);
    
      librarydiv.innerHTML = "";
      previewfileframe.src = filepreview;
      createpreviewdiv.appendChild(createbutton)
      createpreviewdiv.appendChild(previewfileframe);
      librarydiv.appendChild(createpreviewdiv)
      createbutton.addEventListener('click', function() {
        event.preventDefault()
        event.stopPropagation()
        alert('Button was clicked!');
        myRequest()
    });
    
    }
    if(filepreview == undefined || null ){
      console.log(path , "path filepreview")
      console.log(SiteID , "SiteID filepreview")
      console.log(docLibName , "docLibName filepreview")
      console.log(filemasterlist , "filemasterlist filepreview")
      console.log(filepreview , "filepreview filepreview")
      const segments = path.split('/');
    
      // Find the index of 'sites'
      const sitesIndex = segments.indexOf('sites');
    
      // If 'sites' is found and there are enough segments after it
      let myactualdoclib
      if (sitesIndex !== -1 && segments.length > sitesIndex + 3) {
        myactualdoclib = segments[sitesIndex + 3];
        console.log(myactualdoclib , "myactualdoclib")
        // return segments[sitesIndex + 3];  // The document library is the 4th segment after 'sites'
      } else {
        // return null;  // Return null if not enough segments are available
      }
      event.preventDefault()
      event.stopPropagation()
      const createpreviewdiv = document.createElement('div')
      createpreviewdiv.style.display = 'grid'
      const previewfileframe = document.createElement('iframe') 
      previewfileframe.id = 'filePreview'
      previewfileframe.style.width = '930px'
      previewfileframe.style.height = '500px'
      const librarydiv= document.getElementById('files-container')
      const createbutton = document.createElement('button')
      createbutton.textContent = 'Back To DMS';
     console.log("enter here in preview : ",path)
     const encodedFilePath = encodeURIComponent(path);
    console.log(encodedFilePath, "encodedFilePath");
    
    // Extract the parent folder correctly
    const parentFolder = path.substring(0, path.lastIndexOf('/'));
    console.log(parentFolder, "parentFolder");
    
    // Correctly encode the parent folder
    const encodedParentFolder = encodeURIComponent(parentFolder);
    
    // Get the base site URL
    const siteUrl = window.location.origin;
    console.log(siteUrl, "siteUrl");
    
    console.log(path , ".....path")
    // Generate the correct preview URL
    const previewUrl = `${siteUrl}/sites/AlRostmani/${currentEntity}/${myactualdoclib}/Forms/AllItems.aspx?id=${path}&parent=${encodedParentFolder}`;
    
    console.log(previewUrl, "Generated preview URL");
     
      console.log("Generated Preview URL:", previewUrl);
      if(previewUrl){
        librarydiv.innerHTML = "";
        previewfileframe.src = previewUrl;
        createpreviewdiv.appendChild(createbutton)
        createpreviewdiv.appendChild(previewfileframe);
        librarydiv.appendChild(createpreviewdiv)
        createbutton.addEventListener('click', function() {
          event.preventDefault()
          event.stopPropagation()
          alert('Button was clicked!');
          getdoclibdata(currentfolderpath , currentsiteID , currentDocumentLibrary)
      });
      }
    }
   
  }
  // For getting the folder data 
  // const getfolderdata = async (FolderPath:any, siteID:any) => {
  //   console.log("enter here");
  //   // event.preventDefault();
  //   // event.stopPropagation();
  //   currentsiteID=siteID;
  //   currentfolderpath=FolderPath;
  //   //created subsite context
  //   const testidsub = await sp.site.openWebById(siteID)
  //   console.log("Inside Folder directory",testidsub);
  //   const container = document.getElementById("files-container");
  //   container.innerHTML = "";
  //   try {
  
  //     //   const actualpath = `/sites/AlRostmani${FolderPath}`;
  //     //   const folder = await sp.web.getFolderByServerRelativePath(actualpath).files();
  //       const folder = await testidsub.web.getFolderByServerRelativePath(FolderPath).files();
  //       console.log(folder, "folder", typeof(folder), "type of folder");
  //       myfolderdata = folder;
        
  //       console.log(myfolderdata, "myfolderdata");
  
  //       for (const file of folder) {
  //           const fileItem = await testidsub.web.getFileByServerRelativePath(file.ServerRelativeUrl)();
  //           const name = file.Name;
  //           const filesize:any = fileItem.Length;
  //           const Actualfilesize = (filesize / (1024 * 1024)).toFixed(2);
  //           const fileid= file.UniqueId
  //           console.log(name, Actualfilesize, "name and file size");
  
  //           const card = document.createElement("div");
  //           card.className = "card";
  
  //           const Docicon = require("../assets/DOC.png");
  //           const Txticon = require("../assets/TXT.png");
  //           const Pdficon = require("../assets/PDF.png");
  //           const Xlsicon = require("../assets/XLS.png");
  //           const Zipicon = require("../assets/ZIP.png");
  //           let fileIcon;
  //           const fileExtension = name.split(".").pop().toLowerCase(); // Get the file extension
  
  //           switch (fileExtension) {
  //               case "doc":
  //               case "docx":
  //                   fileIcon = Docicon;
  //                   break;
  //               case "txt":
  //                   fileIcon = Txticon;
  //                   break;
  //               case "pdf":
  //                   fileIcon = Pdficon;
  //                   break;
  //               case "xls":
  //               case "xlsx":
  //                   fileIcon = Xlsicon;
  //                   break;
  //               case "zip":
  //                   fileIcon = Zipicon;
  //                   break;
  //               default:
  //                   fileIcon = Docicon; // Default icon if no match
  //                   break;
  //           }
  
  //           card.innerHTML = `
  //               <img class="filextension" src=${fileIcon} alt="${fileExtension} icon"/>
  //               <p class="p1st">${name}</p>
  //               <p class="p2nd"></p>
  //               <p class="p3rd">${Actualfilesize} MB</p>
  //               <div class="three-dots" onclick="toggleMenu2('${fileid}', '${siteID}')">
  //                   <span>...</span>
  //               </div>
  //           `;
  //           const menu = document.createElement("div");
  //           menu.id = `${`menu-${fileid}`}`;
  //           menu.className = "popup-menu";
  //           menu.innerHTML = `
  //             <ul>
  //                 <li onclick="deleteFile('${fileid}','${siteID}')">
  //                 <img src=${deleteIcon} alt="Delete"/>
  //                 Delete
  //               </li>
  //               <li onclick="editFile('${fileid}',  '${siteID}')">
  //                 <img src=${editIcon} alt="AuditHistory"/>
  //                 Audit History
  //               </li>  
  //             </ul>
  //           `;
          
  //           card.appendChild(menu);
  //           container.appendChild(card);
  //       }
  //   } catch (error) {
  //       console.error("Error fetching data:", error);
  //   }
  // };
  
  // Search File Function
  //    const searchFiles = async (event: React.FormEvent) => {
  //     event.preventDefault();
  //     event.stopPropagation();
  //     console.log("Inside the searchFiles");
  //     const searchInput = document.getElementById('searchinput') as HTMLInputElement;
  
  //     console.log(searchInput.value, "searchInput.value");
  //     if (searchInput.value !== "" && searchInput.value !== null) {
  //         console.log(myfolderdata, "my data");
  //         let filteredFiles = myfolderdata.filter((file: any) => file.Name.toLowerCase().includes(searchInput.value.toLowerCase()));
  //         console.log(filteredFiles, "filteredFiles");
  //         const container = document.getElementById("files-container");
  //         container.innerHTML = ""; // Clear previous search results
  
  //         // Process the filtered files
  //         if (filteredFiles.length > 0) {
  //             console.log(filteredFiles, "filteredFiles");
  //             for (const file of filteredFiles) {
  //                 console.log(file.Name, "file.Name");
  //                 console.log(file.Length, "file.Length");
  //                 const Actualfilesize = (file.Length / (1024 * 1024)).toFixed(2);
  //                 const card = document.createElement("div");
  //                 const Docicon = require("../assets/DOC.png");
  //                 const Txticon = require("../assets/TXT.png");
  //                 const Pdficon = require("../assets/PDF.png");
  //                 const Xlsicon = require("../assets/XLS.png");
  //                 const Zipicon = require("../assets/ZIP.png");
  //                 let fileIcon;
  //                 const fileExtension = file.Name.split(".").pop().toLowerCase(); // Get the file extension
          
  //                 switch (fileExtension) {
  //                   case "doc":
  //                     fileIcon = Docicon;
  //                     break;
  //                   case "docx":
  //                     fileIcon = Docicon;
  //                     break;
  //                   case "txt":
  //                     fileIcon = Txticon;
  //                     break;
  //                   case "pdf":
  //                     fileIcon = Pdficon;
  //                     break;
  //                   case "xls":
  //                   case "xlsx":
  //                     fileIcon = Xlsicon;
  //                     break;
  //                   case "zip":
  //                     fileIcon = Zipicon;
  //                     break;
  //                   default:
  //                     fileIcon = Docicon; // Default icon if no match
  //                     break;
  //                 }
          
  //                 card.className = "card";
  //                 card.innerHTML = `         
  //                     <img class="filextension" src=${fileIcon} alt="${fileExtension} icon"/>
  //                     <p class="p1st">${file.Name}</p>
  //                     <p class="p2nd"></p>
  //                     <p class="p3rd">${Actualfilesize} MB</p>
  //                     <div class="three-dots" onclick="toggleMenu2('${file.UniqueId}','${currentsiteID}')">
  //                         <span>...</span>
  //                     </div>
                       
  //                 `;
  //       const menu = document.createElement("div");
  //         menu.id = `${`menu-${file.UniqueId}`}`;
  //         menu.className = "popup-menu";
  //         menu.innerHTML = `
  //           <ul>
  //           <li onclick="deleteFile('${file.UniqueId}','${currentsiteID}')">
  //               <img src=${deleteIcon} alt="Delete"/>
  //               Delete
  //             </li>
  //             <li onclick="editFile('${file.UniqueId}','${currentsiteID}')">
  //               <img src=${editIcon} alt="AuditHistory"/>
  //               Audit History
  //             </li>  
  //           </ul>
  //         `;
        
  //         card.appendChild(menu);
                  
  //                 container.appendChild(card);
  //             }
  //         } else {
  //             console.log("No file found with the name:", searchInput.value);
  //         }
  //     } else {
  //         console.log("outttt");
  //     }
  // };
//   const searchFiles = async (event: React.FormEvent ) => {
//     event.preventDefault();
//     event.stopPropagation();
  
//     const searchInput = document.getElementById('searchinput') as HTMLInputElement;
//     const searchText = searchInput.value;
//     console.log(searchText , "searchText")
//   //  console.log(currentsiteID , "currentsiteID")
//     // const webInfo = await sp.site.openWebById(currentsiteID);
//     // console.log(webInfo , "webinfo")
//     // console.log("WebId: ", (webInfo as any).Id);
//     // const folder = await sp.web.getFolderByServerRelativePath(currentfolderpath)();
//     fetch("/sites/AlRostmani/AARG/_api/web", {
//       method: "GET",
//       headers: {
//         "Accept": "application/json;odata=verbose"
//       }
//     })
//       .then(response => response.json())
//       .then(data => {
//         const webId = data.d.Id;
//         console.log("WebId: ", webId);
//       })
//     // Get the folder ID and other relevant properties
//     // const folderDetails = {
//     //     name: folder.Name,
//     //     uniqueId: folder.UniqueId,  // Unique ID of the folder
//     //     itemCount: folder.ItemCount,
//     //     serverRelativeUrl: folder.ServerRelativeUrl
//     // };
//     // const FolderUID :any = folderDetails.uniqueId
//     // console.log(folderDetails.uniqueId , "folderDetails.uniqueId ")
//     // console.log(`https://officeindia.sharepoint.com/${currentfolderpath}` , "path")
//     const site = await sp.site.getContextInfo()
//     console.log(site , "site")
//     const site2 = await sp.site.getRootWeb()
//     console.log(site2 , "site2")
//     const currentsiteID2 = "338f2337-8cbb-4cd1-bed1-593e9336cd0e"; // siteId of the site collection
// const currentWebId = "c77461a3-065c-47b7-92f2-21fbcf443806"; // webId of the subsite
// const FolderUID = "5358a5c7-69de-4876-9fb9-10de04322671"; // ListId of the document library
// const currentfolderpath = "/sites/AlRostmani/AARG/IOCSignedDocument"; // server-relative path of the folder

//     if (searchText !== "" ) {
//         try {
//           console.log(currentfolderpath, "currentfolderpath")
//             const searchQuery = {
//                   // Querytext: `"${searchText}"`, 
                
//                   Querytext:`${searchText} AND (siteId:${currentsiteID2}) AND (webId:${currentsiteID}) AND (ListId:${FolderUID}) AND (path:"https://officeindia.sharepoint.com/${currentfolderpath}" OR ParentLink:"https://officeindia.sharepoint.com/${currentfolderpath}*")`, 
//                 // Querytext:`"${searchText}" AND ParentLink:"https://officeindia.sharepoint.com${currentfolderpath}"`,
//                 RowLimit: 500,
//                 SelectProperties: ["Title", "Path", "FileExtension", "UniqueId", "Size", "Created", "Modified"],  // Additional file properties
//                 // Refiners: 'FileExtension',
//                 // RefinementFilters: ['FileExtension:equals("docx")', 
//                 //                     'FileExtension:equals("pdf")', 
//                 //                     'FileExtension:equals("pptx")',
//                 //                   ],  
//                 // TrimDuplicates: false
//             };
//             // Performing the search
//             const searchResults = await sp.search(searchQuery);
//             const files = searchResults.PrimarySearchResults;
            
            
//             // console.log("routeToDiffSideBar",routeToDiffSideBar);
  
//             console.log(files, "files");
//             // Clear the previous results
//             const container = document.getElementById("files-container");
//             container.innerHTML = "";
  
//             // Display the search results
//             // start
//           if( routeToDiffSideBar === "" ){
//                 files.forEach((file: any) => {
//                     const card = document.createElement("div");
//                     const {fileIcon} = getFileIcon(file.Title);  
//                     card.className = "card";
//                     card.dataset.fileId = file.UniqueId; 
//                     // console.log(file.UniqueId , "file.UniqueId")
//                     card.innerHTML = `
//                           <div class="IMGContainer">
                    
//                         <img class="filextension" src=${fileIcon} alt="File icon"/>
//                                  </div>   
//                                    <div class="CardTextContainer">
//                         <p class="p1st">${file.Title}</p>
//                         <p class="p3rd">${((file.Size as unknown as number) / (1024 * 1024)).toFixed(2)} MB</p>
//                         <div id="three-dots" class="three-dots" onclick="toggleMenu2('${file.UniqueId}', '${currentsiteID}')">
//                           <span>...</span>
//                         </div>
//                            </div>
//                     `;
//                     const menu = document.createElement("div");
//                     menu.id = `menu-${file.UniqueId}`;
//                     menu.className = "popup-menu";
//                     menu.innerHTML = `
//                       <ul>
//                         <li onclick="confirmDeleteFile('${file.UniqueId}', '${currentsiteID}')">
//                           <img src=${deleteIcon} alt="Delete"/>
//                           Delete
//                         </li>
//                         <li onclick="editFile('${file.UniqueId}', '${currentsiteID}')">
//                           <img src=${editIcon} alt="Edit"/>
//                           Audit History
//                         </li>
//                         <li onclick="PreviewFile('${file.ServerRelativeUrl}', '${currentsiteID}' , '${currentDocumentLibrary}')">
//                           <img src=${editIcon} alt="Preview"/>
//                           Preview File
//                         </li>
//                         <li id="favouriteToggle-${file.UniqueId}" onclick="toggleFavourite('${file.UniqueId}', '${currentsiteID}')">
//                           <img src=${UnFillFavouriteFile} alt="Mark as Favourite" class="mark-as-favourite"/>
//                           <img src=${FillFavouriteFile} alt="Unmark as Favourite" class="unmark-as-favourite" style="display:none;"/>
//                           <span class="favourite-text">Mark as Favourite</span>
//                         </li>  
//                       </ul>
//                     `;
              
//                     card.appendChild(menu);
//                     container.appendChild(card);
//                 });
//           }else{
//               if( routeToDiffSideBar === "myRequest" ){
//                   myRequest(null,null,searchInput);
//               }
              
//               if( routeToDiffSideBar === "myFavourite" ){
  
//                     // console.log("myFavourite");
//                     myFavorite(null,null,searchInput);
                    
//               }
//               if( routeToDiffSideBar === "myFolder"){
//                     // console.log("Inside search => myFolder");
//                     mycreatedfolders(event,searchInput);
//               }
//           }
//           // end
//         } catch (error) {
//             console.error("Error searching files: ", error);
//         }
//     }
  
  
//   };
const searchFiles = async (event: React.FormEvent ) => {
  event.preventDefault();
  event.stopPropagation();

  const searchInput = document.getElementById('searchinput') as HTMLInputElement;
  const searchText = searchInput.value;
  console.log(searchText , "searchText")
 
 
  // if(currentFolder === ""){
  //   console.log("currentFolder --->",currentFolder);
  //   console.log("currentDocumentLibrary --->",currentDocumentLibrary)
  // }else{
  //   console.log("Inside else");
  //   console.log("currentFolder --->",currentFolder);
  //   console.log("currentDocumentLibrary --->",currentDocumentLibrary)
  // }
  // const testidsub = await sp.site.openWebById(currentsiteID);
  // const library = await testidsub.web.lists.getByTitle(currentDocumentLibrary).select("Id")();
  // console.log("Library",library);
  // console.log(`Document Library ID: ${library.Id}`);

  // const folder = await testidsub.web.getFolderByServerRelativePath(`${currentfolderpath}`).select("UniqueId")();
  // console.log(`Folder ID: ${folder.UniqueId}`);
  // console.log("Folder",folder);


  if (searchText !== "" ) {
      try {
        console.log(currentfolderpath, "currentfolderpath")
          const searchQuery = {
               Querytext:`"${searchText}" AND ParentLink:"https://officeindia.sharepoint.com${currentfolderpath}"`,
              // Querytext: `"${searchText}"`,
              RowLimit: 500,
              SelectProperties: ["Title", "Path", "FileExtension", "UniqueId", "Size", "Created", "Modified"],  // Additional file properties
              // Refiners: 'FileExtension',
              // RefinementFilters: ['FileExtension:equals("docx")',
              //                     'FileExtension:equals("pdf")',
              //                     'FileExtension:equals("pptx")',
              //                   ],  
              // TrimDuplicates: false
          };
          // Performing the search
          const searchResults = await sp.search(searchQuery);
          const files = searchResults.PrimarySearchResults;
         
         
          // console.log("routeToDiffSideBar",routeToDiffSideBar);

          console.log(files, "files");
          // Clear the previous results
          const container = document.getElementById("files-container");
          container.innerHTML = "";

          // Display the search results
          // start
        if( routeToDiffSideBar === "" ){
              files.forEach((file: any) => {
                  const card = document.createElement("div");
                  const {fileIcon} = getFileIcon(file.Title);  
                  card.className = "card";
                  card.dataset.fileId = file.UniqueId;
                  // console.log(file.UniqueId , "file.UniqueId")
                  card.innerHTML = `
                      <img class="filextension" src=${fileIcon} alt="File icon"/>
                      <p class="p1st">${file.Title}</p>
                      <p class="p3rd">${((file.Size as unknown as number) / (1024 * 1024)).toFixed(2)} MB</p>
                      <div id="three-dots" class="three-dots" onclick="toggleMenu2('${file.UniqueId}', '${currentsiteID}')">
                        <span>...</span>
                      </div>
                  `;
                  const menu = document.createElement("div");
                  menu.id = `menu-${file.UniqueId}`;
                  menu.className = "popup-menu";
                  menu.innerHTML = `
                    <ul>
                      <li onclick="confirmDeleteFile('${file.UniqueId}', '${currentsiteID}')">
                        <img src=${deleteIcon} alt="Delete"/>
                        Delete
                      </li>
                      <li onclick="auditHistory('${file.UniqueId}', '${currentsiteID}','${file.Title}')">
                        <img src=${editIcon} alt="Edit"/>
                        Audit History
                      </li>
                      <li onclick="PreviewFile('${file.ServerRelativeUrl}', '${currentsiteID}' , '${currentDocumentLibrary}')">
                        <img src=${editIcon} alt="Preview"/>
                        Preview File
                      </li>
                      <li id="favouriteToggle-${file.UniqueId}" onclick="toggleFavourite('${file.UniqueId}', '${currentsiteID}')">
                        <img src=${UnFillFavouriteFile} alt="Mark as Favourite" class="mark-as-favourite"/>
                        <img src=${FillFavouriteFile} alt="Unmark as Favourite" class="unmark-as-favourite" style="display:none;"/>
                        <span class="favourite-text">Mark as Favourite</span>
                      </li>  
                    </ul>
                  `;
           
                  card.appendChild(menu);
                  container.appendChild(card);
              });
        }else{
            if( routeToDiffSideBar === "myRequest" ){
                myRequest(null,null,searchInput);
            }
           
            else if( routeToDiffSideBar === "myFavourite" ){

                  // console.log("myFavourite");
                  myFavorite(null,null,searchInput);
                 
            }
            else if( routeToDiffSideBar === "myFolder"){
                  // console.log("Inside search => myFolder");
                  mycreatedfolders(event,searchInput);
            }
            else if(routeToDiffSideBar === "shareWithOthers"){
                ShareWithOther(null,searchInput);
            }
            else if(routeToDiffSideBar === "shareWithMe"){
                ShareWithMe(null,searchInput);
            }
        }
        // end
      } catch (error) {
          console.error("Error searching files: ", error);
      }
  }


};

// Share With ME & Share With Others ///
const ShareWithOther=async(event:React.MouseEvent<HTMLButtonElement>=null,searchText:HTMLInputElement=null)=>{
  const wait = document.getElementById('files-container')
  wait.classList.remove('hidemydatacards')
  const hidegidvewlistviewbutton = document.getElementById('hidegidvewlistviewbutton')
  if (hidegidvewlistviewbutton) {
   console.log("enter here .....................")
   hidegidvewlistviewbutton.style.display = 'none'
  
 }
  if(createFileButton2){
    createFileButton2.style.display = 'none'
    }
    if(createFileButton){
    createFileButton.style.display = 'none'
    }
  if(event){
    event.preventDefault();
    event.stopPropagation();
  }
  console.log("Share with others called");
  console.log("searchInput",searchText);

  const container = document.getElementById("files-container");
  container.innerHTML="";

  const FilesItems = await sp.web.lists
  .getByTitle("MasterSiteURL")
  .items.select("Title", "SiteID", "FileMasterList", "Active")
  .filter(`Active eq 'Yes'`)();

  // console.log("Files items", FilesItems);
  FilesItems.forEach(async(fileItem)=>{
    if(fileItem.FileMasterList !== null){
      // console.log(files.FileMasterList);

      const filesData = await sp.web.lists
      .getByTitle(`${fileItem.FileMasterList}`)
      .items.select("FileName", "FileUID", "FileSize", "FileVersion","ShareWithOthers")
      .filter(
        `CurrentUser eq '${currentUserEmailRef.current}'`
      )();

     
      console.log("Files Data ",filesData);
      routeToDiffSideBar="shareWithOthers"
      let filteredFileData=[];
      if(searchText !== null){
            filteredFileData=filesData.filter((file: any) => file?.FileName?.toLowerCase().includes(searchText?.value?.toLowerCase()))
      }else{
        filteredFileData=filesData;
      }
      filteredFileData.forEach((file) => {

        if( file.ShareWithOthers !== null ){
         
          const sharedUserInTheFormOFstring = file.ShareWithOthers;
        
          let sharedUsers = JSON.parse(sharedUserInTheFormOFstring);
          console.log(sharedUsers , " here is shared users")
          if(sharedUsers.length === 0){
              return;
          }
         
          // Get the first two users
          const firstTwoUsers = sharedUsers.slice(0, 2);

          // Remaining users count
          const moreUsersCount = sharedUsers.length - 2;

          // Create shared users HTML for the first two users
          let sharedUsersHTML = firstTwoUsers
              .map((user:any) => {
                  let firstNameInitial;
                  console.log(user , "sharewith me users")
                  console.log("user firstnamw", user.SharedWith)
                  console.log("user lastnamw",user.LastName)
                  let lastNameInitial=""
                  if(user.FirstName !== null){
                        firstNameInitial = user.FirstName.charAt(0).toUpperCase();
                  }
                  if(user.LastName !== null){
                        lastNameInitial=user.LastName.charAt(0).toUpperCase();
                  }

                  return `<span  flow="down" tooltip='${user.FirstName }' class="shared-user">${firstNameInitial}${lastNameInitial}</span>`;
                  })
                  .join("");

               let array = ["test1" , "test2" , "test3" , "test4"]
               console.log(array , "array")
          // If there are more users, add "+more"
          if (moreUsersCount > 0) {
                sharedUsersHTML += `<span class="more-users" flow="down" tooltip='${array }'>+${moreUsersCount} more</span>`;
          }
         
          const {fileIcon, fileExtension}= getFileIcon(file.FileName);
          // const card = createFileCard(file, fileIcon, fileItem.SiteID,fileItem.FileMasterList,fileExtension);
          const card = document.createElement("div");
          card.className = "card";
          card.dataset.fileId = file.FileUID; // Store file ID in the card element
          card.dataset.listId = fileItem.SiteID; // Store site ID
       
          card.innerHTML = `        
            <img class="filextension" src=${fileIcon} alt="${fileExtension} icon"/>
            <p class="p1st">${file.FileName}</p>
            <div class="fileSizeAndVersion">
              <p class="p3rd">${file.FileSize} MB</p>
              <p class="p2nd">${file.FileVersion}</p>
            </div>
             <div class="sharedFile">
              ${sharedUsersHTML}
            </div>
          `;
          container.appendChild(card);

        }
      });
     
    }        
  })

}

const ShareWithMe=async(event:React.MouseEvent<HTMLButtonElement>=null,searchText:HTMLInputElement=null)=>{
  const wait = document.getElementById('files-container')
  wait.classList.remove('hidemydatacards')
  if(createFileButton2){
    createFileButton2.style.display = 'none'
    }
    if(createFileButton){
    createFileButton.style.display = 'none'
    }
    const hidegidvewlistviewbutton = document.getElementById('hidegidvewlistviewbutton')
    if (hidegidvewlistviewbutton) {
     console.log("enter here .....................")
     hidegidvewlistviewbutton.style.display = 'none'
    
   }
  if(event){
    event.preventDefault();
    event.stopPropagation();
  }      
  console.log("Share with me called");
  console.log("searchInput",searchText);
  const container = document.getElementById("files-container");
  container.innerHTML="";

  const FilesItems = await sp.web.lists
  .getByTitle("MasterSiteURL")
  .items.select("Title", "SiteID", "FileMasterList", "Active")
  .filter(`Active eq 'Yes'`)();
 
  console.log("MasterSite Items",FilesItems);

  FilesItems.forEach(async(fileItem)=>{

    if(fileItem.FileMasterList !== null){
      // console.log(files.FileMasterList);

      const filesData = await sp.web.lists
      .getByTitle(`${fileItem.FileMasterList}`)
      .items.select("FileName", "FileUID", "FileSize", "FileVersion","ShareWithMe")
      .filter(
        `CurrentUser ne '${currentUserEmailRef.current}'`
      )();

     
      console.log("Files Data ",filesData);
      routeToDiffSideBar="shareWithMe";
      let filteredFileData=[];
      if(searchText !== null){
            filteredFileData=filesData.filter((file: any) => file?.FileName?.toLowerCase().includes(searchText?.value?.toLowerCase()))
      }else{
        filteredFileData=filesData;
      }
      filteredFileData.forEach((file)=>{
        if( file.ShareWithMe !== null ){
          // console.log("FilesMe",file)
          const sharedUserInTheFormOFstring = file.ShareWithMe;
     
          let sharedUsers = JSON.parse(sharedUserInTheFormOFstring);

          if(sharedUsers.length === 0){
              return;
          }

          let fileShareWithMe=sharedUsers.find( (item:any) =>
            {
              //  console.log(item.SharedWith);
              //  console.log("current User",currentUserEmailRef.current)
               return item.SharedWith === currentUserEmailRef.current
            }
          )

          // console.log("files share with me =>",fileShareWithMe);


          if( fileShareWithMe !== undefined ){

            console.log("This File is Share With me By Other Users",file.FileName);
           
            const {fileIcon, fileExtension}= getFileIcon(file.FileName);

            const card = document.createElement("div");
            card.className = "card";
            card.dataset.fileId = file.FileUID; // Store file ID in the card element
            card.dataset.listId = fileItem.SiteID; // Store site ID
         
            card.innerHTML = `        
              <img class="filextension" src=${fileIcon} alt="${fileExtension} icon"/>
              <p class="p1st">${file.FileName}</p>
              <div class="fileSizeAndVersion">
                <p class="p3rd">${file.FileSize} MB</p>
                <p class="p2nd">${file.FileVersion}</p>
              </div>`;

            container.appendChild(card);

          }

        }
      })

    }

})

}
// This function is called when we click on the share.
window.shareFile=async(fileID:String,siteId:String,currentFolderPathForFile:string,fileName:string)=>{
  console.log("Share File called");
  console.log("file Id",fileID);
  console.log("site Id",siteId);
  console.log("currentFolderPath",currentFolderPathForFile);

  // exreact the Entity from folder path
  const parts = currentFolderPathForFile.split("/");
  const entity = parts[3];
  console.log(entity);

  const fetchUser=async(entity:any)=>{
    const [
      users,
      users1,
      users2,
      users3,
      users4,
    ] = await Promise.all([
      sp.web.siteGroups.getByName(`${entity}_Read`).users(),
      sp.web.siteGroups.getByName(`${entity}_Initiator`).users(),
      sp.web.siteGroups.getByName(`${entity}_Contribute`).users(),
      sp.web.siteGroups.getByName(`${entity}_Admin`).users(),
      sp.web.siteGroups.getByName(`${entity}_View`).users(),
    ]);
    console.log(users, "users ", users1,users2,users3,users4);
    const combineArray = [
      ...(users || []),
      ...(users1 || []),
      ...(users2 || []),
      ...(users3 || []),
      ...(users4 || []),
    ];
     
    const resultArray=combineArray.map((user) => (
      {
        id:String(user.Id),
        value: user.Title,
        email: user.Email
      }
    ))
    console.log("combineArray", combineArray);
    console.log("resultArray",resultArray)

    return resultArray;
  }

  const users=await fetchUser(entity);
  console.log("UserArray",users);
 

// Check if a popup already exists, if so, remove it before creating a new one
const existingPopup = document.getElementById('share-popup');
if (existingPopup) {
existingPopup.remove();
}

// Dummy data
// const users = [
//   { value: 'Test1', id: '14',email:"User1@officeindia.onmicrosoft.com" },
//   { value: 'Test2', id: '31',email:"User2@officeindia.onmicrosoft.com" },
//   { value: 'Test3', id: '137',email:"User3@officeindia.onmicrosoft.com"},
//   { value: 'Test4', id: '33',email:"User4@officeindia.onmicrosoft.com" },
//   { value: 'Test5', id: '32',email:"User5@officeindia.onmicrosoft.com" },
//   { value: 'Test6', id: '34',email:"User6@officeindia.onmicrosoft.com" },
//   { value: 'Test User1', id: '39',email:"User7@officeindia.onmicrosoft.com" },
//   ];


// Declare selectedUsers with an explicit type, assuming user IDs are of type string for selecting the user for share
let selectedUsers: { id: string; value: string; email:string }[] = [];
// Create the pop-up element
const popup = document.createElement("div");
popup.id = 'share-popup';
popup.className = "share-popup";

// Add HTML structure for the pop-up with a dropdown and a close "X" button
popup.innerHTML = `
<div class="share-popup-content">
<div class="share-popup-header">
  <h4>Share</h4>
  <span class="share-close-popup" onClick="hideSharePopUp()">x</span>
</div>
<div class="share-popup-body">
  <div id="share-reactSelect">
      <input type="text" id="userInput" placeholder="Add a Name, Group, or Email" style="
      width: 100%;
      padding: 10px;
      font-size: 14px;
      border-radius: 4px;
      border: 1px solid #ccc;
    "/>
    <div id="userDropdown" class="user-dropdown" style="
      display: none;
      position: absolute;
      width: 29.8%;
      max-height: 150px;
      overflow-y: auto;
      background-color: white;
      border: 1px solid #ccc;
      border-radius: 4px;
      z-index: 1000;
    ">
    </div>
  </div>
  <textarea id="share-message" placeholder="Write a message..." >
  </textarea>
</div>
<div class="share-popup-footer">
  <button id="share-shareFileButton">Share</button>
</div>
</div>
`;

// Append the  popup to the body
document.body.appendChild(popup);

// Get references to the input box and dropdown
const userInput = document.getElementById('userInput') as HTMLInputElement;
const userDropdown = document.getElementById('userDropdown');

// Function to render dropdown options based on user input
function renderDropdown(users: { id: string, value: string,email:string }[]) {
// Clear previous options
userDropdown.innerHTML = '';
users.forEach(user => {
const option = document.createElement('div');
option.className = 'dropdown-item';
option.style.padding = '8px';
option.style.cursor = 'pointer';
option.textContent = user.value;
option.onclick = () => selectUser(user);
userDropdown.appendChild(option);
});
}

// Function to show the dropdown when the input is clicked
userInput.addEventListener('focus', () => {
userDropdown.style.display = 'block';

// Display all users initially
renderDropdown(users);
});

// Filter dropdown based on input value
userInput.addEventListener('input', () => {
const searchValue = userInput.value.toLowerCase();
const filteredUsers:any = users.filter(user => user.value.toLowerCase().includes(searchValue));
renderDropdown(filteredUsers);
});

// Function to select a user and display it inside the input
function selectUser(user: { id: string, value: string,email:string }) {
console.log("selected user",selectedUsers)
if (!selectedUsers.some(selectedUser => selectedUser.id === user.id)) {

selectedUsers.push(user);

// Create a span for the selected user with a close button
const selectedUserDiv = document.createElement('span');
selectedUserDiv.className = 'selected-user';
selectedUserDiv.style.display = 'inline-block';
selectedUserDiv.style.padding = '2px 6px';
selectedUserDiv.style.backgroundColor = '#e0e0e0';
selectedUserDiv.style.borderRadius = '12px';
selectedUserDiv.style.marginRight = '5px';
selectedUserDiv.style.position = 'relative';

selectedUserDiv.textContent = user.value;

// Create close button for deselecting the user
const closeButton = document.createElement('span');
closeButton.textContent = 'x';
closeButton.style.cursor = 'pointer';
closeButton.style.marginLeft = '5px';
closeButton.onclick = () => deselectUser(user.id, selectedUserDiv);
selectedUserDiv.appendChild(closeButton);

// Append the selected user to the input field
userInput.parentNode!.insertBefore(selectedUserDiv, userInput);
userInput.value = '';
}
userDropdown.style.display = 'none';
}

// Function to deselect a user
function deselectUser(userId: string, selectedUserDiv: HTMLElement) {
// selectedUsers = selectedUsers.filter(id => id !== userId);
selectedUsers = selectedUsers.filter(selectedUser => selectedUser.id !== userId);
console.log("selected user",selectedUsers);
selectedUserDiv.remove();
}

// Hide the dropdown if clicked outside
document.addEventListener('click', (event) => {
if (!userInput.contains(event.target as Node) && !userDropdown.contains(event.target as Node)) {
userDropdown.style.display = 'none';
}
});

// Adding event listener to the "Share" button
document.getElementById('share-shareFileButton').addEventListener('click', async function() {
    console.log("selectedUserArray",selectedUsers);
    console.log("Entity",entity);
    console.log("FileId",fileID)
    console.log("SiteId",siteId);

    const listToUpdateWithShareData=`DMS${entity}FileMaster`;
    console.log("listToUpdateWithShareData",listToUpdateWithShareData);

    // Fetch the item from the list using its ID
    const item = await sp.web.lists.getByTitle(listToUpdateWithShareData).items.select("FileName","ShareWithOthers","ShareWithMe","FileUID","ID").filter(`FileUID eq '${fileID}' and CurrentUser eq '${currentUserEmailRef.current}'`)();

    console.log("item",item);

    // let dataArray;
    let dataArray: Array<{ FirstName: string; LastName?: string; SharedWith: string; SharedAt: string; TimeStamp: number; Permission: string,userId:string }> = [];
         
    selectedUsers.forEach(async(user)=>{
   
    const nameParts = user.value.trim().split(" ");
    const firstName = nameParts[0];
    let lastName = nameParts.length > 1 ? nameParts[nameParts.length - 1] : "";
    console.log("firstName",firstName)
    console.log("lastName",lastName);
    if(lastName === ""){
      lastName="";
    }

    const isoDate = new Date().toISOString().slice(0, 19) + 'Z';
    const timestamp = Date.now();
      let userObj={
        FirstName:firstName,
        LastName:lastName,
        SharedWith:user.email,
        SharedAt:isoDate,
        TimeStamp:timestamp,
        Permission:"Read",
        userId:user.id
      }
      dataArray.push(userObj);
      // console.log("userObj",userObj);
    })

    console.log("dataArray",dataArray);

   

    if(item[0].ShareWithMe === null && item[0].ShareWithOthers === null){

          const dataInTheFormoOfString=JSON.stringify(dataArray);
           // Now update specific columns of the item
            const updatedItem = await sp.web.lists.getByTitle(listToUpdateWithShareData).items.getById(item[0].ID).update({
              ShareWithOthers:dataInTheFormoOfString,
              ShareWithMe:dataInTheFormoOfString
            });

            console.log("Data updated when ShareWithMe and ShareWithOthers are null",updatedItem);
    }else{
       const shareWithOthers =JSON.parse(item[0].ShareWithOthers);
       const shareWithMe=JSON.parse(item[0].ShareWithMe);

       dataArray.forEach((user)=>{
            // apply condition for sharing same file with same user multiple time using id of the user
            const alReadySharedUserIndex=shareWithOthers.findIndex((item:any)=>{
                  return item.userId === user.userId
            })
            console.log("alReadySharedUser in shareWithOthers",alReadySharedUserIndex);
            const alReadySharedUserIndex1=shareWithMe.findIndex((item:any)=>{
                return item.userId === user.userId
            })
            console.log("alReadySharedUser in shareWithMe",alReadySharedUserIndex1);

            if(alReadySharedUserIndex !== -1){
                  shareWithOthers.splice(alReadySharedUserIndex, 1);
                  shareWithOthers.push(user);
                  console.log("shareWithOthers",shareWithOthers);
            }else{
              shareWithOthers.push(user);
            }

            if(alReadySharedUserIndex1 !== -1){
              shareWithMe.splice(alReadySharedUserIndex1, 1);
              shareWithMe.push(user);
              console.log("shareWithMe",shareWithMe);
            }else{
              shareWithMe.push(user);
            }
       })

       console.log("shareWithOthers",shareWithOthers);
       console.log("shareWithMe",shareWithMe);

       const dataInTheFormoOfStringForShareWithMe=JSON.stringify(shareWithMe);
       const dataInTheFormoOfStringForShareWithOthers=JSON.stringify(shareWithOthers);
       // Now update specific columns of the item
       const updatedItem = await sp.web.lists.getByTitle(listToUpdateWithShareData).items.getById(item[0].ID).update({
        ShareWithOthers:dataInTheFormoOfStringForShareWithOthers,
        ShareWithMe:dataInTheFormoOfStringForShareWithMe
      });

      console.log("Data updated when ShareWithMe and ShareWithOthers",updatedItem);
    }

});


}

// hide the share popup
// @ts-ignore
window.hideSharePopUp=()=>{
const popup=document.querySelector('.share-popup');

if(popup){
  popup.remove();
}
}


// Sharewith Me And Share With Others
  //Toggle the menu card
  // @ts-ignore
   window.toggleMenu2 = function(fileId: string , siteID:any , listitemid:any , Listname:any) {
    console.log(listitemid , ": listitemid") 
    console.log(Listname , ": fileItem.FileMasterList") 
    console.log("Inside the toggleMenu2");
    console.log(siteID, "siteID")
    console.log(fileId , "fileId")
    console.log("enter here i n menu card")
    const allMenus = document.querySelectorAll('.popup-menu');
    console.log(allMenus , "allMenus")
    allMenus.forEach(menu => {
      console.log(menu , "menu")
      console.log(menu.id , "menu.id")
      console.log(fileId , "fileId")
      if (menu.id !== `menu-${fileId}`) {
        menu.classList.remove("show");
      }
    });
  
    // Toggle the menu for the clicked card
    const menu = document.getElementById(`menu-${fileId}`);
    if (menu) {
      console.log("Toggle the menu for the clicked card")
      menu.classList.toggle("show");
    }
    document.addEventListener('click', (event) => {
    
      // console.log("Outside click Event Called");
    
      const target = event.target as HTMLElement;
    
      // Check if the click was inside any menu or three-dot icon
      const isClickInsideMenu = target.closest('.popup-menu');
      const isClickInsideThreeDots = target.closest('.three-dots');
    
      // console.log("This is nested folder",isClickInsideThreeDots);
    
      if (!isClickInsideMenu && !isClickInsideThreeDots) {
        const allMenus = document.querySelectorAll('.popup-menu');
        allMenus.forEach(menu => {
          menu.classList.remove('show');
        });
      }
    });
  }
  
  
    // Edit file action
     // @ts-ignore

    
  
    // Delete file action
   // @ts-ignore
  
   /// this is pop up function
   window.confirmDeleteFile =async(fileId:string, siteID:string ,ListToUpdate:any=null)=>{
    console.log("list name is " , ListToUpdate)
    // console.log(listToUpdate , "listAnme")
    event.preventDefault();
    event.stopPropagation();
    console.log("Inside The confirmDeleteFile");
    console.log("FileID",fileId);
    console.log("siteId",siteID);
   
    const popupData = await sp.web.lists.getByTitle('DMSPopupMaster').items
    .select('PopupText', 'Typeofpopup', 'Isrequires')();
   
    console.log("popupitems",popupData);
   
   popupData.forEach(async (popItems) => {
   
      // Check the type of popup and if it is required
      switch (popItems.Typeofpopup) {
      
          case 'Delete':
              if (popItems.Isrequires) {
                console.log(popItems.Typeofpopup ,"popItems.Typeofpopup ")
                      console.log("TypeOfPopUp: Delete and Isrequires is true");
                     
                      // Create Pop
                      const deleteConfirmationPop = document.createElement('div');
                      deleteConfirmationPop.className = "popup-modal";
                      deleteConfirmationPop.innerHTML = `
                          <div class="popup-content">
                            <p id="confirmation-text">${popItems.PopupText}</p>
                            <div class="popup-actions">
                                <button id="confirm-yes">Yes</button>
                                <button id="confirm-no">No</button>
                            </div>
                          </div>
                      `;
   
                      document.body.appendChild(deleteConfirmationPop);
   
                      // Handle Yes button click (confirmation for Delete)
                      const yesButton =document.getElementById('confirm-yes');
                      yesButton.addEventListener('click', async () => {
                      const confirmationText = document.getElementById('confirmation-text');
                      confirmationText.innerHTML = 'Loading...';
                      
                      try {
                            console.log("Calling deleteFile from confirm delete");
                            await window.deleteFile(fileId, siteID,ListToUpdate);
  
                            // console.log("Updating List inside the confirem Delete");
                            // if(ListToUpdate){       
                            //         const items999 = await sp.web.lists
                            //         .getByTitle(ListToUpdate).items.filter(`FileUID eq '${fileId}'`).top(1)();
                            //         alert(items999)
                                    
                            //         if (items999.length > 0) {
                            //         const itemId = items999[0].ID;
                            //         console.log(itemId , "itemId")
                            //         console.log(items999 , "item9999")
                            //         // Delete the item by ID
                            //         const mylist = ListToUpdate
                            //         console.log(mylist, "mylist")
                            //         await sp.web.lists.getByTitle(mylist).items.getById(itemId).delete();
                        
                            //         console.log(`Item with FileUid ${fileId} has been deleted.`);
                            //         }
                                   
                        // }
                        confirmationText.innerHTML = 'Your file was deleted successfully.';
                      
                        } catch (error) {
                          confirmationText.innerHTML = 'Something went wrong. Your file was not deleted.';
                      }
   
                      // Remove the popup after 1 second
                      setTimeout(() => document.body.removeChild(deleteConfirmationPop), 1000);
                  });
   
                      // Handle No button click (cancel deletion)
                      document.getElementById('confirm-no').addEventListener('click', () => {
                          document.body.removeChild(deleteConfirmationPop); // Close the popup
                      });
   
                  } else {
                      console.log("TypeOfPopUp: Delete and Isrequires is false");
                      // Directly delete the file if no popup is required
                      try {
                          await window.deleteFile(fileId, siteID,ListToUpdate);
                          alert('Your file was deleted successfully.');
                      } catch (error) {
                          alert('Error deleting file.');
                      }
                  }
                  break;
   
          case 'CreateFile':
                  if (popItems.Isrequires) {
                      console.log("TypeOfPopUp: CreateFile and Isrequires is true");
                     
                      // Show popup for CreateFile
                      const createFileConfirmationPop = document.createElement('div');
                      createFileConfirmationPop.className = "popup-modal";
                      createFileConfirmationPop.innerHTML = `
                          <div class="popup-content">
                            <p id="confirmation-text">${popItems.PopupText}</p>
                            <div class="popup-actions">
                                <button id="confirm-yes">Yes</button>
                                <button id="confirm-no">No</button>
                            </div>
                          </div>
                      `;
   
                      document.body.appendChild(createFileConfirmationPop);
                  } else {
                        // Logic without Pop
                    }
                    break;
   
          // Add more cases here for other types like 'Edit', 'Upload', etc.
          default:
              console.log("Unknown TypeOfpopup: ", popItems.Typeofpopup);
      }
  });
   
  }
  
  
  // Without Pop-up
  // @ts-ignore
    window.deleteFile = async(fileId:string, siteID:string, ListToUpdate:any=null) => {
      console.log("Inside the deleteFile");
      console.log("ListToUpdate",ListToUpdate)
      console.log(siteID ,"siteID")
      console.log(`Delete file with ID: ${fileId}`);
      const testidsub = await sp.site.openWebById(siteID)
      
      const deleteffile =  await testidsub.web.getFileById(fileId).delete();
      console.log(deleteffile , "deleteffile")
       alert(`File with ID: ${fileId} has been deleted successfully.`);
  
       console.log(currentfolderpath , "currentfolderpath")
       console.log("currentEntity",currentEntity);
       
       //start
       if(ListToUpdate || currentEntity){
            console.log("Inside The check Of Entity->",currentEntity,"->",ListToUpdate);
            let currentList;
            if(ListToUpdate){
                currentList=ListToUpdate;
            }
            if(currentEntity){
                currentList=`DMS${currentEntity}FileMaster`;
            }
            console.log("selected List",ListToUpdate);
            const items999 = await sp.web.lists
            .getByTitle(currentList).items.filter(`FileUID eq '${fileId}'`).top(1)();
                            alert(items999)
            
            if (items999.length > 0) {
            const itemId = items999[0].ID;
        
            // Delete the item by ID
            await sp.web.lists.getByTitle(currentList).items.getById(itemId).delete();
            
            console.log(`Item with FileUid ${fileId} has been deleted.`);
            }
        }
        // end
       getdoclibdata(currentfolderpath, currentsiteID , currentDocumentLibrary)
      //  getfolderdata(currentfolderpath,currentsiteID)
    };
    
  
    
    //Manage Folder Permission Action 
  // window.managePermission=(message:string)=>{
  //   console.log(message);
  // }
  
  // Manage Folder WorkFlow Action
  
  // Manage Folder View Action
  window.view=(message:string)=>{
    console.log(message);
  }
  
  
  
  // My ctreated folder 
  const createFileButton2 = document.getElementById('createFileButton2')
  const createFileButton = document.getElementById('createFileButton')
  const mycreatedfolders = async (event:any=null, searchText:any=null )=>{
    const wait = document.getElementById('files-container')
    wait.classList.remove('hidemydatacards')
    if(createFileButton2){
      createFileButton2.style.display = 'none'
      }
      if(createFileButton){
      createFileButton.style.display = 'none'
      }
    setlistorgriddata('')
    setlistorgriddata('')
    setShowMyrequButtons(false)
    setShowMyfavButtons(false)

    if(event){
      event.preventDefault()
      event.stopPropagation()
    }
   
    // start
    // call this function onClick of the myFolder Button
    // handleShowContent(event)
    // end
    if(createFileButton2){
       createFileButton2.style.display = 'none'
    }
     if(createFileButton){
    createFileButton.style.display = 'none'
     }  
     const hidegidvewlistviewbutton = document.getElementById('hidegidvewlistviewbutton')
     if (hidegidvewlistviewbutton) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton.style.display = 'none'
     
    }
    const folderItems = await sp.web.lists
    .getByTitle("DMSFolderMaster")
    .items.select("CurrentUser" , "IsFolder" , "FolderPath" , "DocumentLibraryName","SiteTitle","ID")
    .filter(`IsLibrary eq 1 and CurrentUser eq '${currentUserEmailRef.current}'`)();
    console.log(folderItems , "folderItems");

    // new code to fetch the siteId from the masterSiteURl and map this siteid with corresponding siteTitle forEach folder in the folderData that fetch from the DMSFolderMaster
    const dataFromMasterSiteURL=await sp.web.lists.getByTitle("MasterSiteURL").items.select("Title","SiteID").filter(`Active eq 'Yes'`)();
    console.log("dataFromMasterSiteURL",dataFromMasterSiteURL);

    const siteMap=new Map();
    dataFromMasterSiteURL.forEach(site => {
      siteMap.set(site.Title, site.SiteID);
    });

    const folderDataWithSiteId= folderItems.map(folder => {
       // Get the SiteID or null if not found
      const siteID = siteMap.get(folder.SiteTitle) || null;
      return {
        ...folder,
        // Append SiteID to the folder object
        SiteID: siteID
      };
    });

    console.log("Resultant folder data",folderDataWithSiteId);
    // end new code

    const container = document.getElementById("files-container");
    container.innerHTML = "";
    const folderimg = require('../assets/Folder.png')
   
    // start
    console.log("searchInput",searchText);
    routeToDiffSideBar="myFolder";
    let filteredFileData;
    if(searchText !== null){
      // here we change the array to new siteId containing array
      filteredFileData=folderDataWithSiteId.filter((folder: any) =>
           folder.DocumentLibraryName.toLowerCase().includes(searchText.value.toLowerCase())
      // ||   folder.FolderName.toLowerCase().includes(searchText.value.toLowerCase())
      // ||   folder.ParentFolder.toLowerCase().includes(searchText.value.toLowerCase())
    )
    }else{
      // here we change the array to new siteId containing array
      filteredFileData=folderDataWithSiteId;
    }
    // end
    // change the array name in the for loop
    for(const files of filteredFileData){
      const card = document.createElement("div");
 
      card.className = "card";
      card.innerHTML = `
      <img class="filextension" src=${folderimg} icon"/>
      <p class="p1st">${files.DocumentLibraryName}</p>
      <p class="p2nd"></p>
      <p class="p3rd">${files.SiteTitle}</p>
      <div class="three-dots" onclick="toggleMenu2('${files.ID}')">
          <span>...</span>
      </div>
    `;
    const menu = document.createElement("div");
    menu.id =`menu-${files.ID}`;
    menu.className = "popup-menu";
    menu.innerHTML = `
    <ul>
         <li onclick="managePermission('${files.DocumentLibraryName}','${files.SiteTitle}','${files.SiteID}')">
          <img src=${managePermissionIcon} alt="ManagePermission"/>
          Manage Permission
      </li>
      <li onclick="manageWorkflow('${files.DocumentLibraryName}','${files.SiteTitle}','${files.SiteID}')">
        <img src=${manageWorkFlowIcon} alt="ManageWorkFlow"/>
        Manage Workflow
      </li>
      <li onclick="editFile('${files.SiteTitle}','${files.DocumentLibraryName}')">
        <img src=${editIcon} alt="Edit"/>
        Edit
      </li>
      <li onclick="view('view')">
        <img src=${viewIcon} alt="View"/>
        View
      </li>
      <li onclick="deleteFile('delete')">
        <img src=${deleteIcon} alt="Delete"/>
        Delete
      </li>  
    </ul>
    `;
   
    card.appendChild(menu);
    container.appendChild(card);
    }
   
  }
   // This Function is Called when we click on the MyFavourite
  // This Function is Called when we click on the MyFavourite
   // This Function is Called when we click on the MyFavourite
   const myFavorite= async (event: any = null, siteIdToUpdate: string = null,searchText:any=null) => {
    // // alert()
    // setlistorgriddata('')
    // setMyreqormyfav('Myfavourite')
    // // setShowButtons(true)
    // setShowMyrequButtons(false)
    // setShowMyfavButtons(true)
  
    if(createFileButton2){
      createFileButton2.style.display = 'none'
      }
      if(createFileButton){
      createFileButton.style.display = 'none'
      }
       
  
    setTimeout(() => {
      // alert("set timer")
      setlistorgriddata('');  // Update state to '' after a delay
  
      console.log(listorgriddata, "list")
    }, 100);
    const wait2 = document.getElementById('hidegidvewlistviewbutton')
    wait2.classList.add('hidemydatacards')
    const wait = document.getElementById('files-container')
    wait.classList.remove('hidemydatacards')
    setShowMyrequButtons(false)
    setShowMyfavButtons(true)
    setMyreqormyfav((previous)=>'Myfavourite')
   
    // setlistorgriddata('')
    // const hidegidvewlistviewbutton=document.getElementById("hidegidvewlistviewbutton")
    // if (hidegidvewlistviewbutton) {
    //   console.log("enter here .....................")
    //   hidegidvewlistviewbutton.style.display = 'flex'
     
    // }
  
    const hidegidvewlistviewbutton=document.getElementById("hidegidvewlistviewbutton")
    if (hidegidvewlistviewbutton) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton.style.display = 'none'
     
    }
    const hidegidvewlistviewbutton2=document.getElementById("hidegidvewlistviewbutton2")
    if (hidegidvewlistviewbutton2) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton2.style.display = 'flex'
     
    }
  
    // const hidegidvewlistviewbutton2=document.getElementById("hidegidvewlistviewbutton2")
    // if (hidegidvewlistviewbutton2) {
    //   console.log("enter here .....................")
    //   hidegidvewlistviewbutton2.style.display = 'none'
     
    // }
  
    // const hidegidvewlistviewbutton2=document.getElementById("hidegidvewlistviewbutton2")
    // if (hidegidvewlistviewbutton2) {
    //   console.log("enter here .....................")
    //   hidegidvewlistviewbutton2.style.display = 'flex'
     
    // }
  
    if(event) {
      event.preventDefault();
      event.stopPropagation();
    }
  
    console.log("myFavorite Function is called");
  
    const container = document.getElementById("files-container");
    if(siteIdToUpdate ===  null){
        container.innerHTML="";
    }
  
    
  
    // Fetch the list of active lists
    const FilesItems = await sp.web.lists
      .getByTitle("MasterSiteURL")
      .items.select("Title", "SiteID", "FileMasterList", "Active")
      .filter(`Active eq 'Yes'`)();
  
    console.log("Files items", FilesItems);
    console.log("searchInput",searchText);
    FilesItems.forEach(async (fileItem) => {
      if (fileItem.FileMasterList !== null) {
  
  
        // Skip rendering if we're updating only a specific list
        if (siteIdToUpdate && fileItem.SiteID !== siteIdToUpdate) {
          return;
        }
  
        console.log("SiteId", fileItem.SiteID);
  
        // Fetch files marked as favorite
        const filesData = await sp.web.lists
          .getByTitle(`${fileItem.FileMasterList}`)
          .items.select("FileName", "FileUID", "FileSize", "FileVersion")
          .filter(
            `IsFavourite eq 1 and CurrentUser eq '${currentUserEmailRef.current}'`
          )();
  
        console.log("Files", filesData);
  
        // Remove existing content for this specific list to avoid duplication
        const listElements = document.querySelectorAll(
          `[data-list-id='${fileItem.SiteID}']`
        );
        console.log("ListElemet To update",listElements)
        listElements.forEach((el) => el.remove());
  
        // start
        routeToDiffSideBar="myFavourite";
        let filteredFileData;
        if(searchText !== null){
          filteredFileData=filesData.filter((file: any) => file.FileName.toLowerCase().includes(searchText.value.toLowerCase()))
          // console.log("this is filtered data",filteredFileData)
        }else{
          filteredFileData=filesData;
        }
        // end
  
        // change the array name
        // Render only the updated list's items
        filteredFileData.forEach((file) => {
          const {fileIcon, fileExtension}= getFileIcon(file.FileName);
          const truncateText = (text: string, maxLength: number) => {
            return text.length > maxLength ? text.substring(0, maxLength) + "..." : text;
          };
          const card = createFileCard(file, fileIcon, fileItem.SiteID,fileItem.FileMasterList,fileExtension);
          container.appendChild(card);
        });
      }
    });
  
    return;
  };
  
  // This Function create the File card
  // This Function create the File card
  const createFileCard = (file:any, fileIcon:any, siteId:any,listToUpdate:any,fileExtension:any) => {
     
    const card = document.createElement("div");
    card.className = "card";
    card.dataset.fileId = file.FileUID; // Store file ID in the card element
    card.dataset.listId = siteId; // Store site ID
  
    const truncateText = (text: string, maxLength: number) => {
      return text.length > maxLength ? text.substring(0, maxLength) + "..." : text;
    };
    card.innerHTML = `        
      <img class="filextension" src=${fileIcon} alt="${fileExtension} icon"/>
      <p class="p1st">${truncateText(file.FileName, 10)}</p>
      <p class="p2nd">${file.FileVersion}</p>
      <p class="p3rd">${file.FileSize} MB</p>
      <div id="three-dots" class="three-dots" onclick="toggleMenu2('${file.FileUID}', '${siteId}')">
        <span>...</span>
      </div>
    `;
  
    const menu = document.createElement("div");
    menu.id = `menu-${file.FileUID}`;
    menu.className = "popup-menu";
    menu.innerHTML = `
      <ul>
        <li onclick="confirmDeleteFile('${file.FileUID}', '${siteId}','${listToUpdate}' )">
          <img src=${deleteIcon} alt="Delete"/> Delete
        </li>
        <li onclick="unMarkAsFavorite('${file.FileUID}', '${siteId}','${listToUpdate}')">
          <img src=${FillFavouriteFile} alt="Unmark as Favorite"/> Unmark as Favorite
        </li>
        <li onclick="shareFile('${file.FileUID}', '${siteId}')">
          <img src=${ShareFile} alt="Share"/> Share
        </li>
      </ul>
    `;
  
    card.appendChild(menu);
  
    return card;
  };
  
  
  //Manage UnMark File  
    // @ts-ignore
    window.unMarkAsFavorite = async (fileId: number, siteId: string, listToUpdate: string) => {
      console.log("File Id is ", fileId);
      console.log("siteId is ", siteId,);
      console.log( "List ", listToUpdate);
   
      try {
        const list = sp.web.lists.getByTitle(`${listToUpdate}`);
        console.log("List", list);
        const isFavourite=false;
        const items = await list.items.filter(`FileUID eq '${fileId}' and CurrentUser eq '${currentUserEmailRef.current}'`)();
        console.log("File Data",items)
        if (items.length > 0) {
          if (items[0].IsFavourite && items[0].CurrentUser === currentUserEmailRef.current) {
            const itemId = items[0].Id;
            await list.items.getById(itemId).update({
              IsFavourite: isFavourite
            });
            console.log(`Item with FileUID '${fileId}' updated successfully.`);
            // Re-render only the modified list
            await myFavorite(null, siteId);
          }
       
        } else {
          console.log(`No item found with FileUID '${fileId}'.`);
        }
   
      } catch (error) {
        console.log("This error is from unMarkAsFavorite function ", error);
      }
    };
  
    // function to toggle between Favourite and UnFavourite
  // @ts-ignore
  window.toggleFavourite=async (fileId,siteId)=> {
   
    console.log("SiteId",siteId)
   
    const favouriteToggle = document.getElementById(`favouriteToggle-${fileId}`);  
    const markAsFavouriteIcon = favouriteToggle?.querySelector('.mark-as-favourite') as HTMLElement;
    const unMarkAsFavouriteIcon = favouriteToggle?.querySelector('.unmark-as-favourite') as HTMLElement;
    const textElement = favouriteToggle?.querySelector('.favourite-text') as HTMLElement;
   
    console.log("current Entity",currentEntity);
    let listToUpdate=`DMS${currentEntity}FileMaster`;
   
    async function markAsFavourite(fileId:any, siteId:any){
         
          const siteContext = await sp.site.openWebById(siteId);
          const folderData = await siteContext.web.getFolderByServerRelativePath(currentfolderpath).files
          .expand('ListItemAllFields')();
   
          const isFavourite=true;
          const payload={
            FileName:"",
            FileUID:fileId,
            FileVersion:"Version No.30.0",
            FileSize:"",
            IsFavourite:isFavourite,
            CurrentUser:currentUserEmailRef.current,
            CurrentFolderPath:currentfolderpath,
            DocumentLibraryName:currentDocumentLibrary,
            FolderName:currentFolder
          }
   
          folderData.forEach(async (file)=>{
            if(file.UniqueId === fileId){
              payload.FileName=file.Name;
              payload.FileSize=((file.Length as unknown as number) / (1024 * 1024)).toFixed(2);                
            }
          })
          console.log(payload);
   
          // Get the list by name
          const list = sp.web.lists.getByTitle(listToUpdate);
   
          const data=await sp.web.lists.getByTitle(listToUpdate).items
          .filter(`FileUID eq '${fileId}' and CurrentUser eq '${currentUserEmailRef.current}'`)();
          console.log("Data",data);
   
          // Add the new item to the list
          if(data.length>0){
            const itemId = data[0].Id;
            console.log("items ID",itemId);
            if(!data[0].IsFavourite && currentUserEmailRef.current === data[0].CurrentUser){
           
                const updatedData=await sp.web.lists.getByTitle(listToUpdate).items.getById(itemId).update({
                  IsFavourite:true
                });
                console.log("Updated data",updatedData)
          }
   
          }else{
              const addedItem = await list.items.add(payload);
              console.log("New item added successfully:", addedItem);
          }
         
    }
   
   
    async function UnmarkAsFavourite(fileId:any){
     
     
      try {
       
        const data=await sp.web.lists.getByTitle(listToUpdate).items
        .filter(`FileUID eq '${fileId}' and CurrentUser eq '${currentUserEmailRef.current}'`)();
   
        console.log("Data",data);
        const isFavourite=false;
   
        if (data.length > 0) {
          const itemId = data[0].Id;
          console.log("items ID",itemId);
          if(data[0].IsFavourite && data[0].CurrentUser === currentUserEmailRef.current){
              const updatedData=await sp.web.lists.getByTitle(listToUpdate).items.getById(itemId).update({
                IsFavourite:isFavourite
              });
   
              console.log("Updated data",updatedData);
          }else{
            console.log("Can not find item relataed to current user to unmark");
          }
         
       
        } else {
          console.log("No items found with FileUID:",  fileId);
        }
       
      } catch (error) {
        console.error("Error updating the list item:", error);
      }
    }
   
    try {
         
          if ( markAsFavouriteIcon && unMarkAsFavouriteIcon && textElement) {
         
            // Toggle visibility between the two SVGs and text content
            if (markAsFavouriteIcon.style.display === 'none') {
              markAsFavouriteIcon.style.display = 'inline';
              unMarkAsFavouriteIcon.style.display = 'none';
              textElement.textContent = 'Mark as Favourite';
                   
              // Call function to unmark as favourite.
              UnmarkAsFavourite(fileId);
            } else {
              markAsFavouriteIcon.style.display = 'none';
              unMarkAsFavouriteIcon.style.display = 'inline';
              textElement.textContent = 'Unmark as Favourite';
             
              // Call function to mark as favourite.
              markAsFavourite(fileId, siteId);
            }
          }
    } catch (error) {
             console.log("This Error From toggleFavourite Function",error);
    }
   
  }
   
  // This function give the File Icon
  // const getFileIcon = (fileName:any) => {
         
     
  //   const fileExtension = fileName.split(".").pop().toLowerCase();
  //   let fileIcon;
  //   switch (fileExtension) {
  //     case "doc":
  //     case "docx":
  //       fileIcon = require("../assets/DOC.png");
  //       break;
  //     case "txt":
  //       fileIcon = require("../assets/TXT.png");
  //       break;
  //     case "pdf":
  //       fileIcon = require("../assets/PDF.png");
  //       break;
  //     case "xls":
  //     case "xlsx":
  //       fileIcon = require("../assets/XLS.png");
  //       break;
  //     case "zip":
  //       fileIcon = require("../assets/ZIP.png");
  //       break;
  //     default:
  //       fileIcon = require("../assets/DOC.png"); // Default icon if no match
  //       break;
  //   }
  //   return {fileIcon,fileExtension};
  // };
  
    //My request Files
    const myRequest = async (event:React.MouseEvent<HTMLButtonElement>=null, siteIdToUpdate: string = null,    searchText:any=null ) => {
      const wait2 = document.getElementById('hidegidvewlistviewbutton')
      wait2.classList.remove('hidemydatacards')
      
      setTimeout(() => {
        // alert("set timer")
        setlistorgriddata('');  // Update state to '' after a delay
   
        console.log(listorgriddata, "list")
      }, 100);
      
      const wait = document.getElementById('files-container')
      wait.classList.remove('hidemydatacards')
      setShowMyrequButtons(true)
      setShowMyfavButtons(false)
      setMyreqormyfav('Myrequest')
      // setlistorgriddata('')
      const hidegidvewlistviewbutton=document.getElementById("hidegidvewlistviewbutton")
      if (hidegidvewlistviewbutton) {
        console.log("enter here .....................")
        hidegidvewlistviewbutton.style.display = 'flex'
       
      }
      const hidegidvewlistviewbutton2=document.getElementById("hidegidvewlistviewbutton2")
      if (hidegidvewlistviewbutton2) {
        console.log("enter here .....................")
        hidegidvewlistviewbutton2.style.display = 'none'
       
      }
  
  
      // console.log(listorgriddata , "listorgriddata")
      console.log("searchInput",searchText);
      console.log("siteIdToUpdate",siteIdToUpdate);
  
      if(event){
        event.preventDefault();
        event.stopPropagation();
      }
      
  
  
  
  
      // call this function onClick of the myRequest
      // handleShowContent(event)
      const createFileButton2 = document.getElementById('createFileButton2')
      const createFileButton = document.getElementById('createFileButton')
  
      if(createFileButton2){
      createFileButton2.style.display = 'none'
      }
      if(createFileButton){
      createFileButton.style.display = 'none'
      }
       
  
      
      if(event) {
        event.preventDefault();
        event.stopPropagation();
      }
    
      // console.log("myFavorite Function is called");
    
      const container = document.getElementById("files-container");
      if(siteIdToUpdate ===  null){
          container.innerHTML="";
          // console.log("siteToUpdate")
      }
     
      // console.log("beforeFetchItems");
      // Fetch the list of active lists
      const FilesItems = await sp.web.lists
        .getByTitle("MasterSiteURL")
        .items.select("Title", "SiteID", "FileMasterList", "Active")
        .filter(`Active eq 'Yes'`)();
    
       console.log("Active Sites List Names", FilesItems);
    
      FilesItems.forEach(async (fileItem) => {
        if (fileItem.FileMasterList !== null) {
    
          // console.log("FilesItesms");
          // Skip rendering if we're updating only a specific list
          if (siteIdToUpdate && fileItem.SiteID !== siteIdToUpdate) {
            return;
          }
    
          // console.log("SiteId", fileItem.SiteID);
    
          const filesData = await sp.web.lists
          .getByTitle(`${fileItem.FileMasterList}`)
          .items.select("ID" , "FileName", "FileUID", "FileSize", "FileVersion" ,"Status" , "SiteID","CurrentFolderPath","DocumentLibraryName","SiteName")
          .filter(
            `CurrentUser eq '${currentUserEmailRef.current}'`
          )();
          // console.log("My reaquest Called");
  
          // console.log("enter in the myRequest------")
          console.log("FilesData my req",filesData)
        // route to different-2 sideBar
  
        // start
        routeToDiffSideBar="myRequest";
        let filteredFileData=[];
        if(searchText !== null){
              filteredFileData=filesData.filter((file: any) => file.FileName.toLowerCase().includes(searchText.value.toLowerCase()))
              // console.log("this is filtered data",filteredFileData)
        }else{
          filteredFileData=filesData;
        }
        // end 
  
        // change the array
        filteredFileData.forEach((file) => {
        //  console.log(file.ID , "file.odata.id ")
         // Function to truncate text
    const truncateText = (text: string, maxLength: number) => {
      return text.length > maxLength ? text.substring(0, maxLength) + "..." : text;
    };
        console.log(file, "here is our file")
        const card = document.createElement("div");
        
        // console.log("searchArray",searchArray);
        let fileIcon;
        const fileExtension = file.FileName.split(".").pop().toLowerCase(); // Get the file extension
        switch (fileExtension) {
          case "doc":
          case "docx":
            fileIcon = Docicon;
            break;
          case "txt":
            fileIcon = Txticon;
            break;
          case "pdf":
            fileIcon = Pdficon;
            break;
          case "xls":
          case "xlsx":
            fileIcon = Xlsicon;
            break;
          case "zip":
            fileIcon = Zipicon;
            break;
          default:
            fileIcon = Docicon; // Default icon if no match
            break;
        }
    
        card.className = "card";
        card.innerHTML = `  
            <div class="IMGContainer">
             
          <img class="filextension" src=${fileIcon} alt="${fileExtension} icon"/>
  
           </div>     
                   <div class="CardTextContainer">
          <p class="p1st">${truncateText(file.FileName, 10)}</p>
          <p class="p2nd"></p>
          <p class="p3rd">${((file.FileSize as unknown as number) / (1024 * 1024)).toFixed(2)} MB</p>
          <p class="filestatus"> ${file.Status}  </p>
          <div class="three-dots" onclick="toggleMenu2('${file.FileUID}','${file.SiteID}','${file.ID}' , '${file.FileMasterList}', '${file.FilePreviewURL}')  ">
              <span>...</span>
          </div>
           </div>
        `;
    
        const menu = document.createElement("div");
        // console.log(menu , "menu is here")
        menu.id = `menu-${file.FileUID}`;
        menu.className = "popup-menu";
        const showaudit = <FontAwesomeIcon style={{color: "black"}} icon={faListSquares}/>
        menu.innerHTML = `
         <ul>
        <li onclick="confirmDeleteFile('${file.FileUID}','${file.SiteID}','${fileItem.FileMasterList}')">
          <img src=${deleteIcon} alt="Delete"/> Delete
        </li>
   <li onclick="auditHistory('${file.FileUID}', '${file.SiteID}','${file.DocumentLibraryName}','${file.SiteName}')">
              <img src=${editIcon} alt="Edit"/>
                          Audit History
        </li>
        <li onclick="shareFile('${file.FileUID}', '${file.SiteID}','${file.CurrentFolderPath}','${file.FileName}')"
 >
          <img src=${ShareFile} alt="Share"/> Share
        </li>
        <li onclick="PreviewFile('${file.FileUID}','${file.SiteID}','${file.ID}' , '${file.FileMasterList}', '${file.FilePreviewURL}')">
          <img src=${ShareFile} alt="Preview File"/> Preview File
        </li>
      </ul>
        `;
    
        card.appendChild(menu);
    
        // Change the background color and text color based on FileStatus
        const fileStatusElement = card.querySelector(".filestatus") as HTMLElement;
        switch (file.Status) {
          case "Approved":
            fileStatusElement.style.backgroundColor = "#b5e7d3";
            fileStatusElement.style.color = "#008751";
            break;
          case "Rejected":
            fileStatusElement.style.backgroundColor = "rgba(241, 85, 108, 0.1)";
            fileStatusElement.style.color = "#f1556c";
            break;
          case "Rework":
            fileStatusElement.style.backgroundColor = "#ffecc4";
            fileStatusElement.style.color = "rgba(247, 184, 75)";
            break;
            case "Pending":
              fileStatusElement.style.backgroundColor = "rgb(91 156 187 / 25%)";
              fileStatusElement.style.color = "#000b56";
              break;
              default:
                fileStatusElement.style.backgroundColor = "gray";
                fileStatusElement.style.color = "white";
                break;
        }
    
        container.appendChild(card);
          });
        }
      });
    
    };
  
    
    const [activeComponent, setActiveComponent] = useState<string | 'MyRequest'>('');
    const [listorgriddata, setlistorgriddata] = useState<string>('');
    const handleButtonClickShow = (componentName:any) => {
      setActiveComponent(componentName); // Set the active component based on the button clicked
    };
    const handleReturnToMain = () => {
      setActiveComponent(''); // Reset to show the main component
    };
    
  
    const MyrequestshowListView = (componentName:any)=>{
      const wait = document.getElementById('files-container')
      wait.classList.add('hidemydatacards')
      setlistorgriddata('showListView');
    }
  
    // side text content based on click 
    // Handle button click and show the text of the clicked button
    const [selectedText,setSelectedText]=useState<string | null>(null);
    const [dynamicContent, setDynamicContent] = useState<string | null>(null);
     // Function to update the breadcrumb navigation
     const updateBreadcrumb = (path:any) => {
        console.log(path, "path")
      // For toggle the breadcrumb and selectedTextForSideBar
      const selectedTextDiv=document.getElementById('selectedText');
      const breadcrumbElement=document.getElementById("breadcrumb");
  
      if(breadcrumbElement){
        // breadcrumbElement.style.position = "absolute"
        breadcrumbElement.style.width = ""      // Change width  by Amjad
        breadcrumbElement.style.top = "115px"
        breadcrumbElement.style.display='block';
      }
   if(selectedTextDiv){
    selectedTextDiv.style.display='none';
   }
     
   
   
      if (breadcrumbElement) {
        breadcrumbElement.textContent = path;
      }
    };
   // Function to handle navigation and update breadcrumb
   const handleNavigation = (title:string,Devision:string  , Department:string ,  docLibName:string=null, folderName:string=null) => {
    let path = title;
    if(Devision) {
      path += ` > ${Devision}`;
    }
    if(Department) {
      path += ` > ${Department}`;
    }
    if (docLibName) {
      path += ` > ${docLibName}`;
    }
  
    if (folderName) {
      path += ` > ${folderName}`;
    }
  
    updateBreadcrumb(path);
  };
    const handleShowContent = (event: React.MouseEvent<HTMLButtonElement>) => {
      // console.log("enter here")
      event.preventDefault();
     
      //toggle the breadcrumb and selectedText For SideBar
      const selectedTextDiv=document.getElementById('selectedText');
      const breadcrumbElement=document.getElementById("breadcrumb");
      breadcrumbElement.style.display='none';
      selectedTextDiv.style.display='block';
   
   
      // Find the button that was clicked
      const button = event.currentTarget;
   
     
      const spanElement = button.querySelector('.sidebarText');
      const text = spanElement?.textContent;
   
      if (text) {
        setSelectedText(text);
   
        // Update dynamic content based on the button clicked
        switch (text) {
          case 'My Request':
            setDynamicContent('Mentioned below are the documents submitted by logged in user.');
            break;
          case 'My Favourite':
            setDynamicContent('All the files and folder which is marked as Favourite.');
            break;
          case 'My Folder':
            setDynamicContent('All the folder Created by me.');
            break;
          case 'Share with Other':
            setDynamicContent('My files shared with other users.');
            break;
          case 'Share with me':
            setDynamicContent('File upload by other team members and shared with me.');
            break;
          default:
            setDynamicContent(null);
        }
      }
  };
  
  const search = document.getElementById('searchinput')
   if(search){search.addEventListener("keydown", function(event) {
    if (event.key === "Enter") {
      event.preventDefault();  // Prevent default behavior
    }
  })};

  // // Manage Folder WorkFlow Action
  window.manageWorkflow=async(DocumentLibraryName:string,SiteTilte:string, SiteID:string)=>{
    setShowWorkflow(true)
    console.log("enter in workflow")
    // const workflowdiv= document.getElementById('workflowdiv')
    // if(workflowdiv){
    //   workflowdiv.classList.remove('workflowdivhide')
    //   alert(workflowdiv.classList)
    //   workflowdiv.classList.add('workflowdiv')
    // }
    // setshowworkflowdiv("true")

    // setshowworkflowdiv(true)
    // console.log("Inside manageWorkFlow");
  
    propsForManageWorkFlow.SiteTitle=SiteTilte;
    propsForManageWorkFlow.DocumentLibraryName=DocumentLibraryName;
    propsForManageWorkFlow.SiteID=SiteID;
    // handleButtonClickShow("manageWorkFlow");
 
  }

   //Manage Folder Permission Action
   window.managePermission=(documentLibraryName:string,SiteTilte:string,SiteID:string)=>{
    setShowfolderpermission(true)
    // console.log(message);
    console.log("documentLibraryName",documentLibraryName)
    console.log("SiteName",SiteTilte);
    console.log("siteId",SiteID);
 
    managePermissionProps.DocumentLibraryName=documentLibraryName;
    managePermissionProps.SiteTitle=SiteTilte;
    managePermissionProps.SiteID=SiteID;
    // handleButtonClickShow("managePermission");
 
  }

  // Edit File
  window.editFile = async (siteName: string, documentLibraryName:string ) => {
    console.log("Inside the editFile");

    // Fetch the existing columns and types from the list
    const existingColumns = await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("ColumnName", "ColumnType","ID").filter(`SiteName eq '${siteName}' and DocumentLibraryName eq '${documentLibraryName}' and IsDocumentLibrary eq 0`)();

    console.log("existingColumns",existingColumns);

    // Create the popup container dynamically
    const popupContainer = document.createElement("div");
    popupContainer.className = "edit-popup";

    // Append to body
    document.body.appendChild(popupContainer);

    // Create close button
    const closeButton = document.createElement("span");
    closeButton.innerHTML = 'x';
    closeButton.className = 'close-button';
    closeButton.style.cursor = 'pointer';
    closeButton.style.fontSize = '24px';
    closeButton.style.position = 'absolute';
    closeButton.style.top = '10px';
    closeButton.style.right = '15px';

    // Append close button to popup
    popupContainer.appendChild(closeButton);


    // Add form elements for each existing column
    const formContent = document.createElement("div");
    formContent.innerHTML = existingColumns.map((col) => `
      <div class="form-group">
        <div class="col-md-5">
          <label>Field Name</label>
          <input type="text" class="form-control" value="${col.ColumnName}" disabled />
        </div>
        <div class="col-md-5">
          <label>Field Type</label>
          <input type="text" class="form-control" value="${col.ColumnType}" disabled />
        </div>
      </div>
    `).join('');


    // Add button for adding new fields
    const addFieldButton = document.createElement("a");
    addFieldButton.innerHTML = `
      <img class="bi bi-plus" src="${require("../assets/plus.png")}" alt="add" style="width: 50px; height: 50px;" />
    `;

    // Append the content to the popup
    popupContainer.appendChild(formContent);
    popupContainer.appendChild(addFieldButton);

    // Add event listener for "+" button to add new editable fields
    addFieldButton.addEventListener('click', () => {
      const newFieldHTML = `
        <div class="form-group">
          <div class="col-md-5">
            <label>Field Name</label>
            <input type="text" class="form-control" placeholder="Enter new field name" />
          </div>
          <div class="col-md-5">
            <label>Field Type</label>
            <select class="form-control">
              <option value="" disabled selected>Select Type</option>
              <option value="Single Line of Text">Single Line of Text</option>
              <option value="Mulitple Line of Text">Multiple Line of Text</option>
              <option value="Yes or No">Yes or No</option>
              <option value="Date & Time">Date & Time</option>
              <option value="Number">Number</option>
            </select>
          </div>
          <div class="col-md-2">
            <img class="delete-column"  src="${require("../assets/delete.png")}" alt="add" style="width: 40px; margin-top:25px; cursor:pointer;" />
          </div>
        </div>
      `;
      formContent.insertAdjacentHTML('beforeend', newFieldHTML);
    });

    // Add save button
    const saveButton = document.createElement("button");
    saveButton.innerText = "Save";
    popupContainer.appendChild(saveButton);

    saveButton.addEventListener('click', () => {
       // Collect data from newly added fields
      const newFields = Array.from(formContent.querySelectorAll('.form-group.row')).map((group) => {
        const fieldNameInput = (group.querySelector('input[type="text"]') as HTMLInputElement);
        const fieldTypeSelect = (group.querySelector('select') as HTMLSelectElement);
       
        return {
            columnName: fieldNameInput.value,
            columnType: fieldTypeSelect.value
        };
    });

    console.log("New Fields:", newFields);

    try {
      const payloadForPreviewFormMaster={
        SiteName:siteName,
        DocumentLibraryName:documentLibraryName,
        IsRequired:true,
        AddorRemoveThisColumn:"Add To Library",
        IsModified:true
      }

    // existingColumns.forEach(async(column)=>{
    //       try {
    //           // Deleting the item with the provided ID from the DMSPreviewFormMaster list
    //           await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.getById(column.ID).delete();
    //           console.log(`Item with ID ${column.ID} deleted successfully from list DMSPreviewFormMaster`);
    //       } catch (error) {
    //           console.log(`Error deleting item with ID ${column.ID} from list DMSPreviewFormMaster`, error);
    //       }
    //   });

      newFields.forEach(async(column)=>{
        (payloadForPreviewFormMaster as any).ColumnName=column.columnName.replace(/\s+/g,'');;
        (payloadForPreviewFormMaster as any).ColumnType=column.columnType;
       
        const addedItem = await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.add(payloadForPreviewFormMaster);
        console.log("Item added successfully in the DMSPreviewFormField", addedItem);

      })
    } catch (error) {
      console.log("Error in adding columns in the DMSPreviewFormMaster inside the editFile onclick of the save button",error)
    }

     
    });

// Add event listener to close button
closeButton.addEventListener('click', () => {
      popupContainer.style.display = 'none';
 });

//add event listener to  removing the field
 // Event delegation
  formContent.addEventListener('click', (event) => {
    const target = event.target as HTMLElement;
    if (target.classList.contains("delete-column")) {
        const fieldGroup = target.closest('.form-group.row');
        if (fieldGroup) {
            fieldGroup.remove();
        }
    }
});
};

//  Share File
window.shareFile=async(fileID:String,siteId:String,currentFolderPathForFile:string,fileName:string)=>{
  console.log("Share File called");
  console.log("file Id",fileID);
  console.log("site Id",siteId);
  console.log("currentFolderPath",currentFolderPathForFile);

  // exreact the Entity from folder path
  const parts = currentFolderPathForFile.split("/");
  console.log(parts , "parts")
  const entity = parts[3];
  console.log(entity);

  const fetchUser=async(entity:any)=>{
    // const [
    //   users,
    //   users1,
    //   users2,
    //   users3,
    //   users4,
    // ] = await Promise.all([
    //   sp.web.siteGroups.getByName(`${entity}_Read`).users(),
    //   sp.web.siteGroups.getByName(`${entity}_Initiator`).users(),
    //   sp.web.siteGroups.getByName(`${entity}_Contribute`).users(),
    //   sp.web.siteGroups.getByName(`${entity}_Admin`).users(),
    //   sp.web.siteGroups.getByName(`${entity}_View`).users(),
    // ]);
    // console.log(users, "users ", users1,users2,users3,users4);
    // const combineArray = [
    //   ...(users || []),
    //   ...(users1 || []),
    //   ...(users2 || []),
    //   ...(users3 || []),
    //   ...(users4 || []),
    // ];

    // const siteContext = await sp.site.openWebById(OthProps.siteID);
    const user0 = await sp.web.siteUsers();
    const combineUsersArray=user0.map((user)=>(
          {
            id:String(user.Id),
            value: user.Title,
            email: user.Email,
          }
        ))
    console.log("Sub site users",combineUsersArray);
     
    // const resultArray=combineUsersArray.map((user) => (
    //   {
    //     id:String(user.Id),
    //     value: user.Title,
    //     email: user.Email
    //   }
    // ))
    // console.log("combineArray", combineArray);
    // console.log("resultArray",resultArray)

    return combineUsersArray;
  }

  const users=await fetchUser(entity);
  console.log("UserArray",users);
 

// Check if a popup already exists, if so, remove it before creating a new one
const existingPopup = document.getElementById('share-popup');
if (existingPopup) {
existingPopup.remove();
}

// Dummy data
// const users = [
//   { value: 'Test1', id: '14',email:"User1@officeindia.onmicrosoft.com" },
//   { value: 'Test2', id: '31',email:"User2@officeindia.onmicrosoft.com" },
//   { value: 'Test3', id: '137',email:"User3@officeindia.onmicrosoft.com"},
//   { value: 'Test4', id: '33',email:"User4@officeindia.onmicrosoft.com" },
//   { value: 'Test5', id: '32',email:"User5@officeindia.onmicrosoft.com" },
//   { value: 'Test6', id: '34',email:"User6@officeindia.onmicrosoft.com" },
//   { value: 'Test User1', id: '39',email:"User7@officeindia.onmicrosoft.com" },
//   ];


// Declare selectedUsers with an explicit type, assuming user IDs are of type string for selecting the user for share
let selectedUsers: { id: string; value: string; email:string }[] = [];
// Create the pop-up element
const popup = document.createElement("div");
popup.id = 'share-popup';
popup.className = "share-popup";

// Add HTML structure for the pop-up with a dropdown and a close "X" button
popup.innerHTML = `
<div class="share-popup-content">
<div class="share-popup-header">
  <h4>Share</h4>
  <span class="share-close-popup" onClick="hideSharePopUp()">x</span>
</div>
<div class="share-popup-body">
  <div id="share-reactSelect">
      <input type="text" id="userInput" placeholder="Add a Name, Group, or Email" style="
      width: 100%;
      padding: 10px;
      font-size: 14px;
      border-radius: 4px;
      border: 1px solid #ccc;
    "/>
    <div id="userDropdown" class="user-dropdown" style="
      display: none;
      position: absolute;
      width: 29.8%;
      max-height: 150px;
      overflow-y: auto;
      background-color: white;
      border: 1px solid #ccc;
      border-radius: 4px;
      z-index: 1000;
    ">
    </div>
  </div>
  <textarea id="share-message" placeholder="Write a message..." >
  </textarea>
</div>
<div class="share-popup-footer">
  <button id="share-shareFileButton">Share</button>
</div>
</div>
`;

// Append the  popup to the body
document.body.appendChild(popup);

// Get references to the input box and dropdown
const userInput = document.getElementById('userInput') as HTMLInputElement;
const userDropdown = document.getElementById('userDropdown');

// Function to render dropdown options based on user input
function renderDropdown(users: { id: string, value: string,email:string }[]) {
// Clear previous options
userDropdown.innerHTML = '';
users.forEach(user => {
const option = document.createElement('div');
option.className = 'dropdown-item';
option.style.padding = '8px';
option.style.cursor = 'pointer';
option.textContent = user.value;
option.onclick = () => selectUser(user);
userDropdown.appendChild(option);
});
}

// Function to show the dropdown when the input is clicked
userInput.addEventListener('focus', () => {
userDropdown.style.display = 'block';

// Display all users initially
renderDropdown(users);
});

// Filter dropdown based on input value
userInput.addEventListener('input', () => {
const searchValue = userInput.value.toLowerCase();
const filteredUsers= users.filter(user => user.value.toLowerCase().includes(searchValue));
renderDropdown(filteredUsers);
});

// Function to select a user and display it inside the input
function selectUser(user: { id: string, value: string,email:string }) {
console.log("selected user",selectedUsers)
if (!selectedUsers.some(selectedUser => selectedUser.id === user.id)) {

selectedUsers.push(user);

// Create a span for the selected user with a close button
const selectedUserDiv = document.createElement('span');
selectedUserDiv.className = 'selected-user';
selectedUserDiv.style.display = 'inline-block';
selectedUserDiv.style.padding = '2px 6px';
selectedUserDiv.style.backgroundColor = '#e0e0e0';
selectedUserDiv.style.borderRadius = '12px';
selectedUserDiv.style.marginRight = '5px';
selectedUserDiv.style.position = 'relative';

selectedUserDiv.textContent = user.value;

// Create close button for deselecting the user
const closeButton = document.createElement('span');
closeButton.textContent = 'x';
closeButton.style.cursor = 'pointer';
closeButton.style.marginLeft = '5px';
closeButton.onclick = () => deselectUser(user.id, selectedUserDiv);
selectedUserDiv.appendChild(closeButton);

// Append the selected user to the input field
userInput.parentNode!.insertBefore(selectedUserDiv, userInput);
userInput.value = '';
}
userDropdown.style.display = 'none';
}

// Function to deselect a user
function deselectUser(userId: string, selectedUserDiv: HTMLElement) {
// selectedUsers = selectedUsers.filter(id => id !== userId);
selectedUsers = selectedUsers.filter(selectedUser => selectedUser.id !== userId);
console.log("selected user",selectedUsers);
selectedUserDiv.remove();
}

// Hide the dropdown if clicked outside
document.addEventListener('click', (event) => {
if (!userInput.contains(event.target as Node) && !userDropdown.contains(event.target as Node)) {
userDropdown.style.display = 'none';
}
});

// Adding event listener to the "Share" button
document.getElementById('share-shareFileButton').addEventListener('click', async function() {
    console.log("selectedUserArray",selectedUsers);
    console.log("Entity",entity);
    console.log("FileId",fileID)
    console.log("SiteId",siteId);

    const listToUpdateWithShareData=`DMS${entity}FileMaster`;
    console.log("listToUpdateWithShareData",listToUpdateWithShareData);

    // Fetch the item from the list using its ID
    const item = await sp.web.lists.getByTitle(listToUpdateWithShareData).items.select("FileName","ShareWithOthers","ShareWithMe","FileUID","ID").filter(`FileUID eq '${fileID}' and CurrentUser eq '${currentUserEmailRef.current}'`)();

    console.log("item",item);

    // let dataArray;
    let dataArray: Array<{ FirstName: string; LastName?: string; SharedWith: string; SharedAt: string; TimeStamp: number; Permission: string,userId:string }> = [];
         
    selectedUsers.forEach(async(user)=>{
   
    const nameParts = user.value.trim().split(" ");
    const firstName = nameParts[0];
    let lastName = nameParts.length > 1 ? nameParts[nameParts.length - 1] : "";
    console.log("firstName",firstName)
    console.log("lastName",lastName);
    if(lastName === ""){
      lastName="";
    }

    const isoDate = new Date().toISOString().slice(0, 19) + 'Z';
    const timestamp = Date.now();
      let userObj={
        FirstName:firstName,
        LastName:lastName,
        SharedWith:user.email,
        SharedAt:isoDate,
        TimeStamp:timestamp,
        Permission:"Read",
        userId:user.id
      }
      dataArray.push(userObj);
      // console.log("userObj",userObj);
    })

    console.log("dataArray",dataArray);

   

    if(item[0].ShareWithMe === null && item[0].ShareWithOthers === null){

          const dataInTheFormoOfString=JSON.stringify(dataArray);
           // Now update specific columns of the item
            const updatedItem = await sp.web.lists.getByTitle(listToUpdateWithShareData).items.getById(item[0].ID).update({
              ShareWithOthers:dataInTheFormoOfString,
              ShareWithMe:dataInTheFormoOfString
            });

            console.log("Data updated when ShareWithMe and ShareWithOthers are null",updatedItem);
    }else{
       const shareWithOthers =JSON.parse(item[0].ShareWithOthers);
       const shareWithMe=JSON.parse(item[0].ShareWithMe);

       dataArray.forEach((user)=>{
            // apply condition for sharing same file with same user multiple time using id of the user
            const alReadySharedUserIndex=shareWithOthers.findIndex((item:any)=>{
                  return item.userId === user.userId
            })
            console.log("alReadySharedUser in shareWithOthers",alReadySharedUserIndex);
            const alReadySharedUserIndex1=shareWithMe.findIndex((item:any)=>{
                return item.userId === user.userId
            })
            console.log("alReadySharedUser in shareWithMe",alReadySharedUserIndex1);

            if(alReadySharedUserIndex !== -1){
                  shareWithOthers.splice(alReadySharedUserIndex, 1);
                  shareWithOthers.push(user);
                  console.log("shareWithOthers",shareWithOthers);
            }else{
              shareWithOthers.push(user);
            }

            if(alReadySharedUserIndex1 !== -1){
              shareWithMe.splice(alReadySharedUserIndex1, 1);
              shareWithMe.push(user);
              console.log("shareWithMe",shareWithMe);
            }else{
              shareWithMe.push(user);
            }
       })

       console.log("shareWithOthers",shareWithOthers);
       console.log("shareWithMe",shareWithMe);

       const dataInTheFormoOfStringForShareWithMe=JSON.stringify(shareWithMe);
       const dataInTheFormoOfStringForShareWithOthers=JSON.stringify(shareWithOthers);
       // Now update specific columns of the item
       const updatedItem = await sp.web.lists.getByTitle(listToUpdateWithShareData).items.getById(item[0].ID).update({
        ShareWithOthers:dataInTheFormoOfStringForShareWithOthers,
        ShareWithMe:dataInTheFormoOfStringForShareWithMe
      });

      console.log("Data updated when ShareWithMe and ShareWithOthers",updatedItem);
    }

});


}

// hide the share popup
// @ts-ignore
window.hideSharePopUp=()=>{
const popup=document.querySelector('.share-popup');

if(popup){
  popup.remove();
}
}


 // show the audit history popup
  // @ts-ignore
  window.auditHistory=async(fileId:string, siteId:string,DocumentLibraryName:string,SiteName:String)=>{
    console.log("Audit History called",fileId,siteId);
    console.log("Audit History called",SiteName);
    console.log("Audit History called",DocumentLibraryName);
 
 
    const {web}=await sp.site.openWebById(siteId)
 
     // Get the list item  corresponding to the file
     const fileItem:any = await web.getFileById(fileId).expand("ListItemAllFields")();
    //  console.log("fileItem status",fileItem.Status);
     console.log("fileItem",fileItem.ListItemAllFields.Status);
   
    // fetched the columns details corresponding to the file
    const fileColumns =await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("ColumnName","SiteName","DocumentLibraryName").filter(`SiteName eq '${SiteName}' and DocumentLibraryName eq '${DocumentLibraryName}' and IsDocumentLibrary ne 1`)();
    console.log("fileColumns",fileColumns);
 
    // Create an array of objects to store the columnName with there corresponding value
    const resultArrayThatContainstheColumnDetails = fileColumns.map((column) => {
    const columnName = column.ColumnName;
    const columnValue = fileItem.ListItemAllFields[columnName];
 
      return {
        label: columnName,
        value: columnValue !== undefined ? columnValue : null // Handle missing fields
      };
    });
 
    const objectForStatus={
      label:"Status",
      value:fileItem.ListItemAllFields.Status || ""
    }
 
    resultArrayThatContainstheColumnDetails.push(objectForStatus);
    console.log("result",resultArrayThatContainstheColumnDetails);
 
    // get the details of approver
    const itemsFromTaskList = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select(
      "Log","CurrentUser","Remark"    
           ,"LogHistory","ID"                  
           ,"FileUID/FileUID"          
           ,"FileUID/SiteName"              
           ,"FileUID/DocumentLibraryName"
           ,"FileUID/FileName"                
           ,"FileUID/Status"    
           ,"FileUID/RequestedBy"  
           ,"FileUID/Created"  
           ,"FileUID/ApproveAction"
           ,"MasterApproval/ApprovalType"
           ,"MasterApproval/Level"  
           ,"MasterApproval/DocumentLibraryName"
           ,"Modified"
        )
        .expand("FileUID", "MasterApproval")
        .filter(`FileUID/FileUID eq '${fileId}'`)
        .orderBy("Modified", false)();
 
        console.log("itemsFromTaskList",itemsFromTaskList);
 
 
    // Mapping to the desired format
    const approverDetailsArray = itemsFromTaskList.map(task => ({
      level: `Level ${task.MasterApproval.Level}`,
      approver: task.CurrentUser,
      actionDateTime:task.Modified,
      status: task.Log || "",
      remark: task.Remark || ""
    }));
 
    console.log("approverDetailsArray",approverDetailsArray);
 
  // Generate the dynamic HTML for the detail rows
  let detailRowsHTML = "";
  resultArrayThatContainstheColumnDetails.forEach((item, index) => {
      // Start a new row every 3rd item (when index is 0, 3, 6, ...)
      if (index % 3 === 0) {
        detailRowsHTML += '<div class="detail-row">';
      }
 
      // Add each detail column
      detailRowsHTML += `
        <div class="detail-column">
          <div class="detail-label">${item.label}</div>
          <div class="detail-value">${item.value}</div>
        </div>
      `;
 
      // Close the row after 3 items (when index is 2, 5, 8, ...)
      if ((index + 1) % 3 === 0) {
        detailRowsHTML += '</div>';
      }
  });
 
  // If there are leftover columns (less than 3 in the last row), close the row
  if (resultArrayThatContainstheColumnDetails.length % 3 !== 0) {
    detailRowsHTML += '</div>';
  }
 
   // Generate the dynamic HTML for the approver details
   let approverRowsHTML = "";
   approverDetailsArray.forEach((approver) => {
     approverRowsHTML += `
       <div class="detail-row-value-approver">
         <div class="detail-value-approver">${approver.level}</div>
         <div class="detail-value-approver">${approver.approver}</div>
         <div class="detail-value-approver">${approver.actionDateTime}</div>
         <div class="detail-value-approver">${approver.status}</div>
         <div class="detail-value-approver">${approver.remark}</div>
       </div>
     `;
   });
 
     // Create the popup
    const popup = document.createElement("div");
    popup.className = "audit-history-popup";
    popup.innerHTML = `
    <div class="popup-content-auditHistory">
      <div class="popup-header">
        <h5>Audit History</h5>
        <span class="close-btn" onclick="hideAuditHistoryPopup()">&times;</span>
      </div>
      <div class="popup-details">
        ${detailRowsHTML}
        <div class="detail-row-approver">
          <div class="detail-label-approver">Approval Level</div>
          <div class="detail-label-approver">Approver</div>
          <div class="detail-label-approver">Action DateTime</div>
          <div class="detail-label-approver">Status</div>
          <div class="detail-label-approver">Remark</div>
        </div>
        ${approverRowsHTML}
      </div>
    </div>
    `;
 
  // Append to body
  document.body.appendChild(popup);
  }
 
  // function to hide audit history pop
  // @ts-ignore
  window.hideAuditHistoryPopup=()=> {
    const popup = document.querySelector(".audit-history-popup");
    if (popup) {
      popup.remove();
    }
  }
  // start
    // Ref for MyFolder
    const myFolderButtonRef = useRef(null);
    const [triggerClick, setTriggerClick] = useState(false);
 
    const handleReturnToMainFromManageWorkFlow=()=>{
      setShowWorkflow(false);
      setShowfolderpermission(false)
      setActiveComponent('');
      setTriggerClick(true);
    }
 
    useEffect(() => {
      if (triggerClick && myFolderButtonRef.current) {
        myFolderButtonRef.current.click();  
        setTriggerClick(false);
      }
    }, [triggerClick]);
   useEffect(()=>{
    const workflowdiv= document.getElementById('workflowdiv')
    if(workflowdiv){
      workflowdiv.classList.add('workflowdivhide')
    }
   })
    // end
    return (
      <div id="wrapper" ref={elementRef}>
      <div
        className="app-menu"
        id="myHeader">
        <VerticalSideBar _context={sp} />
      </div>
      <div className="content-page">
        <HorizontalNavbar/>
        <div className="content" style={{marginLeft: `${!useHide ? '240px' : '80px'}`,marginTop:'0.8rem'}}>
         
        <div className="container-fluid  paddb">
                  {activeComponent === "" ? (
                    <div className=" dmsmaincontainer">
                      {showWorkflow && (
        <div id="workflowdiv">
          <ManageWorkFlow
            OthProps={propsForManageWorkFlow}
            onReturnToMain={handleReturnToMainFromManageWorkFlow}
          />
        </div>
      )}
      {showfolderpermission  && (
        <div id="showfolderpermission">
                   <ManageFolderPermission
                        OthProps={managePermissionProps}
                        onReturnToMain={handleReturnToMainFromManageWorkFlow}
                        />
        </div>
                       
                      )}
                  {/* Start Code Update by Amjad */}
                      <div className="row">
                               <div className="col-lg-6">
                                  <h4 className="page-title fw-bold mb-1 font-20">Dossier</h4>
                                  <ol className="breadcrumb m-0">
                      {" "}
                      <li className="breadcrumb-item">Home</li>
                     
                      <li className="breadcrumb-item">&gt;</li>{" "} 
                      <li className="breadcrumb-item active">Settings</li>
                    </ol>
                              </div>
  
                              
                              <div className="col-lg-6">
                          <div id="hidegidvewlistviewbutton" className="view-buttons mt-2">
                                  <button className="btn btngridview grid-view active"    
                                  onClick={(event: any = null, siteIdToUpdate: string = null)=>myRequest(event) }>
                                    <a className="listviewfonticon">          
                                      <FontAwesomeIcon style={{color: "black"}} icon={faTableCells}/> </a>Grid View
                                  </button>
                                  <button className="btn btnlistview list-view" onClick={(event:any)=>MyrequestshowListView('ListViewComponent')}>
                                    <a className="listviewfonticon">
                                    <FontAwesomeIcon style={{color: "black"}} icon={faListSquares}/>
                                    </a>
                                List View
                                  </button>
                                 
                            </div>
                            {showMyfavButtons && ( <div id="hidegidvewlistviewbutton2"  className="view-buttons mt-2">
                                    <button className="btn btngridview grid-view active"    
                                    onClick={(e)=>myFavorite(e)}>
                                      <a className="listviewfonticon">          
                                        <FontAwesomeIcon style={{color: "black"}} icon={faTableCells}/> </a>Grid View
                                    </button>
                                    <button className="btn btnlistview list-view" onClick={(event:any)=>MyrequestshowListView('ListViewComponent')}>
                                      <a className="listviewfonticon">
                                      <FontAwesomeIcon style={{color: "black"}} icon={faListSquares}/>
                                      </a>
                                      List View
                                    </button>
                            </div>) 
                            }
                            </div>

                            </div>
                             {/* End Code Update by Amjad */}
                      <div className="mainsidebardms">
                      
                      
                          
                        
                        <div className="sidebardms">
                     
                          <button
                          id= "Myrequestbutton"
                            className={`sidebardmsButton ${
                              activeButton === "MyRequest" ? "active" : ""
                            }`}
                            // onClick={() => handleClick('MyRequest')}
                            onClick={
                              (event)=>{
                                
                                myRequest(event);
                                handleShowContent(event)
                            }
                          }
                          >
                            <span className="sidebarIcon">
                              {/* <FontAwesomeIcon icon={faList} /> */}
                              <img className="sidebariconssmall" src={listicon}></img>
                            </span>
                            <span className="sidebarText">My Request</span>
                          </button>
  
                          <button
                            className={`sidebardmsButton ${
                              activeButton === "MyFavourite" ? "active" : ""
                            }`}
                            onClick={(event) => {  myFavorite(event);
                              handleShowContent(event);}}
                          >
                            <span className="sidebarIcon">
                            <img className="sidebariconssmall" src={starticon}></img>
                              {/* <FontAwesomeIcon icon={faStarRegular} /> */}
                            </span>
                            <span className="sidebarText">My Favourite</span>
                          </button>
  
                          <button
                            className={`sidebardmsButton ${
                              activeButton === "MyFolder" ? "active" : ""
                            }`}
                            onClick={(event)=>{
                              mycreatedfolders(event);
                              handleShowContent(event)
                            }}
                          >
                            <span className="sidebarIcon">
                            <img className="sidebariconssmall" src={foldericon}></img>
                              {/* <FontAwesomeIcon icon={faFolderRegular} /> */}
                            </span>
                            <span className="sidebarText">My Folder</span>
                          </button>
  
                          <button
                            className={`sidebardmsButton ${
                              activeButton === "ShareWithOther" ? "active" : ""
                            }`}
                            onClick={(e)=>ShareWithOther(e)}
                          >
                            <span className="sidebarIcon">
                              {/* <FontAwesomeIcon icon={faShareAlt} /> */}
                              <img className="sidebariconssmall" src={sharewithothericon}></img>
                            </span>
                            <span className="sidebarText">Share with Other</span>
                          </button>
  
                          <button
                             onClick={(e)=>ShareWithMe(e)}
                            className={`sidebardmsButton ${
                              activeButton === "ShareWithMe" ? "active" : ""
                            }`}
                          >
                            <span className="sidebarIcon">
                            <img className="sidebariconssmall" src={sharewithmeicon}></img>
                              {/* <FontAwesomeIcon icon={faShareAlt} /> */}
                            </span>
                            <span className="sidebarText">Share with me</span>
                          </button>
                        </div>
                        <div id="folderContainer2"></div>
                      </div>
                      <div className="librarydata">
                        {showDeletepopup && (
                          <div className="popup">This is a small popup!</div>
                        )}
                         {/* Start Code Update by Amjad */}
                        <div className="row">
                          <div className="col-xl-7">
                        <div
                          id="selectedText"
                          style={{
                            display: "none",
                          }}
                        >
                          {selectedText ? (
                            <h6 className="mb-1 fw-bold text-dark header-title"
                              style={{
                                color: "black",
                                marginBottom: "0px",
                                fontSize: "16px",
                              }}
                            >
                              {selectedText}
                            </h6>
                          ) : (
                            <p className="sub-header font-14"></p>
                          )}
                          {dynamicContent && (
                            <p className="sub-header font-14" style={{ color: "#6c757d" }}>{dynamicContent}</p>
                          )}
                        </div>  
                        <div
                          id="breadcrumb"
                          style={{
                            display: "none",
                          }}
                        ></div>
  </div>
  <div className="col-xl-5">
  <div className="search-container position-relative">
                          <input
                            id="searchinput"
                            type="text"
                            className="search-input"
                            placeholder="Search files..."
                          />
                          <a className="searchbutton" onClick={searchFiles}>
                            <img
                              src={require("../assets/searchicon.png")}
                              alt="Search"
                              className="search-icon"
                            />
                          </a>
                        </div>
  
    </div>
    </div>
                     {/* End Code Update by Amjad */} 
  
                         <div id="files-container"></div>
                       {
                            // listorgriddata === ''  ? (
                            //   <div id="files-container"></div>
                            // ) : (
                            //   listorgriddata === 'showListView' && (
                            //     <Table
                            //     onReturnToMain={handleReturnToMain}
                            //     Currentbuttonclick={{ "buttonclickis": Myreqormyfav }}
                            //   />
                            //   )
                            // )
  
                            listorgriddata === ''  ? (
                              <div id="files-container"></div>
                            ) : (
                              listorgriddata === 'showListView' && (
                                <Table
                                onReturnToMain={handleReturnToMain}
                                Currentbuttonclick={{ "buttonclickis": Myreqormyfav }}
                              />
                              )
                            )
                       }
                       
                       
                      </div>
                     <div id="createuploadfilecont" className="createuploadfilecont"> 
                     <button
                        className="mybutton1"
                        id="createFileButton"
                        onClick={() => handleButtonClickShow("UploadFile")}
                      >
                        + Create File
                      </button>
                      <button
                        className="mybutton2"
                        id="createFileButton2"
                        onClick={() => handleButtonClickShow("CreateFolder")}
                      >
                        + Create Folder
                      </button>
                      </div>
                <div className="Manageworkflow">
             
                </div>
                    </div>
                  ) : (
                    <div>
                      {activeComponent === "UploadFile" && (
                        <UploadFile
                        currentfolderpath={{
                           "Entity" : currentEntity,
                           "Entityurl": currentEntityURL,
                           "siteID": currentsiteID,
                           "Devision":  currentDevision,
                           "Department" : currentDepartment,
                           "DocumentLibrary": currentDocumentLibrary,
                           "Folder" :currentFolder,
                           "folderpath": currentfolderpath
                          }}
                          onReturnToMain={handleReturnToMain}
                        />
                      )}
                        {activeComponent === "CreateFolder" && (
                      <CreateFolder  OthProps={{
                        "Entity" : currentEntity,
                        "Entityurl": currentEntityURL,
                        "siteID": currentsiteID,
                        "Devision":  currentDevision,
                        "Department" : currentDepartment,
                        "DocumentLibrary": currentDocumentLibrary,
                        "Folder" :currentFolder,
                        "folderpath": currentfolderpath
                       }}
                       onReturnToMain={handleReturnToMain} />
                    )}
                    
                     
                    </div>
                  )}
                </div>
              </div>
            </div>
            </div>
          
     
          
    );
  };
  
  
  
  const DMSMain: React.FC<IDmsMusaibProps> = (props) =>{
    return (
      <Provider>
        <ArgPoc  props={props}/>
      </Provider>
    );
  };
  
  export default DMSMain;
  