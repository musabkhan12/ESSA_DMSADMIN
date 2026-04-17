declare global {
  interface Window {
    // managePermission:(documentLibraryName:string,SiteName:string,SiteID:string)=> void;
    // manageWorkflow:(documentLibraryName:string,SiteName:string,SiteID:string)=> void;
    // view:(message:string) => void;
    // PreviewFile: (path: string, siteID: string, docLibName:any) => void;
    // deleteFile:(fileId: string , siteID:string, IsHardDelete:any, listToUpdate:any ) => void;
  }
}
interface UploadFileProps {
  currentfolderpath: {
    CurrentEntity: string;
    currentEntityURL: string;
    currentsiteID: string;
    // ... other properties
  };
}

// @ts-ignore
import * as React from "react";
// import { graph } from "@pnp/graph/presets/all";
import "@pnp/graph/groups";
import { getSP, getGraphClient } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "bootstrap/dist/css/bootstrap.min.css";
import { MSGraphClient } from "@microsoft/sp-http";
import styles from './Form.module.scss'
// import "bootstrap//dist/"
import "../../../CustomCss/mainCustom.scss";
import "../../verticalSideBar/components/VerticalSidebar.scss";
import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";
import UserContext from "../../../GlobalContext/context";
// import { useState , useEffect } from "react";
import Provider from "../../../GlobalContext/provider";
import { useMediaQuery } from "react-responsive";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/sites"
import "@pnp/sp/presets/all"
import "@pnp/sp/site-groups";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { spfi, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import { PermissionKind } from "@pnp/sp/security";
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import "../../verticalSideBar/components/VerticalSidebar.scss";
import "./dmscss";
import "./DMSAdmincss"
import Manageuserpermissioninonego from "./Manageuserpermissioninonego";
import {ManageSuperAdmin} from "./ManageSuperAdmin"
import ManageEssaCoreGroups from "./ManageEssaCoreGroups"
import { useState, useRef, useEffect } from "react";

// import {IDmsMusaibProps} from './IDmsMusaibProps'
import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
import EntityMapping from "./EntityMapping";
import Devision from "./Division";
import Department from "./Department";
// import CreateEntity from "./CreateEntity";
import CreateEntity from './Entity'
import Select from "react-select";
import Swal from 'sweetalert2';
import { ManagePermission } from "./Managepermission";
import { ManageSuper } from "./ManageSuper";
// import { MSGraphClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import ManageFolderDeligation from "./ManageFolderDeligation";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faBars } from "@fortawesome/free-solid-svg-icons";
interface IMyComponentProps {
  context: WebPartContext;
}
let superadmin = require('../assets/superadmin.svg');
let managepermission = require('../assets/managepermission.png');


// let selectedUsersForPermission:any;
let selectedGroupUsers: any[];
let superA = false;
let selectedEntityForPermission: any;
let selectedGropuForPermission: any;
let selectedUsersForPermission: any;

let usersFromGroups: any[] = [];

let superAdminArray: any[];
// let IsAdmin=false;

let groupDetails: any;

let SITEID: any
let WEBID: any
let LISTID: any

interface IDmsAdminComponentProps {
  context: WebPartContext;
  someOtherProp?: any;
}
interface IDmsMusaibProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  siteUrl: string;
  context: WebPartContext;
  someOtherProp?: any;
}
// const Dmsadmincomponent = ({ props }: any) => {
//   console.log(props , " here is my prop")

const Dmsadmincomponent: React.FC<IDmsAdminComponentProps> = ({ context, someOtherProp }) => {
  const sp: SPFI = getSP();
  const data = async () => {
    const d = await sp.web()
    console.log(d, " here is my web")

  }
  data()
  const { useHide }: any = React.useContext(UserContext);
  const elementRef = React.useRef<HTMLDivElement>(null);

  // New Code
  const [allUsersFromGroups, setAllUsersFromGroups] = useState<any[]>([]);
  // console.log("allUsersFromGroups outside select entity",allUsersFromGroups);
  const [toggleManagePermission, setToggleManagePermission] = useState('Yes');
  const [adminPermissionEntity, setAdminPermissionEntity] = useState<any[]>([]);
  const [user, setUser] = useState<any[]>([]);
  const [groups, setGroups] = useState<any[]>([]);
  const [IsSuperAdmin, setIsSuperAdmin] = useState(false);

  const [toggleManagePermissionCard, setToggleManagePermissionCard] = useState("No");
  const manageUserAndPermissionImage = require('../assets/manage-users.png');
  const managefolderdeligation = require('../assets/manage-f.png');
  const managesuperadmin = require('../assets/man-s_ad.png');
   const managesuperadmin1 = require('../assets/managesuperadmin1.png');
    const managesuperadmin2 = require('../assets/managesuperadmin2.png');
  const sitep = require('../assets/site_permission_new.png');
  const viewi = require('../assets/viewocpn.png');
  //aman 27/2/26
  const [activeSp, setActiveSp] = useState<SPFI>(sp);
  // srs 26/2/26
  const [siteCollections, setSiteCollections] = useState<any[]>([]);
  const [selectedGlobalSite, setSelectedGlobalSite] = useState<any>(null);
  // Aman 16/03/26: State for table data
  const [locationData, setLocationData] = useState<any[]>([]);
  // Aman 16/03/26: States for creating new site collection
  const [showCreateForm, setShowCreateForm] = useState(false);
  const [newSiteData, setNewSiteData] = useState({ Title: '', Siteurl: '', IsActive: true });
  const [validationErrors, setValidationErrors] = useState({ location: false, department: false, groups: false, users: false });

  const fetchSiteCollections = async () => {
    try {
      // Fetches Title and SiteURL from the MasterSiteCollection list
      const items = await sp.web.lists.getByTitle("MasterSiteCollection").items.select("Title", "SiteURL", "Id", "IsActive").filter("IsActive eq 1")(); //aman 16/03/26 change to fetch only active site collections
      const options = items.map((item) => ({
        value: item.SiteURL, // Use the URL as the value for logic
        label: item.Title,   // Use Title for the display label
        id: item.Id,
        siteUrl: item.SiteURL
      }));
      setSiteCollections(options);
    } catch (error) {
      console.error("Error fetching MasterSiteCollection:", error);
    }
  };

  // aman 16/03/26: Fetch data from MasterSiteCollection list to show in the table
  const fetchMasterSiteData = async () => {
    try {
      const items = await sp.web.lists.getByTitle("MasterSiteCollection")
        .items.select("Id", "Title", "SiteURL", "SharewithOtherMeMasterSite", "IsActive")();
      setLocationData(items);
      setActiveComponent('Create Location');
    } catch (error) {
      console.error("Error:", error);
    }
  };

  // Aman 16/03/26: Toggle IsActive status
  const handleToggleStatus = async (id: number, currentStatus: boolean) => {
    try {
      await sp.web.lists.getByTitle("MasterSiteCollection").items.getById(id).update({
        IsActive: !currentStatus
      });
      Swal.fire("Updated", "Status changed", "success");
      fetchMasterSiteData();
      fetchSiteCollections();
    } catch (error) {
      console.error("Error:", error);
    }
  };

  // Aman 16/03/26: Function to save new site and auto-generate URL
  const handleSaveNewSite = async () => {
    if (!newSiteData.Title) {
      Swal.fire("Error", "Please enter a title", "error");
      return;
    }
    if (!newSiteData.Siteurl) {
      Swal.fire("Error", "Please enter a Siteurl", "error");
      return;
    }
    try {
      // URL auto generation: spaces remove karke lower case mein
      const generatedUrl = `https://officeindia.sharepoint.com/sites/${newSiteData.Title.replace(/\s+/g, '')}`;

      await sp.web.lists.getByTitle("MasterSiteCollection").items.add({
        Title: newSiteData.Title,
        // SiteURL: generatedUrl,
        SiteURL: newSiteData.Siteurl,



        IsActive: newSiteData.IsActive,
        // Aman 16/03/26: Dynamic path for share list
        SharewithOtherMeMasterSite: `${generatedUrl}/Lists/DMSShareWithOtherMaster/AllItems.aspx`
      });

      Swal.fire("Success", "New Site Collection added", "success");
      setShowCreateForm(false);
      setNewSiteData({ Title: '', Siteurl: '', IsActive: true }); // Reset form
      fetchMasterSiteData(); // Refresh table data
    } catch (error) {
      console.error("Error adding site:", error);
      Swal.fire("Error", "Could not add site", "error");
    }
  };

  //after sourish 26/2/26 change aman changes in handleGlobalSiteChange function 27/2/26
  const handleGlobalSiteChange = async (selected: any, preserveManageCard = false) => {
     //adddhyan - 03/04/2026 start
    setSelectedGlobalSite(selected);
    // Clear location error when location is selected
    if (selected) {
      setValidationErrors(prev => ({ ...prev, location: false }));
    }
 //adddhyan - 03/04/2026 start
    if (preserveManageCard) {
      // Keep the Manage Users/Permission panel open and cards hidden
      setToggleManagePermissionCard("Yes");
      setToggleManagePermission("No");
    } else {
      setActiveComponent('');           // Aman 16/03/26 
      setToggleManagePermissionCard("No"); // Ritik 25-03-2026 added code to reset the states when site collection changes
      setToggleManagePermission("Yes"); // Ritik 25-03-2026 added code to reset the states when site collection changes
    }
    //adddhyan - 03/04/2026 start

    if (selected && selected.siteUrl) {
      try {
        const targetSp = spfi(selected.siteUrl).using(SPFx(context));

        // ✅ REQUIRED: update active SP instance
        setActiveSp(targetSp);
        setUser([]);                   //Aman 2/3/26 change clear user state when site collection changes to avoid showing users from previous site collection
        await fetchUsers(targetSp);    //Aman 2/3/26 change fetch users for the selected site collection

        // 1. Fetch the cards for the new site
        await getmasterlis(targetSp);

        // 2. Fetch Entities (MasterSiteURL list)
        const entityDetails = await targetSp.web.lists
          .getByTitle("MasterSiteURL")
          .items.select("SiteURL", "Title", "Active", "SiteID")
          .filter(`Active eq 'Yes'`)();

        const entityArray = entityDetails.map((entity) => ({
          value: entity.Title,
          label: entity.Title,
          SiteID: entity.SiteID,
          SiteURL: entity.SiteURL,
        }));

        setAdminPermissionEntity(entityArray);

      } catch (error) {
        console.error("Error connecting to target site:", error);
        Swal.fire(
          "Access Denied",
          "Ensure the target site has the required lists.",
          "error"
        );
      }
    } else {
      setToggleManagePermissionCard("No"); // Ritik 25-03-2026 added code to reset the states when site collection changes
      setToggleManagePermission("Yes"); // Ritik 25-03-2026 added code to reset the states when site collection changes
      // ✅ Reset back to host site (ESSA)
      setActiveSp(sp);
      getmasterlis(sp);
      getDetailsOfAdmin();
    }
  };



  useEffect(() => {
    fetchSiteCollections();
  }, []);

// aman change start 3/4/26 (NEW useEffect)
useEffect(() => {
  if (siteCollections.length === 0) return;

  const loadCards = async () => {
    try {
      const firstSite = siteCollections[0];

      const targetSp = spfi(firstSite.siteUrl).using(SPFx(context));

      await getmasterlis(targetSp);

    } catch (error) {
      console.error("Error loading cards:", error);
    }
  };

  loadCards();
}, [siteCollections]);
// aman change end 3/4/26
  // addhyan - 01/04/2026 
  const handleViewSiteFromTable = async (item: any) => {
    const selected = {
      value: item.SiteURL,
      label: item.Title,
      siteUrl: item.SiteURL,
      id: item.Id,
    };
    await handleGlobalSiteChange(selected);
  };

  const handleBackToSettings = async () => {
    await handleGlobalSiteChange(null);
  };
  // addhyan - 01/04/2026 

  const handleToggleCard = (event: any, name: any) => {
    event.preventDefault();
    // alert(name)
    setToggleManagePermissionCard(name);
    setToggleManagePermission("No");
    setActiveComponent(null);

  }

  const handleToggleSuper = (event: any, name: any) => {
    event.preventDefault();
    setToggleManagePermission("No");
    setActiveComponent(name);
  }
  // const [permissionTable,setPermissionTable]=useState<string >('')
  const currentUserEmailRef = useRef('');

  const getcurrentuseremail = async () => {
    // const users = await graph.groups.getById("group-id").members();
    // console.log(users);
    // const graphClient: MSGraphClientV3 = await getGraphClient(context);
    const graphClient = await context.msGraphClientFactory.getClient("3");
    console.log(graphClient, "graphClient ")
    //  let groupId = '2ff98c8a-bbff-4615-bbe9-d87f0e486e33'
    console.log(graphClient, 'graphClient initialized');
    // try {
    //   const data = "https://login.microsoftonline.com/79a9a17c-1d27-470f-bf5a-3b542a3563ab/oauth2/v2.0/token"
    // } catch (error) {

    // }

    //       const groupId = '2e60dbe9-8981-4d64-b19c-2c763f202f1d'; // Replace with your specific group ID
    // const response = await graphClient
    //   .api(`groups`)
    //   .version('v1.0')
    //   .get();
    // // console.log(response, "Specific group details"); 
    // //        const response = await graphClient
    // //          .api('/groups')
    // //          .version('v1.0').get();
    //          console.log(response, "response");
    //  const response2 = await graphClient.api(`/groups/${groupId}/members`).version('v1.0').get();
    //  console.log(response2, "response 22");
    //  // Fetch groups from Microsoft Graph API
    //  const response = await graphClient
    //  .api("/groups")
    //  .version("v1.0")
    //  .select("id,displayName,mail")
    //  .get();

    const userdata = await sp.web.currentUser();
    currentUserEmailRef.current = userdata.Email;
    getDetailsOfSuperAdmin();
    getDetailsOfAdmin();
    fetchUsers(sp);
  }

  const getDetailsOfSuperAdmin = async () => {
    try {
      const usersFromDMSSuperAdmin = await sp.web.siteGroups.getByName('DMSSuper_Admin').users();

      superAdminArray = usersFromDMSSuperAdmin;
      usersFromDMSSuperAdmin.forEach((user) => {
        console.log("current user email from super admin", currentUserEmailRef.current);
        console.log("user email from super admin group", user.Email);
        if (user.Email === currentUserEmailRef.current) {
          superA = true;
          console.log(superA, "is super admin");
          setIsSuperAdmin(true);
          // setToggleManagePermission('Yes');
        }
      })
      console.log("usersFromDMSSuperAdmin", usersFromDMSSuperAdmin);
    } catch (error) {
      console.log("error in getting the details of super admin", error);
    }
  }
  console.log("IsSuperAdmin", IsSuperAdmin);
  const getDetailsOfAdmin = async () => {
    try {
      const entityDetails = await sp.web.lists.getByTitle("MasterSiteURL").items.select("SiteURL", "Title", "Active", "SiteID").filter(`Active eq 'Yes'`)();
      console.log("entityDetails", entityDetails);
      let entityArray: any[] = []
      const subsiteAdminDetails = await Promise.all(
        entityDetails.map(async (entity) => {
          try {
            const subsiteContext = await sp.site.openWebById(entity.SiteID);
            const usersFromAdmin = await subsiteContext.web.siteGroups.getByName(`${entity.Title}_Admin`).users();
            console.log("IsSuperAdmin from entityDeatils forEach", IsSuperAdmin, superA, usersFromAdmin);
            if (superA || usersFromAdmin.length !== 0) {
              if (usersFromAdmin.length !== 0 && superA === false) {
                // usersFromAdmin.forEach((user)=>{
                for (const user of usersFromAdmin) {
                  if (user.Email === currentUserEmailRef.current) {
                    console.log("current user is Admin", currentUserEmailRef.current);
                    console.log("entity", entity.Title);
                    console.log("all users in the admin group", usersFromAdmin);
                    entityArray.push({
                      value: entity.Title,
                      label: entity.Title,
                      SiteID: entity.SiteID
                    });
                    return usersFromAdmin;
                  }
                }
              }
              if (superA) {
                // console.log("current user is super admin",IsSuperAdmin);
                // console.log("user is super admin",currentUserEmailRef.current);
                // console.log("entity",entity.Title);
                // console.log("users from the admin group",usersFromAdmin);
                entityArray.push({
                  value: entity.Title,
                  label: entity.Title,
                  SiteID: entity.SiteID
                });
                return usersFromAdmin;
              }
            }
            // entityArray.push({
            //   value:entity.Title,
            //   label:entity.Title,
            //   SiteID:entity.SiteID
            // });
            // return usersFromAdmin;
          } catch (error) {
            // If the error is permission-related, return an undefined instead of throwing
            console.log("Error in getting group users. Returning empty array for", entity.Title, error);
            return undefined;
          }
        })
      );
      console.log("subsiteAdminDetails", subsiteAdminDetails);
      console.log("entityArray", entityArray);
      let finalUserArray: any[] = [];
      subsiteAdminDetails.forEach((userArray) => {
        if (userArray) {
          userArray.forEach((user) => {
            // Push the desired object structure into the result array
            finalUserArray.push({
              email: user.Email,
              Id: user.Id,
              value: user.Title,
              label: user.Title
            });
          });
        }
      });
      // Set a flag if all elements are undefined or null
      const allUndefined = subsiteAdminDetails.every(userArray =>
        userArray === undefined || (Array.isArray(userArray) && userArray.every(user => !user))
      );
      if (!allUndefined) {
        // IsAdmin=true;
        setToggleManagePermission("Yes");
      }

      setAdminPermissionEntity(entityArray);
      console.log("allUndefined", allUndefined);
      console.log("finalUserArray", finalUserArray);
    } catch (error) {
      console.log("error getting entity details", error);
    }
  }



  // const fetchUsers = async (spContext: SPFI) => {
  //  const users = await spContext.web.siteUsers();
  //   console.log("users fetch from the site",user);
  //     const usersArray=user.map((user)=>(
  //           {
  //             id:String(user.Id),
  //             value: user.Title,
  //             email: user.Email,
  //             label:user.Title,
  //             loginName:user.LoginName
  //           }
  //     ))
  //     console.log("site users",usersArray);
  //     setUser(usersArray);
  // }

  //Aman 2/3/26 change for fetch selected site collection user in the user field of manage users & Permission
  const fetchUsers = async (spContext: SPFI) => {
    const users = await spContext.web.siteUsers();

    console.log("users fetch from the site", users);
   
    //  Sirf users (group hata diya)
  const onlyUsers = users.filter(
    (u: any) => u.PrincipalType === 1
  );

    // const usersArray = users.map((u: any) => ({
        const usersArray = onlyUsers.map((u: any) => ({

    
      id: String(u.Id),
      value: u.Title,
      email: u.Email,
      label: u.Title,
      loginName: u.LoginName,
    }));

    setUser(usersArray);
  };

  // handle entity Select
  // const handleEntitySelect=async(selectedEntity:any)=>{
  //     console.log("selectedEntity",selectedEntity);
  //     selectedEntityForPermission=selectedEntity;
  //     const subsiteContext = await sp.site.openWebById(selectedEntity.SiteID);
  //     if(IsSuperAdmin){
  //       try {
  //         const usersFromAdmin = await subsiteContext.web.siteGroups.getByName(`${selectedEntity.value}_Admin`).users();
  //         console.log("usersFromAdmin -> IsSuperAdmin",usersFromAdmin);
  //         // console.log("IsAdmin",IsAdmin);
  //         // Check if current user email exists in the usersFromAdmin list
  //         const emailExists = usersFromAdmin.some(user => user.Email.toLowerCase() === currentUserEmailRef.current.toLowerCase());
  //         console.log("emailExists",emailExists);
  //         if(!emailExists){
  //           console.log(`${currentUserEmailRef.current} does not exist in the list. Adding to the admin group.`);
  //           const userObj = await sp.web.ensureUser(currentUserEmailRef.current);
  //           console.log("userObj",userObj);
  //           const users=await subsiteContext.web.siteGroups.getByName(`${selectedEntity.value}_Admin`).users.add(userObj.data.LoginName);
  //           console.log(`User Added Succecssfully in the ${selectedEntity.value}_Admin`,users);
  //         }else{
  //           console.log(`${currentUserEmailRef.current} already exists in the list.`);
  //         }

  //       } catch (error) {
  //         console.log(`Error in Adding super admin to the ${selectedEntity.value}_Admin gropup`,error)
  //       }
  //     }

  //     // Fetch all the groups in the subsite
  //     interface IMember {
  //       PrincipalType: number;
  //       Title:String;
  //       Id:number 
  //     }
  //     interface IRoleAssignmentInfo {
  //       Member?: IMember; 
  //     }
  //     const groups3:IRoleAssignmentInfo[] = await subsiteContext.web.roleAssignments.expand("Member")();
  //     console.log("groups3",groups3);
  //     // const filteredMembers=groups3.filter(roleAssignment => {
  //     //   return roleAssignment.Member.PrincipalType === 8;
  //     // });

  //     // const filteredGroups = filteredMembers.map((object) => ({
  //     //     value: object.Member.Title,
  //     //     label: object.Member.Title,
  //     //     Id: object.Member.Id,
  //     // }));

  //             // const filteredMembers=groups3.filter(roleAssignment => {
  //     //   return roleAssignment.Member.PrincipalType === 8;
  //     // });

  //     // const filteredGroups = filteredMembers.map((object) => ({
  //     //     value: object.Member.Title,
  //     //     label: object.Member.Title,
  //     //     Id: object.Member.Id,
  //     // }));

  //       const filteredMembers = groups3
  //       .filter(roleAssignment => {
  //         const title = roleAssignment.Member.Title;

  //         return title.endsWith("_Admin") || title.endsWith("_Read") || title.endsWith("_Contribute") || title.endsWith("_Approval");
  //       })
  //       .map(roleAssignment => roleAssignment.Member);


  //       const filteredGroups = filteredMembers.map(member => ({
  //       value: member.Title,
  //       label: member.Title,
  //       Id: member.Id,
  //       }));

  //       console.log(filteredGroups);
  //     console.log("filteredGroups",filteredGroups);
  //     console.log("filteredMembers",filteredMembers);
  //     // filter the DMSSuper_Admin
  //     const filteredRoles = filteredGroups.filter(role => role.value !== "DMSSuper_Admin");
  //     console.log("filteredRoles before permission",filteredRoles);
  //     // let usersFromGroups:any[]=[];
  //     usersFromGroups=[];
  //     // const addedNewInfoToGroups=filteredRoles.map(async(group)=>{
  //     //     const result=group.value.split("_")[1];
  //     //     // console.log("result",result);
  //     //     if(result === "Admin"){
  //     //       (group as any).permission="Admin";
  //     //       (group as any).Description="Full Control - Has full control.";
  //     //       const usersFromAdmin = await subsiteContext.web.siteGroups.getByName(`${group.value}`).users();
  //     //       console.log("usersFromAdmin",usersFromAdmin);
  //     //       usersFromAdmin.forEach((user)=>{
  //     //         const userObject={
  //     //           user:"",
  //     //           groupName:"",
  //     //           permission:"",
  //     //           Descirption:"",
  //     //           email:"",
  //     //         }
  //     //         userObject.user=user.Title;
  //     //         userObject.email=user.Email;
  //     //         userObject.groupName=`${group.value}`;
  //     //         userObject.permission="Admin";
  //     //         userObject.Descirption="Full Control - Has full control.";
  //     //         (userObject as any).userId=user.Id
  //     //         usersFromGroups.push(userObject);
  //     //       })
  //     //     }else if(result === "Contribute"){
  //     //       (group as any).permission="Contribute";
  //     //       (group as any).Description="Contribute - Can view, add, update, and delete documents.";
  //     //       const usersFromContribute = await subsiteContext.web.siteGroups.getByName(`${group.value}`).users();
  //     //       console.log("usersFromContribute",usersFromContribute);
  //     //       usersFromContribute.forEach((user)=>{
  //     //         const userObject={
  //     //           user:"",
  //     //           groupName:"",
  //     //           permission:"",
  //     //           Descirption:"",
  //     //           email:"",
  //     //         }
  //     //         userObject.user=user.Title;
  //     //         userObject.email=user.Email;
  //     //         userObject.groupName=`${group.value}`;
  //     //         userObject.permission="Contribute";
  //     //         userObject.Descirption="Contribute - Can view, add, update, and delete documents.";
  //     //         (userObject as any).userId=user.Id
  //     //         usersFromGroups.push(userObject);
  //     //       })
  //     //     }else if(result === "Read"){
  //     //       (group as any).permission="Read";
  //     //       (group as any).Description="Read - Can view pages and download documents."
  //     //       const usersFromRead = await subsiteContext.web.siteGroups.getByName(`${group.value}`).users();
  //     //       console.log("usersFromRead",usersFromRead);
  //     //       usersFromRead.forEach((user)=>{
  //     //         const userObject={
  //     //           user:"",
  //     //           groupName:"",
  //     //           permission:"",
  //     //           Descirption:"",
  //     //           email:"",
  //     //         }
  //     //         userObject.user=user.Title;
  //     //         userObject.email=user.Email;
  //     //         userObject.groupName=`${group.value}`;
  //     //         userObject.permission="Read";
  //     //         userObject.Descirption="Read - Can view pages and download documents.";
  //     //         (userObject as any).userId=user.Id
  //     //         usersFromGroups.push(userObject);
  //     //       })
  //     //     }else if(result === "View"){
  //     //       (group as any).permission="View";
  //     //       (group as any).Description="View - Can only view content."
  //     //       const usersFromView = await subsiteContext.web.siteGroups.getByName(`${group.value}`).users();
  //     //       console.log("usersFromView",usersFromView);
  //     //       usersFromView.forEach((user)=>{
  //     //         const userObject={
  //     //           user:"",
  //     //           groupName:"",
  //     //           permission:"",
  //     //           Descirption:"",
  //     //           email:"",
  //     //         }
  //     //         userObject.user=user.Title;
  //     //         userObject.email=user.Email;
  //     //         userObject.groupName=`${group.value}`;
  //     //         userObject.permission="View";
  //     //         userObject.Descirption="View - Can only view content.";
  //     //         (userObject as any).userId=user.Id
  //     //         usersFromGroups.push(userObject);
  //     //       })
  //     //     }else if(result === "Initiator"){
  //     //       (group as any).permission="Initiator";
  //     //       (group as any).Description="Initiator - Can view, add, update and delete documents."
  //     //       const usersFromInitiator = await subsiteContext.web.siteGroups.getByName(`${group.value}`).users();
  //     //       console.log("usersFromInitiator",usersFromInitiator);
  //     //       usersFromInitiator.forEach((user)=>{
  //     //         const userObject={
  //     //           user:"",
  //     //           groupName:"",
  //     //           permission:"",
  //     //           Descirption:"",
  //     //           email:"",
  //     //         }
  //     //         userObject.user=user.Title;
  //     //         userObject.email=user.Email;
  //     //         userObject.groupName=`${group.value}`;
  //     //         userObject.permission="Initiator";
  //     //         userObject.Descirption="Initiator - Can view, add, update and delete documents.";
  //     //         (userObject as any).userId=user.Id
  //     //         usersFromGroups.push(userObject);
  //     //       })
  //     //     }else if(result === "Approval"){
  //     //       (group as any).permission="Approval";
  //     //       (group as any).Description="Approval - Can view, add, update and delete documents."
  //     //       const usersFromApproval = await subsiteContext.web.siteGroups.getByName(`${group.value}`).users();
  //     //       console.log("usersFromApproval",usersFromApproval);
  //     //       usersFromApproval.forEach((user)=>{
  //     //         const userObject={
  //     //           user:"",
  //     //           groupName:"",
  //     //           permission:"",
  //     //           Descirption:"",
  //     //           email:"",
  //     //         }
  //     //         userObject.user=user.Title;
  //     //         userObject.email=user.Email;
  //     //         userObject.groupName=`${group.value}`;
  //     //         userObject.permission="Approval";
  //     //         userObject.Descirption="Approval - Can view, add, update and delete documents.";
  //     //         (userObject as any).userId=user.Id
  //     //         usersFromGroups.push(userObject);
  //     //       })
  //     //     }else if(result === "AllUsers"){
  //     //       (group as any).permission="AllUsers";
  //     //       (group as any).Description=" AllUsers - Can view, add, update and delete documents."
  //     //       const usersFromAllUsers = await subsiteContext.web.siteGroups.getByName(`${group.value}`).users();
  //     //       console.log("usersFromAllUsers",usersFromAllUsers);
  //     //       usersFromAllUsers.forEach((user)=>{
  //     //         const userObject={
  //     //           user:"",
  //     //           groupName:"",
  //     //           permission:"",
  //     //           Descirption:"",
  //     //           email:"",
  //     //         }
  //     //         userObject.user=user.Title;
  //     //         userObject.email=user.Email;
  //     //         userObject.groupName=`${group.value}`;
  //     //         userObject.permission="AllUsers";
  //     //         userObject.Descirption="AllUsers - Can view, add, update and delete documents.";
  //     //         (userObject as any).userId=user.Id
  //     //         usersFromGroups.push(userObject);
  //     //       })
  //     //     }else{
  //     //       (group as any).permission="Unknown";
  //     //       (group as any).Description="UnKnown Role"
  //     //       const usersFromUnKnown = await subsiteContext.web.siteGroups.getByName(`${group.value}`).users();
  //     //       console.log(`usersFromUnKnown_${group.value}`,usersFromUnKnown);

  //     //       usersFromUnKnown.forEach((user)=>{
  //     //         const userObject={
  //     //           user:"",
  //     //           groupName:"",
  //     //           permission:"",
  //     //           Descirption:"",
  //     //           email:"",
  //     //         }
  //     //         userObject.user=user.Title;
  //     //         userObject.email=user.Email;
  //     //         userObject.groupName=`${group.value}`;
  //     //         userObject.permission="Unknown";
  //     //         userObject.Descirption="UnKnown Role";
  //     //         (userObject as any).userId=user.Id
  //     //         usersFromGroups.push(userObject);
  //     //       })
  //     //     }

  //     //     return group;
  //     // })

  //     await Promise.all(filteredRoles.map(async (group) => {
  //       const result = group.value.split("_")[1];
  //       let permission = "";
  //       let description = "";

  //       // Determine permission and description based on result
  //       switch (result) {
  //         case "Admin":
  //           permission = "Admin";
  //           description = "Full Control - Has full control.";
  //           break;
  //         case "Contribute":
  //           permission = "Contribute";
  //           description = "Can view, add, update, and download documents.";
  //           break;
  //         case "Read":
  //           permission = "Read";
  //           description = "Can view pages and download documents.";
  //           break;
  //         case "View":
  //           permission = "View";
  //           description = "Can only view content.";
  //           break;
  //         case "Initiator":
  //           permission = "Initiator";
  //           description = "Can view, add, update, and download documents.";
  //           break;
  //         case "Approval":
  //           permission = "Approval";
  //           description = "Can view, add, update, and download documents.";
  //           break;
  //         case "AllUsers":
  //           permission = "AllUsers";
  //           description = "Can view, add, update, and download documents.";
  //           break;
  //         default:
  //           permission = "Unknown";
  //           description = "Unknown role";
  //           break;
  //       }

  //       (group as any).permission=permission;
  //       (group as any).Description=description;

  //       // Fetch users based on group name
  //       const users = await subsiteContext.web.siteGroups.getByName(`${group.value}`).users();

  //       // Add each user to the usersFromGroups array
  //       users.forEach((user) => {
  //         usersFromGroups.push({
  //           user: user.Title,
  //           email: user.Email,
  //           groupName: group.value,
  //           permission:permission,
  //           Descirption: description,
  //           userId: user.Id,
  //         });
  //       });
  //     }));

  //     console.log("allUsersFromGroups",usersFromGroups);
  //     console.log("filteredRoles",filteredRoles);
  //     // New code start filter the group and its description
  //     groupDetails=filteredRoles.find(item => item.value === `${selectedEntity.value}_Admin`);
  //     console.log("groupDetails",groupDetails);
  //     selectedGropuForPermission=groupDetails;
  //     try {
  //       const subsiteContext = await sp.site.openWebById(selectedEntityForPermission.SiteID);
  //       const usersFromSelectedGroups = await subsiteContext.web.siteGroups.getByName(`${groupDetails.value}`).users();
  //       console.log("usersFromSelectedGroups",usersFromSelectedGroups);
  //       selectedGroupUsers=usersFromSelectedGroups;

  //         const showUsersFromGroupsOnTable=usersFromSelectedGroups.map((user)=>{
  //           return {
  //             user: user.Title,
  //             email: user.Email,
  //             groupName: groupDetails.value,
  //             permission:groupDetails.permission,
  //             Descirption: groupDetails.Description,
  //             userId: user.Id,
  //           }
  //         })
  //       console.log("showUsersFromGroupsOnTable1",showUsersFromGroupsOnTable);
  //       setAllUsersFromGroups([]);
  //       setAllUsersFromGroups(showUsersFromGroupsOnTable);
  //       setShowGroupsUsers("Yes");
  //     } catch (error) {
  //       console.log("error from getting the users from the groups after selecting the groups",error);
  //     }
  //     // end
  //     setGroups(filteredRoles);
  //     // setAllUsersFromGroups([]);
  //     // setAllUsersFromGroups(usersFromGroups);
  //     setShowGroupsTable("Yes");
  // }
  // srs 26/2/26
  const handleEntitySelect = async (selectedEntity: any) => {
    console.log("selectedEntity", selectedEntity);
    selectedEntityForPermission = selectedEntity;
    // Clear department error when department is selected
    if (selectedEntity) {
      setValidationErrors(prev => ({ ...prev, department: false }));
    }

    try {
      // --- FIX: Use spfi with the Absolute URL for cross-site support ---
      const subsiteContext = spfi(selectedEntity.SiteURL).using(SPFx(context));

      if (IsSuperAdmin) {
        try {
          const adminGroupName = `${selectedEntity.value}_Admin`;
          const usersFromAdmin = await subsiteContext.web.siteGroups.getByName(adminGroupName).users();

          const emailExists = usersFromAdmin.some(user =>
            user.Email.toLowerCase() === currentUserEmailRef.current.toLowerCase()
          );

          if (!emailExists) {
            // Get user from host site and ensure them on the target subsite
            const userObj = await sp.web.ensureUser(currentUserEmailRef.current);
            await subsiteContext.web.siteGroups.getByName(adminGroupName).users.add(userObj.data.LoginName);
          }
        } catch (error) {
          console.log(`Error in adding super admin to the group`, error);
        }
      }

      // --- THE LOGIC BELOW REMAINS EXACTLY AS YOUR ORIGINAL ---
      interface IMember {
        PrincipalType: number;
        Title: String;
        Id: number;
      }
      interface IRoleAssignmentInfo {
        Member?: IMember;
      }

      const groups3: IRoleAssignmentInfo[] = await subsiteContext.web.roleAssignments.expand("Member")();
      console.log("groups3", groups3);

      const filteredMembers = groups3
        .filter(roleAssignment => {
          const title = roleAssignment.Member.Title;
          return title.endsWith("_Admin") || title.endsWith("_Read") || title.endsWith("_Contribute") || title.endsWith("_Approval");
        })
        .map(roleAssignment => roleAssignment.Member);

      const filteredGroups = filteredMembers.map(member => ({
        value: member.Title,
        label: member.Title,
        Id: member.Id,
      }));

      const filteredRoles = filteredGroups.filter(role => role.value !== "DMSSuper_Admin");

      usersFromGroups = [];

      await Promise.all(filteredRoles.map(async (group) => {
        const result = group.value.split("_")[1];
        let permission = "";
        let description = "";

        switch (result) {
          case "Admin":
            permission = "Admin";
            description = "Full Control - Has full control.";
            break;
          case "Contribute":
            permission = "Contribute";
            description = "Can view, add, update, and download documents.";
            break;
          case "Read":
            permission = "Read";
            description = "Can view pages and download documents.";
            break;
          case "View":
            permission = "View";
            description = "Can only view content.";
            break;
          case "Initiator":
            permission = "Initiator";
            description = "Can view, add, update, and download documents.";
            break;
          case "Approval":
            permission = "Approval";
            description = "Can view, add, update, and download documents.";
            break;
          case "AllUsers":
            permission = "AllUsers";
            description = "Can view, add, update, and download documents.";
            break;
          default:
            permission = "Unknown";
            description = "Unknown role";
            break;
        }

        (group as any).permission = permission;
        (group as any).Description = description;

        const users = await subsiteContext.web.siteGroups.getByName(`${group.value}`).users();

        users.forEach((user) => {
          usersFromGroups.push({
            user: user.Title,
            email: user.Email,
            groupName: group.value,
            permission: permission,
            Descirption: description,
            userId: user.Id,
          });
        });
      }));

      groupDetails = filteredRoles.find(item => item.value === `${selectedEntity.value}_Admin`);
      selectedGropuForPermission = groupDetails;

      try {
        // Using subsiteContext instead of sp.site.openWebById
        const usersFromSelectedGroups = await subsiteContext.web.siteGroups.getByName(`${groupDetails.value}`).users();
        selectedGroupUsers = usersFromSelectedGroups;

        const showUsersFromGroupsOnTable = usersFromSelectedGroups.map((user) => {
          return {
            user: user.Title,
            email: user.Email,
            groupName: groupDetails.value,
            permission: groupDetails.permission,
            Descirption: groupDetails.Description,
            userId: user.Id,
          }
        });

        setAllUsersFromGroups(showUsersFromGroupsOnTable);
        setShowGroupsUsers("Yes");
      } catch (error) {
        console.log("error from getting users after selection", error);
      }

      setGroups(filteredRoles);
      setShowGroupsTable("Yes");

    } catch (error) {
      console.error("Error in cross-site handleEntitySelect", error);
    }
  };
  //  handle groups select
  const handleGroupsSelect = async (selectedGrous: any) => {
    // Set selected groups start
    groupDetails = selectedGrous;
    // Clear groups error when groups is selected
    if (selectedGrous) {
      setValidationErrors(prev => ({ ...prev, groups: false }));
    }
    // End
    console.log("selectedGrous", selectedGrous);
    console.log("selectedEntityForPermission", selectedEntityForPermission);
    selectedGropuForPermission = selectedGrous;
    try {
      const subsiteContext = await sp.site.openWebById(selectedEntityForPermission.SiteID);
      const usersFromSelectedGroups = await subsiteContext.web.siteGroups.getByName(`${selectedGrous.value}`).users();
      console.log("usersFromSelectedGroups", usersFromSelectedGroups);
      selectedGroupUsers = usersFromSelectedGroups;

      const showUsersFromGroupsOnTable = usersFromSelectedGroups.map((user) => {
        return {
          user: user.Title,
          email: user.Email,
          groupName: selectedGrous.value,
          permission: selectedGrous.permission,
          Descirption: selectedGrous.Description,
          userId: user.Id,
        }
      })
      console.log("showUsersFromGroupsOnTable", showUsersFromGroupsOnTable);
      setAllUsersFromGroups([]);
      setAllUsersFromGroups(showUsersFromGroupsOnTable);
      setShowGroupsUsers("Yes");
    } catch (error) {
      console.log("error from getting the users from the groups after selecting the groups", error);
    }


  }

  const handleUsersSelect = (selectedUser: any) => {
    console.log("selectedUser", selectedUser);
    selectedUsersForPermission = selectedUser;
    // Clear users error when user selects users
    setValidationErrors(prev => ({ ...prev, users: false }));
  }

  const handleAddUsers = async () => {
    console.log("selectedUsersForPermission", selectedUsersForPermission);
    console.log("selectedGropuForPermission", selectedGropuForPermission);
    console.log("selectedEntityForPermission", selectedEntityForPermission);
    console.log("selectedGlobalSite", selectedGlobalSite);

    const errors = { location: false, department: false, groups: false, users: false };
    const missingFields = [];

    // Validate Location
    if (!selectedGlobalSite) {
      errors.location = true;
      missingFields.push('Location');
    }

    // Validate Department
    if (selectedEntityForPermission === undefined) {
      errors.department = true;
      missingFields.push('Department');
    }

    // Validate Groups
    if (selectedGropuForPermission === undefined) {
      errors.groups = true;
      missingFields.push('Groups');
    }

    // Validate Users
    if (selectedUsersForPermission === undefined || selectedUsersForPermission.length === 0) {
      errors.users = true;
      missingFields.push('Users');
    }

    // If there are errors, show them and set error states
    if (missingFields.length > 0) {
      setValidationErrors(errors);
      Swal.fire({
        icon: 'error',
        title: 'Please fill out the fields!',
        html: `<b>Please select the following fields:</b><br/>${missingFields.join(', ')}`,
        confirmButtonColor: '#d33'
      });
      return;
    }

    // const subsiteContext = await sp.site.openWebById(selectedEntityForPermission.SiteID);
    const subsiteContext = spfi(selectedEntityForPermission.SiteURL).using(SPFx(context)); //addhyan 
    //wait for all add operations to complete
    const addUsersPromises = selectedUsersForPermission.map(async (user: any) => {
      try {
        // const userObj = await sp.web.ensureUser(user.email);
        const userObj = await subsiteContext.web.ensureUser(user.email); //addhyan
        console.log("userObj", userObj);
        const users = await subsiteContext.web.siteGroups.getByName(`${selectedGropuForPermission.value}`).users.add(userObj.data.LoginName);
        console.log(`${user.email} added to the group successfully.`, users);
        
      } catch (error) {
        console.error(`Failed to add ${user.email} to the group: `, error);
      }
    });

    await Promise.all(addUsersPromises);
    // Clear all errors on success
    setValidationErrors({ location: false, department: false, groups: false, users: false });
    onSuccess(selectedGropuForPermission.value);
    // Call handleEntitySelect once all users have been added
    // to refresh the user table
    // handleEntitySelect(selectedEntityForPermission);
    // selectedUsersForPermission=undefined;
    handleGroupsSelect(selectedGropuForPermission);
  }

  const hanldeManagePermission = () => {
    if (selectedGroupUsers === undefined && selectedGropuForPermission === undefined) {
      checkValidation();
    } else {
      setActiveComponent("ManagePermission");
      setToggleManagePermission('No');
      setToggleManagePermissionCard("No");
    }
  }

  const handleReturnToMainFromPermissionTable = (Name: any) => {
    setToggleManagePermission('Yes');
    setActiveComponent(Name);
    setToggleManagePermissionCard("No");
    setShowGroupsUsers("No");
    setShowGroupsTable("No");
  }

  const checkValidation = () => {
    Swal.fire("Please fill out the fields!", "All fields are required");
  }

  const handleCreate = (Name: any) => {
    setActiveComponent(Name);
    setToggleManagePermission('No');
  }

  const handleBackToManagePermissionCard = () => {
    setToggleManagePermissionCard('Yes');
    setActiveComponent(null);
  }

  // New Code Added for Show the selceted entity Groups in table form and also  show the all users of the groups
  const [showGroupsTable, setShowGroupsTable] = useState("No");
  const [showGroupsUsers, setShowGroupsUsers] = useState("No");
  // const [refresh,setRefresh]=useState(false);

  // this function remove the user from groups
  const handleDeleteUser = async (userId: any, groupName: any) => {
    console.log("UserId", userId);
    try {

      // const subsitecontext=await sp.site.openWebById(selectedEntityForPermission.SiteID);
      const subsiteContext = spfi(selectedEntityForPermission.SiteURL).using(SPFx(context)); //addhyan 

      // Get the group by name


      const group = subsiteContext.web.siteGroups.getByName(groupName); // addhyan 
      // Remove the user from the group using their userId
      confirmDelete(group, userId, groupName);
      // await group.users.removeById(userId);
    } catch (error) {
      console.error("Error removing user from group: ", error);
    }
  }

  const confirmDelete = (group: any, userId: any, groupName: any) => {
    Swal.fire({
      title: "Are you sure?",
      text: "You won't be able to revert this!",
      icon: "warning",
      showCancelButton: true,
      confirmButtonColor: "#3085d6",
      cancelButtonColor: "#d33",
      confirmButtonText: "Yes, Removed it!"
    }).then(async (result) => {
      if (result.isConfirmed) {
        await group.users.removeById(userId);
        console.log(`User with ID ${userId} has been removed from the group '${groupName}'`);
        // to refresh the user table
        // handleEntitySelect(selectedEntityForPermission);
        handleGroupsSelect(selectedGropuForPermission);
        Swal.fire({
          title: "Removed!",
          text: `User Suucessfuly removed from ${groupName}.`,
          icon: "success"
        });
      }
    });
  }

  const onSuccess = (groupName: any) => {
    Swal.fire({
      title: "Added!",
      text: `User Added Suucessfuly to the ${groupName}.`,
      icon: "success",
    });
  }
  //  End


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

  const [Mylistdata, setMylistdata] = useState([]);
  ////
  const [activeComponent, setActiveComponent] = useState<string>('');
  ////
  console.log(activeComponent, "activeComponent")
  const handleReturnToMain = (Name: any) => {
    setActiveComponent(Name); // Reset to show the main component
    console.log(activeComponent, "activeComponent updated")
    setToggleManagePermission('Yes');
  };
  // const getmasterlis = async () => {
  //   try {
  //     const site = await sp.site();
  //     SITEID = site.Id;
  //     const web = await sp.web();
  //     WEBID = web.Id;
  //     const list = await sp.web.lists.getByTitle("DMSAdmin").select("Id")();
  //     LISTID = list.Id;
  //     // alert(`this is site ID ${site.Id} and this is web ID ${web.Id} and this is list ID ${list.Id}`);
  //     const items = await sp.web.lists.getByTitle('DMSAdmin').items();
  //     console.log(items, "getmasterlis");
  //     setMylistdata(items);

  //   } catch (error) {
  //     console.error("Error fetching list items:", error);
  //   }
  // };
  // useEffect(() => {
  //   getmasterlis();

  // }, []);
  // srs 26/2/26
  const getmasterlis = async (targetSp: SPFI = sp) => {
    try {
      // We need the SiteID and WebID of the TARGET site to construct image URLs
      const site = await targetSp.site();
      const web = await targetSp.web();

      // Get the list ID from the target site
      const list = await targetSp.web.lists.getByTitle("DMSAdmin").select("Id")();

      // Update global/state variables for the IDs so the render uses the correct ones
      SITEID = site.Id;
      WEBID = web.Id;
      LISTID = list.Id;

      // const items = await targetSp.web.lists.getByTitle('DMSAdmin').items();
      // console.log(items, "Fetched cards from target site");
      // setMylistdata(items);
      //addhyan 2/3/26 filter the items based on isActive field
      const items = await targetSp.web.lists.getByTitle('DMSAdmin').items();
      console.log(items, "Fetched cards from target site");
      //const filteredItems = items.filter(item => item.isActive === true);
      //aman change start 3/4/26
      const filteredItems = items.filter((item: any) => {
        return (
          item.isActive === true ||
          item.IsActive === true ||
          item.isActive === "Yes" ||
          item.IsActive === "Yes"
        );
      });
      //aman change end 3/4/26
      setMylistdata(filteredItems);


    } catch (error) {
      console.error("Error fetching list items from target site:", error);
      // Optionally clear cards if the target site doesn't have the DMSAdmin list
      setMylistdata([]);
    }
  };
  console.log(Mylistdata, "Mylistdata")


  useEffect(() => {
    getcurrentuseremail();

  }, []);
  const siteUrl = someOtherProp.siteUrl;
  console.log(siteUrl, "siteUrl")


  // Add pagination start
  const [currentPage, setCurrentPage] = React.useState(1);
  const itemsPerPage = 10;
  const totalPages = Math.ceil(allUsersFromGroups.length / itemsPerPage);

  const handlePageChange = (pageNumber: any) => {
    if (pageNumber > 0 && pageNumber <= totalPages) {
      setCurrentPage(pageNumber);
    }
  };

  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const currentData = allUsersFromGroups.slice(startIndex, endIndex);

  interface PaginationProps {
    currentPage: number;
    totalPages: any;
    handlePageChange: any;
  }
  const Pagination = ({ currentPage, totalPages, handlePageChange }: PaginationProps) => {
    const pageLimit = 5; // Number of visible page items

    // Determine the start and end page based on the current page and total pages
    const startPage = Math.max(1, currentPage - Math.floor(pageLimit / 2));
    const endPage = Math.min(totalPages, startPage + pageLimit - 1);

    // Adjust start page if it's too close to the end
    const adjustedStartPage = Math.max(1, Math.min(startPage, totalPages - pageLimit + 1));

    // Create an array for the visible page numbers
    const visiblePages = Array.from(
      { length: Math.min(pageLimit, totalPages) },
      (_, index) => adjustedStartPage + index
    );

    return (
      <nav className="pagination-container">
        <ul className="pagination">
          {/* Previous Button */}
          <li className={`page-item ${currentPage === 1 ? 'disabled' : ''}`}>
            <a
              className="page-link PreviousPage"
              onClick={() => handlePageChange(currentPage - 1)}
              aria-label="Previous"
            >
              «
            </a>
          </li>

          {/* Render visible page numbers */}
          {visiblePages.map((pageNumber) => (
            <li
              key={pageNumber}
              className={`page-item ${currentPage === pageNumber ? 'active' : ''}`}
            >
              <a className="page-link" onClick={() => handlePageChange(pageNumber)}>
                {pageNumber}
              </a>
            </li>
          ))}

          {/* Next Button */}
          <li className={`page-item ${currentPage === totalPages ? 'disabled' : ''}`}>
            <a
              className="page-link NextPage"
              onClick={() => handlePageChange(currentPage + 1)}
              aria-label="Next"
            >
              »
            </a>
          </li>
        </ul>
      </nav>
    );
  };
  // End

  return (
    <div id="wrapper" ref={elementRef}>
      <div
        className="app-menu"
        id="myHeader">
        <VerticalSideBar _context={sp} />
      </div>
      <div className="content-page">
        <HorizontalNavbar _context={sp} siteUrl={siteUrl} />
        <div className="content" style={{ marginLeft: `${!useHide ? '240px' : '80px'}`, marginTop: '1.5rem' }}>
          {/* addhyan 01/04/26: */}
          {/* srs 26/2/26 */}
          {/* GLOBAL SITE SELECTION BAR */}
          <div className="container-fluid">
            <h4 className="page-title fw-bold mb-1 font-20">Settings</h4>
            <div className="Route">
              {" "}
              <h2 className="Home">Home</h2>
              <span className="greater">&gt;</span>{" "}
              <h2 className="Setting">Settings </h2>{" "}
            </div>


            {/* {selectedGlobalSite ? <div className="text-secondary" >Active Site Collection  {' > '}
              {selectedGlobalSite && <span className="fw-bold text-dark">{selectedGlobalSite.label}</span>}
            </div> : ""} */}

          </div>


          {/* addhyan 01/04/26: */}

          <div className="container-fluid mt-3">
            {IsSuperAdmin ? (<>
              {activeComponent === "" ?
                (<div>
                  <div className="DMSMasterContainer">
                    {/* <h4 className="page-title fw-bold mb-1 font-20">Settings</h4>  */}



                    {/* <div className="Route">
                      {" "}
                      <h2 className="Home">Home</h2>
                      <span className="greater">&gt;</span>{" "}
                      <h2 className="Setting">Settings</h2>{" "}
                    </div> */}

                    {/* addhyan 01/04/2026 */}
                    {/* {selectedGlobalSite && ( */}
                      {/* <div className="mb-3 position-relative"> */}
                        {/* <button className="btn btn-secondary newbutton_align" onClick={handleBackToSettings}>Back to Home </button> */}
                        {/* <button className="btn btn-secondary back-to-admin" onClick={handleBackToSettings}>Back to Home</button> */}
                      {/* </div> */}
                    {/* )} */}

                    
                    <div className="row manage-master mt-0">
                      {/* {!selectedGlobalSite && ( */}
                        <div className="col-sm-3 col-md-3 mt-2">
                          <div className="card-master box1" onClick={fetchMasterSiteData}>
                            <div className="icon" style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '70px', fontSize: '32px', color: '#2575a7' }}>
                              <img className="" src={sitep} />
                            </div>
                            <p className="text-dark">Manage Location</p>
                          </div>
                        </div>
                      {/* )} */}
                      {/* addhyan 01/04/2026 end */}
                      {Mylistdata.map((item: any) => {
                        const imageData = JSON.parse(item.Image); // Assuming 'ImageColumn' is the column name
                        const itemid = String(item.Id);
                        console.log(itemid, "itemsid");
                        console.log(imageData, "imagedata");
                        const siteUrl = window.location.origin;
                        let locationPath = window.location.pathname.match(/\/sites\/[^\/]+/)[0];
                        // const imageUrl = `https://officeindia.sharepoint.com/sites/AlRostmaniSpfx2/_api/v2.1/sites('${SITEID},${WEBID}')/lists('${LISTID}')/items('${itemid}')/attachments('${imageData?.fileName}')/thumbnails/0/c3000x2000/content?prefer=noredirect,closestavailablesize`;
                        // const imageUrl = `https://officeindia.sharepoint.com/sites/AlRostmani/_api/v2.1/sites('${SITEID},${WEBID}')/lists('${LISTID}')/items('${itemid}')/attachments('${imageData?.fileName}')/thumbnails/0/c3000x2000/content?prefer=noredirect,closestavailablesize`;
                        // const imageUrl = `https://officeindia.sharepoint.com/sites/AlRostmanispfx2/_api/v2.1/sites('${SITEID},${WEBID}')/lists('${LISTID}')/items('${itemid}')/attachments('${imageData?.fileName}')/thumbnails/0/c3000x2000/content?prefer=noredirect,closestavailablesize`;
                        const imageUrl = `${siteUrl}${locationPath}/_api/v2.1/sites('${SITEID},${WEBID}')/lists('${LISTID}')/items('${itemid}')/attachments('${imageData?.fileName}')/thumbnails/0/c3000x2000/content?prefer=noredirect,closestavailablesize`;
                        console.log(imageUrl, "imageurl");
                        return (

                          //  <div className="col-sm-3 col-md-3 mt-2"> addhyan 01/04/2026
                          <div className="col-sm-3 col-md-3 mt-2" key={item.Id}>
                            <a href={item?.LinkUrl}>
                              <div className="card-master box1" onClick={() => handleCreate(item.Name)}>
                                <div className="icon">
                                  <img className="" src={imageUrl} />
                                </div>
                                <p className="text-dark">{item.Name}</p>
                              </div>
                            </a>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                </div>) : (
                  <div className="position-relative">
                    {activeComponent === 'Create Location' || activeComponent === 'Create Department' && (
                      <div>
                        <button className="btn back-to-admin" onClick={() => handleReturnToMain('')}> Back to Home </button>
                        <CreateEntity context={context}
                          siteCollections={siteCollections} />  
                      </div>

                    )}
                    {activeComponent === 'Manage user permission' && (
                      <div>
                        <button className="btn back-to-admin newbutton_align" onClick={() => handleReturnToMain('')}> Back to Home </button>
                        {/* <Manageuserpermissioninonego /> */}
                        {/* srs 17/3/26 Ritik 03/04/2026 */}
                        <Manageuserpermissioninonego sp={activeSp} siteCollections={siteCollections} context={context} />
                      </div>
                    )}
                    {activeComponent === 'Create Division' && (
                      <div className="position-relative">
                        <button className="btn back-to-admin" onClick={() => handleReturnToMain('')}> Back to Home </button>
                        {/*Aman 27/2/26*/}
                        <Devision sp={activeSp} />
                      </div>

                    )}
                    {/* Aman 16/03/26: Table for viewing data of MasterSiteCollection list */}

                    {activeComponent === 'Create Location' && (
                      <div className="position-relative mt-4">
                        {showCreateForm ? (
                          <>
                           

                            <div className="card p-4" style={{ border: "1px solid #ccc", borderRadius: "8px", background: "#fff" }}>
 <div className='page-title fw-bold mb-3 font-20'>Add Location</div>
                              <div className="row mb-4">

                                <div className="col-md-4">
                                  <label className="fw-bold mb-1" style={{ color: '#666' }}>Title <span className="text-danger">*</span></label>
                                  <input
                                    type="text"
                                    className="form-control shadow-sm"
                                    style={{ borderRadius: "6px", border: "1px solid #ccc", padding: '10px' }}
                                    onChange={(e) => setNewSiteData({ ...newSiteData, Title: e.target.value })}
                                  />

                                </div>
                                <div className="col-md-4">
                                  <label className="fw-bold mb-1" style={{ color: '#666' }}>Site Url <span className="text-danger">*</span></label>
                                  <input
                                    type="text"
                                    className="form-control shadow-sm"
                                    style={{ borderRadius: "6px", border: "1px solid #ccc", padding: '10px' }}
                                    onChange={(e) => setNewSiteData({ ...newSiteData, Siteurl: e.target.value })}
                                  />

                                </div>
                                <div className="col-md-4 ">
                                  <label className="fw-bold mb-2" style={{ color: '#666' }}>Active <span className="text-danger">*</span></label>
                                  <div className="d-flex gap-4 align-items-center">
                                    <label className="d-flex align-items-center gap-2" style={{ cursor: 'pointer', fontWeight: '500' }}>
                                      <input type="radio" name="isActive" checked={newSiteData.IsActive === true} onChange={() => setNewSiteData({ ...newSiteData, IsActive: true })} /> Yes
                                    </label>
                                    <label className="d-flex align-items-center gap-2" style={{ cursor: 'pointer', fontWeight: '500' }}>
                                      <input type="radio" name="isActive" checked={newSiteData.IsActive === false} onChange={() => setNewSiteData({ ...newSiteData, IsActive: false })} /> No
                                    </label>
                                  </div>
                                </div>
                              </div>

                              <div className="d-flex justify-content-center gap-2 mt-3">
                                <button
                                  className="btn text-white"
                                  style={{ backgroundColor: "#2c9942", padding: "8px 25px", border: 'none', borderRadius: '6px' }}
                                  onClick={handleSaveNewSite}
                                >
                                  Submit
                                </button>
                                <button
                                  className="btn btn-secondary shadow-sm"
                                  style={{ backgroundColor: "#6c757d", padding: "8px 25px", border: 'none', borderRadius: '6px' }}
                                  onClick={() => setShowCreateForm(false)}
                                >
                                  Cancel
                                </button>
                              </div>
                            </div>

                          </>

                        ) : (
                          <>
                            <div className="d-flex justify-content-end gap-2 mb-3 mt-minus30">
                              <button className="btn btn-primary shadow-sm" style={{ backgroundColor: '#2c9942', border: 'none' }} onClick={() => setShowCreateForm(true)}>Add New</button>
                              <button className="btn btn-secondary shadow-sm" onClick={() => handleReturnToMain('')}>Back to Home</button>
                            </div>

                            <div style={{ padding: '15px', clear: 'both', float: 'left', width: '100%', background: '#fff', borderRadius: '10px', border: '1px solid #ccc', marginTop: '15px' }}>
                              <header style={{ padding: '0px 0px 10px 0px' }}>
                                <div className='page-title fw-bold mb-1 font-20'>Manage Location</div>
                              </header>

                              <table className='mtbalenew'>
                                <thead>
                                  <tr>
                                    <th style={{ minWidth: '50px', maxWidth: '50px' }}>S.No.</th>
                                    <th style={{ textAlign: 'center', minWidth: '120px', maxWidth: '120px' }}>Title</th>
                                    <th style={{ textAlign: 'center', minWidth: '250px', maxWidth: '250px' }}>Site URL</th>
                                    {/* <th style={{ maxWidth: '200px', minWidth: '200px', textAlign: 'center' }}>Master Site Share Path</th> */}
                                    <th style={{ minWidth: '70px', maxWidth: '70px', textAlign: 'center' }}>Status</th>
                                    <th style={{ minWidth: '80px', maxWidth: '80px', textAlign: 'center' }}>Action</th>
                                    {/* addhyan 01/04/26: */}
                                    
                                  </tr>
                                </thead>
                                <tbody>
                                  {locationData.map((item: any, index: number) => (
                                    <tr key={item.Id}>
                                      <td style={{ minWidth: '50px', maxWidth: '50px' }}>
                                    <div className="d-flex align-items-center justify-content-center">  <span style={{marginLeft:'0px'}} className="indexdesign">{index + 1}</span></div>  </td>
                                      <td style={{ maxWidth: '120px', minWidth: '120px', textAlign: 'center' }}>{item.Title}</td>
                                      <td title={item.SiteURL} style={{ maxWidth: '250px', minWidth: '250px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{item.SiteURL}</td>
                                      {/* <td title={item.SharewithOtherMeMasterSite} style={{ maxWidth: '200px', minWidth: '200px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{item.SharewithOtherMeMasterSite}</td> */}
                                      <td style={{ minWidth: '70px', maxWidth: '70px', textAlign: 'center' }}>
                                        {item.IsActive ? <span className="badge bg-success">Active</span> : <span className="badge bg-danger">Inactive</span>}
                                      </td>
                                      <td style={{ minWidth: '80px', maxWidth: '80px', textAlign: 'center' }}>
                                        <button className="btn btn-sm btn-outline-primary" onClick={() => handleToggleStatus(item.Id, item.IsActive)}>
                                          {item.IsActive ? "Deactivate" : "Activate"}
                                        </button>
                                      </td>
                                      
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </>
                        )}
                      </div>
                    )}
                    {/* {activeComponent === 'Create Department' && (
                      <div className="position-relative">
                        <button className="btn back-to-admin" onClick={() => handleReturnToMain('')}> Back to Home </button>
                        <Department sp={activeSp} />
                      </div>

                    )} */}
                    {activeComponent === 'Map Division & Department' && (
                      <div className="position-relative">
                        <button className="btn back-to-admin" onClick={() => handleReturnToMain('')}> Back to Home </button>
                        <EntityMapping sp={activeSp} />
                      </div>

                    )}
                  </div>
                )
              }
            </>
            ) : (

              <div>
                {/* <h1>You Don't Have Access</h1> */}
              </div>
            )}
          {/* Aman change 3/4/26 */}
            {activeComponent !== 'Create Location' && toggleManagePermission === "Yes" && toggleManagePermissionCard === "No" ? (
              //  adddhyan - 03/04/2026 start
              <div className="DMSMasterContainer">
                {/* {!selectedGlobalSite && ( */}
                <div className="row manage-master mt-3">
                 {IsSuperAdmin && (
                    
                    <div className="col-sm-3 col-md-3">
                      <a href="">
                        <div className="card-master box1" onClick={(event) => { handleToggleSuper(event, "ManageSuperAdmin") }}>
                          <div className="icon">
                            <img className="" src={managesuperadmin} />
                          </div>
                          <p className="text-dark">Manage Super Admin</p>
                        </div>
                      </a>
                    </div>
                  )}
                 {IsSuperAdmin && (
                    
                    <div className="col-sm-3 col-md-3">
                      <a href="">
                        <div className="card-master box1" onClick={(event) => { handleToggleSuper(event, "ManageGlobalPermission") }}>
                          <div className="icon">
                            <img className="" src={managesuperadmin2} />
                          </div>
                          <p className="text-dark">Manage Team Permission</p>
                        </div>
                      </a>
                    </div>
                  )}
                  {IsSuperAdmin && (
                    
                    <div className="col-sm-3 col-md-3">
                      <a href="">
                        <div className="card-master box1" onClick={(event) => { handleToggleSuper(event, "ManageSuper") }}>
                          <div className="icon">
                            <img className="" src={managesuperadmin1} />
                          </div>
                          <p className="text-dark">Manage Location Admin</p>
                        </div>
                      </a>
                    </div>
                  )}
                  <div className="col-sm-3 col-md-3 ">
                    <a href="">
                      <div className="card-master box1" onClick={(event) => {
                        handleToggleCard(event, "Yes")
                      }
                      }>
                        <div className="icon">
                          <img className="" src={manageUserAndPermissionImage} />
                        </div>
                        <p className="text-dark">Manage Group Permission</p>
                      </div>
                    </a>
                  </div>
                  <div className="col-sm-3 col-md-3 ">
                    <a href="">
                      <div className="card-master box1" onClick={(event) => {
                        handleToggleCard(event, "Manage Folder Deligation")
                      }
                      }>
                        <div className="icon">
                          <img className="" src={managefolderdeligation} />
                        </div>
                        <p className="text-dark">Manage Folder Deligation</p>
                      </div>
                    </a>
                  </div>
                </div>
                 {/* )} */}
              </div>
            ) : (
              <>

              </>
            )
            }
            {toggleManagePermissionCard === "Yes" && (
              <div className="position-relative">
                <div className="mt-minus30">
                  <button className="btn back-to-admin" onClick={() => handleReturnToMainFromPermissionTable('')}>Back To Main</button>

                </div>
                <div style={{

                  position: "relative",

                  marginTop: "70px",
                  padding: "20px",
                  border: "1px solid #ccc",
                  borderRadius: "8px",
                  background: "#fff",
                  clear: "both",
                  float: "left",
                  width: "100%"

                }}>
                  <p className="page-title font-20 text-dark fw-bold mb-2" style={{

                  }}>Manage Group Permission</p>
                  <div className="row">
                    <div className="col-md-4">
  <label>Location</label>
  <div className={`select-wrapper ${validationErrors.location ? 'select-error' : ''}`}>
    <Select
      options={siteCollections}
      value={selectedGlobalSite}
      onChange={(selected: any) => handleGlobalSiteChange(selected, true)}
      onKeyDown={(e: any) => {
        if (e.key === 'Enter') {
          e.preventDefault();
        }
      }}
       //adddhyan - 03/04/2026 start
      placeholder="Select Location..."
      styles={{
        control: (base) => ({
          ...base,
          borderColor: validationErrors.location ? '#dc3545' : base.borderColor,
          borderWidth: validationErrors.location ? '2px' : '1px',
          boxShadow: validationErrors.location ? '0 0 0 0.2rem rgba(220, 53, 69, 0.25)' : base.boxShadow
        })
      }}
    />
  </div>
</div>
                    <div className="col-sm-4">
                      <label>Department</label>
                      <div className={`select-wrapper ${validationErrors.department ? 'select-error' : ''}`}>
                        <Select
                          options={adminPermissionEntity}
                          onChange={(selected: any) =>
                            handleEntitySelect(selected)
                          }
                          onKeyDown={(e: any) => {
                            if (e.key === 'Enter') {
                              e.preventDefault();
                            }
                          }}
                          placeholder="Select Department..."
                          noOptionsMessage={() => "No Department Found..."}
                          styles={{
                            control: (base) => ({
                              ...base,
                              borderColor: validationErrors.department ? '#dc3545' : base.borderColor,
                              borderWidth: validationErrors.department ? '2px' : '1px',
                              boxShadow: validationErrors.department ? '0 0 0 0.2rem rgba(220, 53, 69, 0.25)' : base.boxShadow
                            })
                          }}
                        />
                      </div>
                    </div>
                    <div className="col-sm-4 ">
                      <label>Groups</label>
                      <div className={`select-wrapper ${validationErrors.groups ? 'select-error' : ''}`}>
                        <Select
                          options={groups}
                          onChange={(selected: any) =>
                            handleGroupsSelect(selected)
                          }
                          onKeyDown={(e: any) => {
                            if (e.key === 'Enter') {
                              e.preventDefault();
                            }
                          }}
                          placeholder="Select Groups..."
                          noOptionsMessage={() => "No Groups Found..."}
                          styles={{
                            control: (base) => ({
                              ...base,
                              borderColor: validationErrors.groups ? '#dc3545' : base.borderColor,
                              borderWidth: validationErrors.groups ? '2px' : '1px',
                              boxShadow: validationErrors.groups ? '0 0 0 0.2rem rgba(220, 53, 69, 0.25)' : base.boxShadow
                            })
                          }}
                        />
                      </div>
                    </div>
                    {<div className="col-sm-4 mt-3">
                      <label>Users</label>
                      <div className={`select-wrapper ${validationErrors.users ? 'select-error' : ''}`}>
                        <Select
                          isMulti
                          options={user}
                          onChange={(selected: any) =>
                            handleUsersSelect(selected)
                          }
                          onKeyDown={(e: any) => {
                            if (e.key === 'Enter') {
                              e.preventDefault();
                            }
                          }}
                          placeholder="Select User..."
                          noOptionsMessage={() => "No User Found..."}
                          styles={{
                            control: (base) => ({
                              ...base,
                              borderColor: validationErrors.users ? '#dc3545' : base.borderColor,
                              borderWidth: validationErrors.users ? '2px' : '1px',
                              boxShadow: validationErrors.users ? '0 0 0 0.2rem rgba(220, 53, 69, 0.25)' : base.boxShadow
                            })
                          }}
                        />
                      </div>
                    </div>
                    }

                  </div>
                  <div style={{
                    display: "flex",
                    gap: "10px",
                    justifyContent: "center"

                  }}>
                    <button style={{ padding: '8px 10px', borderRadius: '4px' }} type="button" className="mt-4 btn btn-primary" onClick={handleAddUsers}>
                      Add
                    </button>

                  </div>
                </div>

                {showGroupsTable === "Yes" && (
                  <div>

                    <div style={{ padding: '15px', clear: 'both', float: 'left', marginTop: '15px' }} className={styles.container}>
                      <header style={{ padding: '0px 0px 5px 0px' }}>
                        <div className='page-title fw-bold mb-1 font-20'>{selectedEntityForPermission.value} &gt; {groupDetails?.value && groupDetails?.value.includes('_')
                          ? groupDetails?.value.split('_')[1]
                          : groupDetails?.value || ''} &gt; Details
                        </div>
                      </header>
                      <table className='mtbalenew'>

                        <thead>
                          <tr>
                            <th>Title</th>

                            <th >Description</th>
                          </tr>
                        </thead>
                        <tbody>

                          <tr>
                            <td>
                              {groupDetails?.value && groupDetails?.value.includes('_')
                                ? groupDetails?.value.split('_')[1]
                                : groupDetails?.value || ''}
                            </td>
                            <td>
                              {groupDetails?.Description || ''}
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                    {showGroupsUsers === "Yes" && (<>
                      <div style={{ padding: '15px', clear: 'both', float: 'left', marginTop: '15px' }} className={styles.container}>
                        <header style={{ padding: '0px 0px 5px 0px' }}>
                          <div className='page-title fw-bold mb-1 font-20'>
                            {selectedEntityForPermission.value} &gt;
                            {groupDetails.value && groupDetails.value.includes('_')
                              ? groupDetails.value.split('_')[1]
                              : groupDetails.value || ''}
                            &gt; Users
                          </div>
                        </header>
                        <table className='mtbalenew'>

                          <thead>
                            <tr>
                              <th style={{ minWidth: '55px', maxWidth: '55px' }}>S.No.</th>
                              <th>User</th>
                              <th>User Email</th>
                              <th>Group Name</th>
                              <th>Permission</th>
                              <th>Description</th>
                              <th style={{ minWidth: '65px', maxWidth: '65px' }}>Action</th>
                            </tr>
                          </thead>
                          <tbody>
                            {currentData.map((item: any, index: any) => (
                              <React.Fragment key={item.userId}>
                                <tr>
                                  <td style={{ minWidth: '55px', maxWidth: '55px' }}>
                                    <span className="indexdesign">
                                      {/* {index + 1} */}
                                      {(currentPage - 1) * itemsPerPage + index + 1}
                                    </span>
                                  </td>
                                  <td>
                                    {item.user || ''}
                                  </td>
                                  <td >
                                    {item.email || ''}
                                  </td>
                                  <td>
                                    {item.groupName || ''}
                                  </td>
                                  <td>
                                    {item.permission || ''}
                                  </td>
                                  <td>
                                    {item.Descirption || ''}
                                  </td>
                                  <td style={{ minWidth: '65px', maxWidth: '65px' }}>
                                    <img
                                      className={styles.deleteicon}
                                      src={require("../assets/del.png")}
                                      alt="Delete"
                                      onClick={(event) => {
                                        handleDeleteUser(item.userId, item.groupName)
                                      }}
                                    />
                                  </td>
                                </tr>
                              </React.Fragment>
                            ))}
                          </tbody>
                        </table>
                        <Pagination
                          currentPage={currentPage}
                          totalPages={totalPages}
                          handlePageChange={handlePageChange}

                        />
                      </div>
                    </>)
                    }

                  </div>

                )}

              </div>
            )
            }

            {activeComponent === "ManagePermission" &&
              (
                <div className="position-relative">
                  <button className="btn back-to-admin newbutton_align" onClick={() => handleReturnToMainFromPermissionTable('')}>Back To Main</button>
                  <ManagePermission
                    // selectedGroupUsers={selectedGroupUsers}
                    selectedGropuForPermission={selectedGropuForPermission}
                    selectedEntityForPermission={selectedEntityForPermission}
                    onBack={() => handleBackToManagePermissionCard()}
                  />
                </div>
              )
            }
             {activeComponent === "ManageGlobalPermission" &&
              ( 
                <div className="position-relative">
                  <div className="mt-minus30">
                  <button className="btn back-to-admin mt-minus30" onClick={() => handleReturnToMainFromPermissionTable('')}>Back To Main</button>
              </div>    {/*Aman 27/2/26*/}
                  {/* <ManageSuper sp={activeSp} /> */}
                  <ManageEssaCoreGroups  context={context} />
                </div>
              )
            }
             {activeComponent === "ManageSuperAdmin" &&
              (
                <div className="position-relative">
                  <button className="btn back-to-admin" onClick={() => handleReturnToMainFromPermissionTable('')}>Back To Main</button>
                  {/*Aman 27/2/26*/}
                  {/* <ManageSuper sp={activeSp} /> */}
                  <ManageSuperAdmin  context={context} />
                </div>
              )
            }
            {activeComponent === "ManageSuper" &&
              (
                <div className="position-relative">
                  <button className="btn back-to-admin" onClick={() => handleReturnToMainFromPermissionTable('')}>Back To Main</button>
                  {/*Aman 27/2/26*/}
                  {/* <ManageSuper sp={activeSp} /> */}
                  <ManageSuper sp={activeSp} context={context} />
                </div>
              )
            }
            {toggleManagePermissionCard === "Manage Folder Deligation" &&
              (
                <div className="position-relative">
                  <button className="btn back-to-admin newbutton_align" onClick={() => handleReturnToMainFromPermissionTable('')}>Back To Main</button>
                  {/* <ManageFolderDeligation
                    sp={activeSp}
                    // selectedGroupUsers={selectedGroupUsers}
                    //  selectedGropuForPermission={selectedGropuForPermission}
                    //  selectedEntityForPermission={selectedEntityForPermission}
                    onBack={() => handleBackToManagePermissionCard()}
                  />  Ritik 03/04/2026 commented */}
                {/* Ritik 03/04/2026 added Props  */}
                <ManageFolderDeligation 
  sp={activeSp}
  siteCollections={siteCollections}
  context={context}
  onBack={()=>handleBackToManagePermissionCard()}
                />
                {/* end here  */}
                </div>
              )
            }
          </div>
        </div>
      </div>
    </div>


  );
};



const DMSAdmin: React.FC<IDmsMusaibProps> = (props) => {
  const { context, someOtherProp } = props;
  return (
    <Provider>
      <Dmsadmincomponent context={context} someOtherProp={props} />
    </Provider>
  );
};

export default DMSAdmin;
