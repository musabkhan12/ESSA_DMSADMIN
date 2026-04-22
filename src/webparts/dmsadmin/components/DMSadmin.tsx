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
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "bootstrap/dist/css/bootstrap.min.css";
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
import { spfi,SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import { PermissionKind } from "@pnp/sp/security";
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import "../../verticalSideBar/components/VerticalSidebar.scss";
import "./dmscss";
import "./DMSAdmincss"
import { useState , useRef , useEffect} from "react";
import {IDmsMusaibProps} from './IDmsMusaibProps'
import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
 import EntityMapping from "./EntityMapping";
import Devision from "./Division";
import Department from "./Department";
// import CreateEntity from "./CreateEntity";
import CreateEntity from "./EntityDemo";
import Select from "react-select";
import Swal from 'sweetalert2';
import { ManagePermission } from "./Managepermission";
import { ManageSuper } from "./ManageSuper";

let selectedEntityForPermission:any;
let selectedGropuForPermission:any;
// let selectedUsersForPermission:any;
let selectedGroupUsers:any[];
let superA=false;
let superAdminArray:any[];
// let IsAdmin=false;

const Dmsadmin = ({ props }: any) => {
  const sp: SPFI = getSP();

  const { useHide }: any = React.useContext(UserContext);
  const elementRef = React.useRef<HTMLDivElement>(null);

  // New Code
  const [toggleManagePermission,setToggleManagePermission]=useState('No');
  const [adminPermissionEntity,setAdminPermissionEntity]=useState<any[]>([]);
  // const [user,setUser]=useState<any[]>([]);
  const [groups,setGroups]=useState<any[]>([]);
  const [IsSuperAdmin,setIsSuperAdmin]=useState(false);

  const [toggleManagePermissionCard,setToggleManagePermissionCard]=useState("No");

  const handleToggleCard=(event:any,name:any)=>{
    event.preventDefault();
    setToggleManagePermissionCard(name);
    setToggleManagePermission("No");
    setActiveComponent(null);
    
  }

  const handleToggleSuper=(event:any,name:any)=>{
    event.preventDefault();
    setToggleManagePermission("No");
    setActiveComponent(name);
  }
  // const [permissionTable,setPermissionTable]=useState<string >('')
  const currentUserEmailRef = useRef('');

  const getcurrentuseremail = async()=>{
    const userdata = await sp.web.currentUser();
    currentUserEmailRef.current = userdata.Email;
    getDetailsOfSuperAdmin();
    getDetailsOfAdmin();  
    // fetchUsers()
  }

  const getDetailsOfSuperAdmin=async()=>{
    try {
        const usersFromDMSSuperAdmin = await sp.web.siteGroups.getByName('DMSSuper_Admin').users();
        superAdminArray=usersFromDMSSuperAdmin;
        usersFromDMSSuperAdmin.forEach((user)=>{
            if(user.Email === currentUserEmailRef.current){
              superA=true;
              setIsSuperAdmin(true);
              // setToggleManagePermission('Yes');
            }
        })
        console.log("usersFromDMSSuperAdmin",usersFromDMSSuperAdmin);
    } catch (error) {
      console.log("error in getting the details of super admin",error);
    }
  }
  console.log("IsSuperAdmin",IsSuperAdmin);
  const getDetailsOfAdmin=async()=>{
    try {
        const entityDetails=await sp.web.lists.getByTitle("MasterSiteURL").items.select("SiteURL","Title","Active","SiteID").filter(`Active eq 'Yes'`)();
        console.log("entityDetails",entityDetails);
        let entityArray:any[]=[]
        const subsiteAdminDetails = await Promise.all(
          entityDetails.map(async (entity) => {
            try {
              const subsiteContext = await sp.site.openWebById(entity.SiteID);
              const usersFromAdmin = await subsiteContext.web.siteGroups.getByName(`${entity.Title}_Admin`).users();
              console.log("IsSuperAdmin from entityDeatils forEach",IsSuperAdmin,superA,usersFromAdmin);
              if(superA || usersFromAdmin.length !== 0){
                  if(usersFromAdmin.length !== 0 && superA === false){
                    // usersFromAdmin.forEach((user)=>{
                    for(const user of usersFromAdmin){
                      if(user.Email === currentUserEmailRef.current){
                          console.log("current user is Admin",currentUserEmailRef.current);
                          console.log("entity",entity.Title);
                          console.log("all users in the admin group",usersFromAdmin);
                          entityArray.push({
                            value:entity.Title,
                            label:entity.Title,
                            SiteID:entity.SiteID
                          });
                          return usersFromAdmin;
                      }
                    }
                  }
                  if(superA){
                    // console.log("current user is super admin",IsSuperAdmin);
                    // console.log("user is super admin",currentUserEmailRef.current);
                    // console.log("entity",entity.Title);
                    // console.log("users from the admin group",usersFromAdmin);
                    entityArray.push({
                      value:entity.Title,
                      label:entity.Title,
                      SiteID:entity.SiteID
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
        console.log("subsiteAdminDetails",subsiteAdminDetails);
        console.log("entityArray",entityArray);
        let finalUserArray:any[]=[];
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
        if(!allUndefined){
          // IsAdmin=true;
          setToggleManagePermission("Yes");
        }
        
        setAdminPermissionEntity(entityArray);
        console.log("allUndefined",allUndefined);
        console.log("finalUserArray",finalUserArray);
    }catch (error) {
      console.log("error getting entity details",error);
    }
  }



  // const fetchUsers=async()=>{
  //   const user = await sp.web.siteUsers();
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

  // handle entity Select
  const handleEntitySelect=async(selectedEntity:any)=>{
      console.log("selectedEntity",selectedEntity);
      selectedEntityForPermission=selectedEntity;
      const subsiteContext = await sp.site.openWebById(selectedEntity.SiteID);
      if(IsSuperAdmin){
        try {
          const usersFromAdmin = await subsiteContext.web.siteGroups.getByName(`${selectedEntity.value}_Admin`).users();
          console.log("usersFromAdmin -> IsSuperAdmin",usersFromAdmin);
          // console.log("IsAdmin",IsAdmin);
          // Check if current user email exists in the usersFromAdmin list
          const emailExists = usersFromAdmin.some(user => user.Email.toLowerCase() === currentUserEmailRef.current.toLowerCase());
          console.log("emailExists",emailExists);
          if(!emailExists){
            console.log(`${currentUserEmailRef.current} does not exist in the list. Adding to the admin group.`);
            const userObj = await sp.web.ensureUser(currentUserEmailRef.current);
            console.log("userObj",userObj);
            const users=await subsiteContext.web.siteGroups.getByName(`${selectedEntity.value}_Admin`).users.add(userObj.data.LoginName);
            console.log(`User Added Succecssfully in the ${selectedEntity.value}_Admin`,users);
          }else{
            console.log(`${currentUserEmailRef.current} already exists in the list.`);
          }

        } catch (error) {
          console.log(`Error in Adding super admin to the ${selectedEntity.value}_Admin gropup`,error)
        }
      }
      
      // Fetch all the groups in the subsite
      interface IMember {
        PrincipalType: number;
        Title:String;
        Id:number 
      }
      interface IRoleAssignmentInfo {
        Member?: IMember; 
      }
      const groups3:IRoleAssignmentInfo[] = await subsiteContext.web.roleAssignments.expand("Member")();
      console.log("groups3",groups3);
      const filteredMembers=groups3.filter(roleAssignment => {
        return roleAssignment.Member.PrincipalType === 8;
      });

      const filteredGroups = filteredMembers.map((object) => ({
          value: object.Member.Title,
          label: object.Member.Title,
          Id: object.Member.Id,
      }));
      console.log("filteredGroups",filteredGroups);
      console.log("filteredMembers",filteredMembers);
      // filter the DMSSuper_Admin
      const filteredRoles = filteredGroups.filter(role => role.value !== "DMSSuper_Admin");
      setGroups(filteredRoles);
  }

  //  handle groups select
  const handleGroupsSelect=async(selectedGrous:any)=>{
      console.log("selectedGrous",selectedGrous);
      console.log("selectedEntityForPermission",selectedEntityForPermission);
      selectedGropuForPermission=selectedGrous;
      try {
        const subsiteContext = await sp.site.openWebById(selectedEntityForPermission.SiteID);
        const usersFromSelectedGroups = await subsiteContext.web.siteGroups.getByName(`${selectedGrous.value}`).users();
        console.log("usersFromSelectedGroups",usersFromSelectedGroups);
        selectedGroupUsers=usersFromSelectedGroups;
      } catch (error) {
        console.log("error from getting the users from the groups after selecting the groups",error);
      }
      

  }

  // const handleUsersSelect=(selectedUser:any)=>{
  //       console.log("selectedUser",selectedUser);
  //       selectedUsersForPermission=selectedUser;
  // }

  // const handleAddUsers=async()=>{
  //     console.log("selectedUsersForPermission",selectedUsersForPermission);
  //     console.log("selectedGropuForPermission",selectedGropuForPermission);
  //     console.log("selectedEntityForPermission",selectedEntityForPermission);

  //     if(selectedUsersForPermission === undefined || selectedUsersForPermission.length === 0){
  //       checkValidation();
  //       return;
  //     }
  //     if(selectedGropuForPermission === undefined){
  //       checkValidation();
  //       return;
  //     }
  //     if(selectedEntityForPermission === undefined){
  //       checkValidation();
  //       return;
  //     }

  //     const subsiteContext = await sp.site.openWebById(selectedEntityForPermission.SiteID); 
  //     selectedUsersForPermission.forEach(async(user:any)=>{
  //       try {
  //         const userObj = await sp.web.ensureUser(user.email);
  //         console.log("userObj",userObj);
  //         const users=await subsiteContext.web.siteGroups.getByName(`${selectedGropuForPermission.value}`).users.add(userObj.data.LoginName);
  //         console.log(`${user.email} added to the group successfully.`,users);
  //       } catch (error) {
  //         console.error(`Failed to add ${user.email} to the group: `, error);
  //       }
  //     })
      
  //   }

  const hanldeManagePermission=()=>{
    if(selectedGroupUsers === undefined && selectedGropuForPermission === undefined){
      checkValidation();
    }else{
      setActiveComponent("ManagePermission");
      setToggleManagePermission('No');
      setToggleManagePermissionCard("No");
    }   
  }

  const handleReturnToMainFromPermissionTable=(Name:any)=>{
    setToggleManagePermission('Yes');
    setActiveComponent(Name);
    setToggleManagePermissionCard("No");
  }

  const checkValidation=()=>{
        Swal.fire("Please fill out the fields!", "All fields are required");
  }

  const handleCreate=(Name:any)=>{
    setActiveComponent(Name);
    setToggleManagePermission('No');
  }

  const handleBackToManagePermissionCard=()=>{
    setToggleManagePermissionCard('Yes');
    setActiveComponent(null);
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
const [activeComponent, setActiveComponent] = useState<string >('');
////
console.log(activeComponent , "activeComponent")
const handleReturnToMain = (Name:any) => {
  setActiveComponent(Name); // Reset to show the main component
  console.log(activeComponent , "activeComponent updated")
  setToggleManagePermission('Yes');
};
  const getmasterlis = async () => {
    try {
      const items = await sp.web.lists.getByTitle('DMSAdmin').items();
      console.log(items, "getmasterlis");
      setMylistdata(items);
      
    } catch (error) {
      console.error("Error fetching list items:", error);
    }
  };
  console.log(Mylistdata , "Mylistdata")
  useEffect(() => {
    getcurrentuseremail();
    getmasterlis();
  }, []);

  return (
    <div id="wrapper" ref={elementRef}>
    <div
      className="app-menu"
      id="myHeader">
      <VerticalSideBar _context={sp} />
    </div>
    <div className="content-page">
      <HorizontalNavbar/>
      <div className="content" style={{marginLeft: `${!useHide ? '240px' : '80px'}`,marginTop:'1.5rem'}}>
       
      <div className="container-fluid  paddb">
    {IsSuperAdmin ? (<>
      {activeComponent === "" ?
               (<div>
                    <div className="DMSMasterContainer">
                <h4 className="page-title fw-bold mb-1 font-20">Settings</h4>
                <div className="Route">
                  {" "}
                  <h2 className="Home">Home</h2>
                  <span className="greater">&gt;</span>{" "}
                  <h2 className="Setting">Settings</h2>{" "}
                </div>
                <div className="row manage-master mt-3">
                  {Mylistdata.map((item: any) => {
                    const imageData = JSON.parse(item.Image); // Assuming 'ImageColumn' is the column name
                    const itemid = String(item.Id);
                    console.log(itemid, "itemsid");
                    console.log(imageData, "imagedata");
                    const imageUrl = `https://officeindia.sharepoint.com//_api/v2.1/sites('338f2337-8cbb-4cd1-bed1-593e9336cd0e,e2837b3f-b207-41eb-940b-71c74da3d214')/lists('3f31e4eb-27b3-4370-b5cd-8cf594981912')/items('${itemid}')/attachments('${imageData.fileName}')/thumbnails/0/c3000x2000/content?prefer=noredirect,closestavailablesize`;
                    console.log(imageUrl, "imageurl");
                    return (
                      <div className="col-sm-3 col-md-3 mt-2">
                        <a href={item?.LinkUrl}>
                          <div className="card-master box1" onClick={()=>handleCreate(item.Name)}>
                            <div className="icon">
                              <img className="CardImage" src={imageUrl} />
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
                <div>
                  {/* {activeComponent === 'Create Entity' && (
                    <div>
                      <button onClick={()=>handleReturnToMain('')}> Back to Home </button>
                      <CreateEntity/>
                    </div>
               
                  )}  */}
                  {activeComponent === 'Create Devision' && (
                    <div>
                      <button onClick={()=>handleReturnToMain('')}> Back </button>
                      {/*Aman 27/2/26*/}
                      <Devision sp={sp}/>
                    </div>
          
                  )} 
                  {activeComponent === 'Create Department' && (
                    <div>
                      <button onClick={()=>handleReturnToMain('')}> Back to Home </button>
                      <Department sp={sp} />
                    </div>
                   
                  )} 
                  {activeComponent === 'Map Devision & Department' && (
                    <div>
                      <button onClick={()=>handleReturnToMain('')}> Back to Home </button>
                      <EntityMapping sp={sp}/>
                    </div>
          
                  )} 
                </div>
               )
      }
      </>
    ) :( 

      <div>
        <h1>You Don't Have Access</h1>
      </div>
    ) }        
    {toggleManagePermission === "Yes" ? (
                  <div className="DMSMasterContainer">
                  <div className="row manage-master mt-3">
                    {IsSuperAdmin && (
                      <div className="col-sm-3 col-md-3 mt-2">
                        <a href="">
                              <div className="card-master box1" onClick={(event)=>{handleToggleSuper(event,"ManageSuper")}}>
                                <div className="icon">
                                  <img className="CardImage" />
                                </div>
                                <p className="text-dark">Manage Super Admin</p>
                              </div>
                        </a>
                    </div>
                    )}
                    <div className="col-sm-3 col-md-3 mt-2">
                          <a href="">
                                  <div className="card-master box1" onClick={(event)=>
                                    {
                                      handleToggleCard(event,"Yes")
                                    }
                                    }>
                                    <div className="icon">
                                      <img className="CardImage" />
                                    </div>
                                    <p className="text-dark">Manage users and permission</p>
                                  </div>
                          </a>
                      </div>
                  </div>
                  </div>
               ) :(
                <>
                  {/* {IsSuperAdmin === false && (
                    <div style={
                    {
                      marginLeft:"30px",
                      marginTop:"30px",
                      color:"#707070"
                    }
                   }>
                     <h6>'Unauthorized access. You do not have permission to view this page.'</h6> 
                </div>
                  )} */}
                </>
               )
    }
    {toggleManagePermissionCard === "Yes" && (
      <div>
        {/* Ritik 22/04/26 */}
          <button className="back-to-admin" onClick={()=>handleReturnToMainFromPermissionTable('')}>Back To Home</button> 
          <div style={{
                    width:"fit-content",
                    position:"relative",
                    marginLeft:"50px",
                    marginTop:"50px",
                    padding:"20px",
                    border:"2px solid #7fc4de",
                    borderRadius:"10px",
                    background:"#fff",

                  }}>
                  <p style={{
                    marginBottom:"20px",
                    marginLeft:"300px"
                  }}>Manage Users And Permission</p>
                  <div style={{
                    gap:"60px",
                    display:"flex"
                  }}>
                    <div style={{
                      width:"220px"
                    }}>
                      <label>Entity</label>
                      <Select                        
                          options={adminPermissionEntity}
                          onChange={(selected: any) =>
                            handleEntitySelect(selected)
                          }
                          placeholder="Select Entity..."
                          noOptionsMessage={() => "No Entity Found..."}
                        />
                    </div>
                    <div  style={{
                      width:"220px"
                    }}>
                      <label>Groups</label>
                      <Select
                          options={groups}
                          onChange={(selected: any) =>
                            handleGroupsSelect(selected)
                          }
                          placeholder="Select Groups..."
                          noOptionsMessage={() => "No Groups Found..."}
                        />
                    </div>
                     {/* <div  style={{
                      width:"220px"
                      }}>
                     <label>Users</label>
                      <Select
                          isMulti
                          options={user}
                          onChange={(selected: any) =>
                            handleUsersSelect(selected)
                          }
                          placeholder="Select User..."
                          noOptionsMessage={() => "No User Found..."}
                        />
                     </div>     */
                     }
                  </div>
                  <div style={{
                    display:"flex",
                    gap:"10px",
                    marginLeft:"300px",
                    marginTop:"30px"
                  }}>
                    <button type="button" onClick={hanldeManagePermission}>
                       Manage Permission
                    </button>
                  </div>
                </div>
                {/* <div>
                  {activeComponent === "ManagePermission" && 
                      (
                        <div>
                            <button onClick={()=>handleReturnToMainFromPermissionTable('')}>BackToMain</button>
                            <ManagePermission
                              // selectedGroupUsers={selectedGroupUsers}
                              selectedGropuForPermission={selectedGropuForPermission}
                              selectedEntityForPermission={selectedEntityForPermission}
                            />
                        </div>
                      )
                  }
                </div> */}
               
              
    </div>)
    }
    {activeComponent === "ManagePermission" && 
                      (
                        <div className="position-relative">
                          {/* Ritik 22/04/26 */}
                            <button className="back-to-admin" onClick={()=>handleReturnToMainFromPermissionTable('')}>Back To Home</button>
                            <ManagePermission
                              // selectedGroupUsers={selectedGroupUsers}
                              selectedGropuForPermission={selectedGropuForPermission}
                              selectedEntityForPermission={selectedEntityForPermission}
                              onBack={()=>handleBackToManagePermissionCard()}
                            />
                        </div>
                      )
    }
    {activeComponent === "ManageSuper" && 
                      (
                        <div className="position-relative">
                          {/* Ritik 22/04/26 */}
                            <button className="back-to-admin" onClick={()=>handleReturnToMainFromPermissionTable('')}>Back To Home</button>
                            {/*Aman 27/2/26*/}
                            {/* <ManageSuper sp={sp}/> */}
                            <ManageSuper sp={sp} context={props.context}/>
                        </div>
                      )
    }
    {/* Old Manage Permission  */}
    {/* {toggleManagePermission === "Yes" ? (
                <div style={{
                    width:"fit-content",
                    position:"relative",
                    marginLeft:"50px",
                    marginTop:"50px",
                    padding:"20px",
                    border:"2px solid #7fc4de",
                    borderRadius:"10px",
                    background:"#fff",

                  }}>
                  <p style={{
                    marginBottom:"20px",
                    marginLeft:"300px"
                  }}>Manage Users And Permission</p>
                  <div style={{
                    gap:"60px",
                    display:"flex"
                  }}>
                    <div style={{
                      width:"220px"
                    }}>
                      <label>Entity</label>
                      <Select                        
                          options={adminPermissionEntity}
                          onChange={(selected: any) =>
                            handleEntitySelect(selected)
                          }
                          placeholder="Select Entity..."
                          noOptionsMessage={() => "No Entity Found..."}
                        />
                    </div>
                    <div  style={{
                      width:"220px"
                    }}>
                      <label>Groups</label>
                      <Select
                          options={groups}
                          onChange={(selected: any) =>
                            handleGroupsSelect(selected)
                          }
                          placeholder="Select Groups..."
                          noOptionsMessage={() => "No Groups Found..."}
                        />
                    </div>
                  </div>
                  <div style={{
                    display:"flex",
                    gap:"10px",
                    marginLeft:"300px",
                    marginTop:"30px"
                  }}>
                    <button type="button" onClick={hanldeManagePermission}>
                       Manage Permission
                    </button>
                  </div>
                </div>
               ) : (
                <div>
                  {activeComponent === "ManagePermission" && 
                      (
                        <div>
                            <button onClick={()=>handleReturnToMainFromPermissionTable('')}>BackToMain</button>
                            <ManagePermission
                              selectedGropuForPermission={selectedGropuForPermission}
                              selectedEntityForPermission={selectedEntityForPermission}
                            />
                        </div>
                      )
                  }
                </div>
               )
    } */}
               
              </div>
            </div>
          </div>
          </div>
   
        
  );
};



const DMSAdmin: React.FC<IDmsMusaibProps> = (props) =>{
  return (
    <Provider>
      <Dmsadmin  props={props}/>
    </Provider>
  );
};

export default DMSAdmin;
