import React, { useEffect, useRef, useState } from "react";
import Select from "react-select";
import { getSP , getGraphClient } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import Swal from "sweetalert2";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faSort } from '@fortawesome/free-solid-svg-icons';//Ritik 22/04/2026
import styles from './Form.module.scss'
let superA=false;
let usersFromGroups:any[]=[];
let selectedEntityForPermission:any;
let groupDetails:any;
let selectedGroupUsers:any
let selectedUsersForPermission:any;
let seleccteduserforapproval:any
let selectedGropuForPermission:any;
let superAdminArray:any[];
// Ritik 03/04/2026 - replaced for the siteCollections prop
interface ManageFolderDeligationProps {
  sp: SPFI;
  onBack: ()=>void;
  siteCollections: { value: string; label: string; siteUrl: string; id: any }[];
  context: any;
}
const ManageFolderDeligation: React.FC<ManageFolderDeligationProps> = ({ sp, onBack, siteCollections, context }) => {
  // end here 
    // const sp: SPFI = getSP();
    const [adminPermissionEntity,setAdminPermissionEntity]=useState<any[]>([]);
    const [IsSuperAdmin,setIsSuperAdmin]=useState(false);
    const [user,setUser]=useState<any[]>([]);
    const [allUsersFromGroups,setAllUsersFromGroups]=useState<any[]>([]);
    const [allUsersFromADMINGroups,setAllUsersFromADMINGroups]=useState<any[]>([]);
    const [toggleManagePermission,setToggleManagePermission]=useState('Yes');
    const [showGroupsTable,setShowGroupsTable]=useState("No");
    const [showGroupsUsers,setShowGroupsUsers]=useState("No");
    const [groups,setGroups]=useState<any[]>([]);
  const currentUserEmailRef = useRef('');
  // Ritik 03/04/2026 - added state for selected site collection
  const [selectedSiteCollection, setSelectedSiteCollection] = React.useState<any>(null);
  const [activeSp, setActiveSp] = React.useState<SPFI>(sp);
  const [selectedEntity, setSelectedEntity] = React.useState<any>(null);
const [selectedUser, setSelectedUser] = React.useState<any>(null);
const [selectedApprovers, setSelectedApprovers] = React.useState<any[]>([]);

// Aman 21/4/26 start
 const [formErrors, setFormErrors] = useState({
  location: false,
  department: false,
  users: false,
  approvers: false
});
// Aman 21/4/26 end
  // end here
  // ritik 03/04/2026 - handle site collection change and set the sp context accordingly replaced with the below code
    useEffect(() => {
  if (selectedSiteCollection) {
    getcurrentuseremail();
  }
}, [activeSp]);
    const getcurrentuseremail = async()=>{
    const userdata = await activeSp.web.currentUser(); //  Ritik 03/04/2026 - ActiveSp Get the current user's email
        currentUserEmailRef.current = userdata.Email;
 
        getDetailsOfAdmin();  
        getDetailsOfSuperAdmin()
        fetchUsers()
      }
      const getDetailsOfSuperAdmin=async()=>{
        try {
            const usersFromDMSSuperAdmin = await activeSp.web.siteGroups.getByName('DMSSuper_Admin').users(); // ritik 03/04/2026 - ActiveSp Get users from DMSSuper_Admin group
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
      const fetchUsers=async()=>{
        const user = await activeSp.web.siteUsers(); // ritik 03/04/2026 - ActiveSp Fetch all site users to show in the dropdown when super admin want to add user in the folder delegation group without adding them in the admin group
        console.log("users fetch from the site",user);
          const usersArray=user.map((user)=>(
                {
                  id:String(user.Id),
                  value: user.Title,
                  email: user.Email,
                  label:user.Title,
                  loginName:user.LoginName
                }
          ))
          console.log("site users",usersArray);
          setUser(usersArray);
      }
    const getDetailsOfAdmin=async()=>{
        try {
            const entityDetails=await activeSp.web.lists.getByTitle("MasterSiteURL").items.select("SiteURL","Title","Active","SiteID").filter(`Active eq 'Yes'`)(); // Ritik 03/04/2026 - ActiveSp Get the list of active entities from MasterSiteURL list
            console.log("entityDetails",entityDetails);
            let entityArray:any[]=[]
            const subsiteAdminDetails = await Promise.all(
              entityDetails.map(async (entity:any) => {
                try {
                  const subsiteContext = await activeSp.site.openWebById(entity.SiteID);
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
  // ritik 03/04/2026 - handle site collection change and set the sp context accordingly replaced with the below code
  const handleSiteChange = async (selected: any) => {
  setSelectedSiteCollection(selected);
  setAdminPermissionEntity([]);
  setGroups([]);
  setShowGroupsTable("No");
  setShowGroupsUsers("No");
  setSelectedEntity(null);
  setSelectedUser(null);
  setSelectedApprovers([]);
  selectedEntityForPermission = undefined;
  selectedGropuForPermission = undefined;
  groupDetails = undefined;
    
  if (selected && selected.siteUrl) {
    const { spfi, SPFx } = await import("@pnp/sp");
    const targetSp = spfi(selected.siteUrl).using(SPFx(context));
    setActiveSp(targetSp);
  } else {
    setActiveSp(sp);
  }
};

      const handleEntitySelect=async(selectedEntity:any)=>{
        console.log("selectedEntity",selectedEntity);
        selectedEntityForPermission=selectedEntity;
        const subsiteContext = await activeSp.site.openWebById(selectedEntity.SiteID);
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
              const userObj = await activeSp.web.ensureUser(currentUserEmailRef.current);
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
        // const onlyFolderDeligationuser = groups3.find( )
        const filteredMembers=groups3.filter(roleAssignment => {
          return roleAssignment.Member.PrincipalType === 8;
        });

        const filteredGroups = filteredMembers.map((object:any) => (
            
            {
       
            value: object.Member.Title,
            label: object.Member.Title,
            Id: object.Member.Id,
        }));
        console.log("filteredGroups",filteredGroups);
        console.log("filteredMembers",filteredMembers);
        // filter the DMSSuper_Admin
        const filteredRoles = filteredGroups.filter(role => role.value !== "DMSSuper_Admin");
        console.log("filteredRoles before permission",filteredRoles);
        // let usersFromGroups:any[]=[];
        usersFromGroups=[];
    

        await Promise.all(filteredRoles.map(async (group) => {
            // alert(group)
          const result = group.value.split("_")[1];
          let permission = "";
          let description = "";
          
          // Determine permission and description based on result
          switch (result) {
            case "Admin":
              permission = "Admin";
              description = "Full Control - Has full control.";
              break;
            case "Contribute":
              permission = "Contribute";
              description = "Can view, add, update, and delete documents.";
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
              description = "Can view, add, update, and delete documents.";
              break;
            case "Approval":
              permission = "Approval";
              description = "Can view, add, update, and delete documents.";
              break;
            case "AllUsers":
              permission = "AllUsers";
              description = "Can view, add, update, and delete documents.";
              break;
            case "FolderDeligation ":
              permission = "FolderDeligation ";
              description = "Can create Folder on behlaf of Admin";
              break;
            default:
              permission = "Unknown";
              description = "Unknown role";
              break;
          }
          
          (group as any).permission=permission;
          (group as any).Description=description;

          // Fetch users based on group name
          const users = await subsiteContext.web.siteGroups.getByName(`${group.value}`).users();
      
          // Add each user to the usersFromGroups array
          users.forEach((user) => {
            usersFromGroups.push({
              user: user.Title,
              email: user.Email,
              groupName: group.value,
              permission:permission,
              Descirption: description,
              userId: user.Id,
            });
          });
        }));
      
        console.log("allUsersFromGroups",usersFromGroups);
        console.log("filteredRoles",filteredRoles);
        // New code start filter the group and its description
        groupDetails=filteredRoles.find(item => item.value === `${selectedEntity.value}_FolderDeligation`);
        let getapprovaluserfromadmingroup = filteredRoles.find(item => item.value === `${selectedEntity.value}_Admin`);
        console.log("groupDetails",groupDetails);
        selectedGropuForPermission=groupDetails;
        try {
          const subsiteContext = await activeSp.site.openWebById(selectedEntityForPermission.SiteID);
          const usersFromSelectedGroups = await subsiteContext.web.siteGroups.getByName(`${groupDetails.value}`).users();
          const usersFromapprovalGroups = await subsiteContext.web.siteGroups.getByName(`${getapprovaluserfromadmingroup.value}`).users();
          console.log("usersFromSelectedGroups",usersFromSelectedGroups);
          console.log("usersFromapprovalGroups",usersFromapprovalGroups);
          selectedGroupUsers=usersFromSelectedGroups;
          
            const showUsersFromGroupsOnTable=usersFromSelectedGroups.map((user)=>{
              return {
                user: user.Title,
                email: user.Email,
                groupName: groupDetails.value,
                permission:groupDetails.permission,
                Descirption: groupDetails.Description,
                userId: user.Id,
              }
            })
            const setusersFromapprovalGroups =usersFromapprovalGroups.map((user)=>{
              return {
                 id:user.Id,
                  value: user.Title,
                  email: user.Email,
                  label:user.Title,
                  loginName:user.LoginName
              }
            })


          console.log("showUsersFromGroupsOnTable1",showUsersFromGroupsOnTable);
          setAllUsersFromGroups([]);
          setAllUsersFromGroups(showUsersFromGroupsOnTable);
          setShowGroupsUsers("Yes");

          setAllUsersFromADMINGroups(setusersFromapprovalGroups)


        } catch (error) {
          console.log("error from getting the users from the groups after selecting the groups",error);
        }
        
        // end
        setGroups(filteredRoles);
        // setAllUsersFromGroups([]);
        // setAllUsersFromGroups(usersFromGroups);
        setShowGroupsTable("Yes");
    }

    const handleGroupsSelect=async(selectedGrous:any)=>{
        // Set selected groups start
        groupDetails=selectedGrous;
        // End
          console.log("selectedGrous",selectedGrous);
          console.log("selectedEntityForPermission",selectedEntityForPermission);
          selectedGropuForPermission=selectedGrous;
          try {
            const subsiteContext = await activeSp.site.openWebById(selectedEntityForPermission.SiteID);
            const usersFromSelectedGroups = await subsiteContext.web.siteGroups.getByName(`${selectedGrous.value}`).users();
            console.log("usersFromSelectedGroups",usersFromSelectedGroups);
            selectedGroupUsers=usersFromSelectedGroups;
  
              const showUsersFromGroupsOnTable=usersFromSelectedGroups.map((user)=>{
                return {
                  user: user.Title,
                  email: user.Email,
                  groupName: selectedGrous.value,
                  permission:selectedGrous.permission,
                  Descirption: selectedGrous.Description,
                  userId: user.Id,
                }
              })
            console.log("showUsersFromGroupsOnTable",showUsersFromGroupsOnTable);
            setAllUsersFromGroups([]);
            setAllUsersFromGroups(showUsersFromGroupsOnTable);
            setShowGroupsUsers("Yes");
          } catch (error) {
            console.log("error from getting the users from the groups after selecting the groups",error);
          }
          
  
      }
      const handleUsersSelect=(selectedUser:any)=>{
        console.log("selectedUser",selectedUser);
        selectedUsersForPermission=selectedUser;
  }
      const handleUsersforapprovalSelect=(selectedUser:any)=>{
        console.log("selectedUser",selectedUser);
        seleccteduserforapproval=selectedUser;
  }
  const onSuccess=(groupName:any)=>{
    Swal.fire({
      // Rohit 21/4/26 start 
      title: "Added Successsfully.",
      //text: `User Added Suucessfuly to the ${groupName}.`,  
      icon: "success"
    });
    // Rohit 21/4/26 end 
  }
  const checkValidation=()=>{
    Swal.fire("Please fill out the fields!", "All fields are required");
}
  const handleAddUsers = async () => {
    console.log("selectedUsersForPermission", selectedUsersForPermission);
    console.log("selectedGropuForPermission", selectedGropuForPermission);
    console.log("selectedEntityForPermission", selectedEntityForPermission);
   // Aman 21/4/26 start
    // if (
    //   selectedUsersForPermission === undefined ||
    //   selectedUsersForPermission.length === 0
    // ) {
    //   checkValidation();
    //   return;
    // }
    let errors = {
  location: !selectedSiteCollection,
  department: !selectedEntityForPermission,
  users: !selectedUsersForPermission,
  approvers: !seleccteduserforapproval || seleccteduserforapproval.length === 0
};

setFormErrors(errors);

if (errors.location || errors.department || errors.users || errors.approvers) {
  checkValidation();
  return;
}
// Aman 21/4/26 end
    if (
      allUsersFromADMINGroups === undefined ||
      allUsersFromADMINGroups.length === 0
    ) {
      checkValidation();
      return;
    }
    if (selectedGropuForPermission === undefined) {
      checkValidation();
      return;
    }
    if (selectedEntityForPermission === undefined) {
      checkValidation();
      return;
    }

    const subsiteContext = await activeSp.site.openWebById(
      selectedEntityForPermission.SiteID
    );
    //wait for all add operations to complete
    // const addUsersPromises = selectedUsersForPermission.map(
    //   async (user: any) => {
    //     try {
    //       const userObj = await sp.web.ensureUser(user.email);
    //       console.log("userObj", userObj);
    //       const users = await subsiteContext.web.siteGroups
    //         .getByName(`${selectedGropuForPermission.value}`)
    //         .users.add(userObj.data.LoginName);
    //       console.log(`${user.email} added to the group successfully.`, users);
           
    //       const addedItem = await sp.web.lists.getByTitle("DMSFolderDeligationApprovalMaster").items.add({
    //         Title : selectedEntityForPermission.value,
    //         CurrentUser	: selectedUsersForPermission[0].email ,
    //         Approvals: {
    //           results: seleccteduserforapproval.map((user:any) => ({
    //             Key: user.loginName
    //           }))
    //         }
    //         // Approvals : 
    //       });


    //     } catch (error) {
    //       console.error(`Failed to add ${user.email} to the group: `, error);
    //     }
    //   }
    // );
// If selectedUsersForPermission is a single object
const user = selectedUsersForPermission;

const getItem = await activeSp.web.lists.getByTitle("DMSFolderDeligationApprovalMaster").items.select("*").filter(`CurrentUser eq '${user.email}' and SiteTitle eq '${selectedEntityForPermission.value}'`)(); // Ritik 03/04/2026 - ActiveSp Check if the user is already added in the delegation list for the same entity

console.log("getItem",getItem);

if(getItem.length > 0){
  Swal.fire("User Already Exists", "This user is already a member of the Delegation group.", "warning");
  return
}

try {
  const userObj = await activeSp.web.ensureUser(user.email);
  console.log("userObj", userObj);
  
  const users = await subsiteContext.web.siteGroups
    .getByName(`${selectedGropuForPermission.value}`)
    .users.add(userObj.data.LoginName);
  console.log(`${user.email} added to the group successfully.`, users);
  // Prepare array of user IDs for the Approvals field
  // const approvalUserIds = seleccteduserforapproval.map((approvalUser: any) => 
  //   Number(approvalUser.id) // Ensure this is converting the user ID to a number
  // );
 
const approvalUserIds = seleccteduserforapproval.map((approvalUser: any) => 
  Number(approvalUser.id) // Ensure this is converting the user ID to a number within an object
);
  const addedItem = await activeSp.web.lists.getByTitle("DMSFolderDeligationApprovalMaster").items.add({
    Title: selectedEntityForPermission.value,
    SiteTitle: selectedEntityForPermission.value,
    CurrentUser: user.email,
    ApprovalsId: approvalUserIds // Ensure this is an array of objects with Key properties
    
  });
  
  console.log("Added item:", addedItem);
  

  console.log("Added item:", addedItem);
} catch (error) {
  console.error(`Failed to add ${user.email} to the group: `, error);
}

    // await Promise.all(users);
    onSuccess(selectedGropuForPermission.value);
    handleGroupsSelect(selectedGropuForPermission);
    // ritik 03/04/2026 - Reset selected users after adding to the group
    setSelectedUser(null);
setSelectedApprovers([]);
selectedUsersForPermission = undefined;
    seleccteduserforapproval = undefined;
    // end here
  };

const confirmDelete=(group:any,userId:any,groupName:any,userEmail:any,siteTitle:any)=>{
    Swal.fire({
      // Rohit 21/4/26
      title: "Do you want to delete this User/Group?", //Ritik 22/04/2026
      //text: "You won't be able to revert this!",
      icon: "warning",
      showCancelButton: true,
      confirmButtonColor: "#3085d6",
      cancelButtonColor: "#d33",
      confirmButtonText: "Ok"      // Rohit 21/4/26 
    }).then(async(result) => {
      if (result.isConfirmed) {
        await group.users.removeById(userId);

        const getItem = await activeSp.web.lists.getByTitle("DMSFolderDeligationApprovalMaster").items.select("*").filter(`CurrentUser eq '${userEmail}' and SiteTitle eq '${siteTitle}'`)();

        console.log("getItem",getItem);

        if(getItem.length > 0){
          await activeSp.web.lists.getByTitle("DMSFolderDeligationApprovalMaster").items.getById(getItem[0].Id).delete()
          console.log("user successfully removed from the lists")
        }
        console.log(`User with ID ${userId} has been removed from the group '${groupName}'`);
          // to refresh the user table
          // handleEntitySelect(selectedEntityForPermission);
          handleGroupsSelect(selectedGropuForPermission);
        Swal.fire({
          // Rohit 21/4/26 start
          title: "Deleted successfully.",                        
          //text: `User Suucessfuly removed from ${groupName}.`,
          // Rohit 21/4/26 end
          icon: "success"
        });
      }
    });
  }
const handleDeleteUser=async(userId:any,groupName:any,item:any)=>{
    console.log("UserId",userId);
    console.log("item",item);
    console.log("selected entity",selectedEntityForPermission)
    try {

        const subsitecontext=await activeSp.site.openWebById(selectedEntityForPermission.SiteID);
        // Get the group by name
        const group =subsitecontext.web.siteGroups.getByName(groupName);
        // Remove the user from the group using their userId
        confirmDelete(group,userId,groupName,item.email,selectedEntityForPermission.value);
        // await group.users.removeById(userId);
    } catch (error) {
        console.error("Error removing user from group: ", error);
    }
}

// Ritik 22/04/2026
const [folderSearchText, setFolderSearchText] = React.useState("");
const [folderSortConfig, setFolderSortConfig] = React.useState({ key: '', direction: 'ascending' });
const [folderColumnSearch, setFolderColumnSearch] = React.useState({
  email: '', groupName: '', permission: '', Descirption: ''
});

const filteredFolderData = React.useMemo(() => {
  let result = allUsersFromGroups;
  if (folderSearchText) {
    const s = folderSearchText.toLowerCase();
    result = result.filter((u: any) => (u.user || '').toLowerCase().includes(s));
  }
  if (folderColumnSearch.email) {
    const s = folderColumnSearch.email.toLowerCase();
    result = result.filter((u: any) => (u.email || '').toLowerCase().includes(s));
  }
  if (folderColumnSearch.groupName) {
    const s = folderColumnSearch.groupName.toLowerCase();
    result = result.filter((u: any) => (u.groupName || '').toLowerCase().includes(s));
  }
  if (folderColumnSearch.permission) {
    const s = folderColumnSearch.permission.toLowerCase();
    result = result.filter((u: any) => 
      'folder deligation'.includes(s)
    );
  }
  if (folderColumnSearch.Descirption) {
    const s = folderColumnSearch.Descirption.toLowerCase();
    result = result.filter((u: any) => 
      'can create folder(folder created by user will go for approval) and can add, view, update , download documents.'.includes(s)
    );
  }
  if (folderSortConfig.key) {
    result = [...result].sort((a: any, b: any) => {
      const aVal = (a[folderSortConfig.key] || '').toLowerCase();
      const bVal = (b[folderSortConfig.key] || '').toLowerCase();
      if (aVal < bVal) return folderSortConfig.direction === 'ascending' ? -1 : 1;
      if (aVal > bVal) return folderSortConfig.direction === 'ascending' ? 1 : -1;
      return 0;
    });
  }
  return result;
}, [allUsersFromGroups, folderSearchText, folderSortConfig, folderColumnSearch]);
//end here
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
  const currentData = filteredFolderData.slice(startIndex, endIndex);//Ritik 22/04/26

  interface PaginationProps{
    currentPage: number;
    totalPages: any;
    handlePageChange: any;
  }
  const Pagination = ( { currentPage, totalPages, handlePageChange }: PaginationProps) => {
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
    return <div>
          <div className="position-relative">
          <div>
            {/* <button className="btn back-to-admin" onClick={()=>handleReturnToMainFromPermissionTable('')}>Back To Main</button> */}
           
           </div>
            <div style={{
                  
                      position:"relative",
                  
                      marginTop:"0px",
                      padding:"20px",
                      border:"1px solid #ccc",
                      borderRadius:"8px",
                      background:"#fff",
                      clear:"both",
                      float:"left",
                      width:"100%"
  
                    }}>
                    <p className="page-title font-20 text-dark fw-bold" style={{ 
                  
          }}>Manage Folder Delegation</p>
          {/* Ritik 03/04/2026 added site collection dropdown */}
                    <div className="row">
  <div className="col-sm-4 mb-3">
    <label>Location <span className="text-danger">*</span></label>
    <Select
     // Aman 21/4/26 start
      options={siteCollections}
      value={selectedSiteCollection}
      styles={{
      control: (base) => ({
        ...base,
        border: formErrors.location ? "2px solid #fe0100" : base.border,
        backgroundColor: formErrors.location ? "#fee6e6" : base.backgroundColor
      })
    }}
      
      //onChange={handleSiteChange}
      onChange={(selected) => {
  handleSiteChange(selected);
  setFormErrors(prev => ({ ...prev, location: false }));
}}
    // Aman 21/4/26 end
      placeholder="Select Location..."
      noOptionsMessage={() => "No Site Collections Found..."}
      onKeyDown={(e: any) => {
      if (e.key === 'Enter') {
        e.preventDefault(); // Prevents the page from submitting/going back
      }
    }}
    />
  </div>
  <div className="col-sm-4">
              <label>Department <span className="text-danger">*</span></label>
              {/* ritik 03/04/2026 added for entity selection and management */}
    <Select                        
    options={adminPermissionEntity}
    value={selectedEntity}
    // Amab 21/4/26 start
     styles={{
      control: (base) => ({
        ...base,
        border: formErrors.department ? "2px solid #fe0100" : base.border,
        backgroundColor: formErrors.department ? "#fee6e6" : base.backgroundColor
      })
    }}
    // Aman 21/4/26 end
    onChange={(selected: any) => {
      handleEntitySelect(selected);
      setSelectedEntity(selected);
      setFormErrors(prev => ({ ...prev, department: false }));  // Aman 21/4/26
    }}
    placeholder="Select Entity..."
    noOptionsMessage={() => "No Entity Found..."}
    onKeyDown={(e: any) => {
      if (e.key === 'Enter') {
        e.preventDefault(); // Prevents the page from submitting/going back
      }
    }}
              />
              {/* end here */}
                      </div>
                      <div className="col-sm-4">
                        <label>Groups <span className="text-danger">*</span></label>
                        <Select
                            isDisabled
                            options={groups}
                            onChange={(selected: any) =>
                              handleGroupsSelect(selected)
                            }
                            placeholder={`${groupDetails?.value }`}
                            noOptionsMessage={() => "No Groups Found..."}
                          />
                      </div>
                         { <div  className="col-sm-4">
                      {/* Aman 21/4/26 end */}
              <label>Users <span className="text-danger">*</span></label>
              {/* Ritik 03/04/2026 added for user selection and management */}
                        <Select
    options={user}
    value={selectedUser}
    // Aman 21/4/26 start
    styles={{
      control: (base) => ({
        ...base,
        border: formErrors.users ? "2px solid #fe0100" : base.border,
        backgroundColor: formErrors.users ? "#fee6e6" : base.backgroundColor
      })
    }}
    // Aman 21/4/26 end
    onChange={(selected: any) => {
      handleUsersSelect(selected);
      setSelectedUser(selected);
      setFormErrors(prev => ({ ...prev, users: false }));
    }}
    placeholder="Select User..."
    noOptionsMessage={() => "No User Found..."}
    onKeyDown={(e: any) => {
      if (e.key === 'Enter') {
        e.preventDefault(); // Prevents the page from submitting/going back
      }
    }}
              />
              {/* end here  */}
                       </div> 
                       } 
                         { 
                          <div  className="col-sm-4">
              <label>Select Approvers <span className="text-danger">*</span></label>
              {/* Ritik 03/04/2026 added for approver selection and management */}
                        <Select
    isMulti
    options={allUsersFromADMINGroups}
    value={selectedApprovers}
    // Aman 21/4/26 start
    styles={{
      control: (base) => ({
        ...base,
        border: formErrors.approvers ? "2px solid #fe0100" : base.border,
        backgroundColor: formErrors.approvers ? "#fee6e6" : base.backgroundColor
      })
    }}
    // Aman 21/4/26 end
    onChange={(selected: any) => {
      handleUsersforapprovalSelect(selected);
      setSelectedApprovers(selected);
      setFormErrors(prev => ({ ...prev, approvers: false }));  // Aman 21/4/26
    }}
    placeholder="Select User..."
    noOptionsMessage={() => "No User Found..."}
    onKeyDown={(e: any) => {
      if (e.key === 'Enter') {
        e.preventDefault(); // Prevents the page from submitting/going back
      }
    }}
              />
              {/* end here  */}
                       </div> 
                       } 
                      
                    </div>
                    <div style={{
                      display:"flex",
                      gap:"10px",
                      justifyContent:"center"
                     
                    }}>
                         <button style={{padding:'8px 10px', borderRadius:'4px'}} type="button" className="mt-4 btn btn-primary" onClick={handleAddUsers}>
                         Add
                      </button>
                  
                    </div>
                  </div>

                  {showGroupsTable ==="Yes" && (
              <div>
                                            
                        <div style={{padding:'15px',clear:'both', float:'left', marginTop:'15px'}} className={styles.container}>
                        <header style={{padding:'0px 0px 5px 0px'}}>
                        <div className='page-title fw-bold mb-1 font-20'>{selectedEntityForPermission.value} &gt; {
                        groupDetails?.value && groupDetails?.value.includes('_') 
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
                                {/* {groupDetails?.Description || ''} */}
                                Can create folder(folder created by user will go for approval) and can add, view, update documents.
                              </td>
                            </tr>
                        </tbody>
                        </table>
                        </div>
                    {showGroupsUsers ==="Yes" && (<>
                      <div style={{padding:'15px',clear:'both', float:'left', marginTop:'15px'}} className={styles.container}>
                        <header style={{padding:'0px 0px 5px 0px'}}>
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
                                <th style={{minWidth:'55px', maxWidth:'55px'}}>S.No.</th>
                                {/* <th>User</th>
                                <th>User Email</th>
                                <th>Group Name</th>
                                <th>Permission</th>
                                <th>Description</th> */}
                                <th>
  <div>
    <button type="button" style={{cursor:'pointer', background:'none', border:'none', padding:'0', fontWeight:'inherit', fontSize:'inherit'}}
      onClick={() => { setFolderSortConfig(prev => ({ key: 'user', direction: prev.key === 'user' && prev.direction === 'ascending' ? 'descending' : 'ascending' })); }}>
      User <FontAwesomeIcon icon={faSort} />
    </button>
    <input type="text" placeholder="Search User" className="inputcss"
      value={folderSearchText}
      onChange={(e) => { setFolderSearchText(e.target.value); setCurrentPage(1); }}
      onKeyDown={(e: any) => { if (e.key === 'Enter') e.preventDefault(); }}
    />
  </div>
</th>
<th>
  <div>
    <button type="button" style={{cursor:'pointer', background:'none', border:'none', padding:'0', fontWeight:'inherit', fontSize:'inherit'}}
      onClick={() => { setFolderSortConfig(prev => ({ key: 'email', direction: prev.key === 'email' && prev.direction === 'ascending' ? 'descending' : 'ascending' })); }}>
      User Email <FontAwesomeIcon icon={faSort} />
    </button>
    <input type="text" placeholder="Search Email" className="inputcss"
      value={folderColumnSearch.email}
      onChange={(e) => { setFolderColumnSearch(prev => ({...prev, email: e.target.value})); setCurrentPage(1); }}
      onKeyDown={(e: any) => { if (e.key === 'Enter') e.preventDefault(); }}
    />
  </div>
</th>
<th>
  <div>
    <button type="button" style={{cursor:'pointer', background:'none', border:'none', padding:'0', fontWeight:'inherit', fontSize:'inherit'}}
      onClick={() => { setFolderSortConfig(prev => ({ key: 'groupName', direction: prev.key === 'groupName' && prev.direction === 'ascending' ? 'descending' : 'ascending' })); }}>
      Group Name <FontAwesomeIcon icon={faSort} />
    </button>
    <input type="text" placeholder="Search Group" className="inputcss"
      value={folderColumnSearch.groupName}
      onChange={(e) => { setFolderColumnSearch(prev => ({...prev, groupName: e.target.value})); setCurrentPage(1); }}
      onKeyDown={(e: any) => { if (e.key === 'Enter') e.preventDefault(); }}
    />
  </div>
</th>
<th>
  <div>
    <button type="button" style={{cursor:'pointer', background:'none', border:'none', padding:'0', fontWeight:'inherit', fontSize:'inherit'}}
      onClick={() => { setFolderSortConfig(prev => ({ key: 'permission', direction: prev.key === 'permission' && prev.direction === 'ascending' ? 'descending' : 'ascending' })); }}>
      Permission <FontAwesomeIcon icon={faSort} />
    </button>
    <input type="text" placeholder="Search Permission" className="inputcss"
      value={folderColumnSearch.permission}
      onChange={(e) => { setFolderColumnSearch(prev => ({...prev, permission: e.target.value})); setCurrentPage(1); }}
      onKeyDown={(e: any) => { if (e.key === 'Enter') e.preventDefault(); }}
    />
  </div>
</th>
<th>
  <div>
    <button type="button" style={{cursor:'pointer', background:'none', border:'none', padding:'0', fontWeight:'inherit', fontSize:'inherit'}}
      onClick={() => { setFolderSortConfig(prev => ({ key: 'Descirption', direction: prev.key === 'Descirption' && prev.direction === 'ascending' ? 'descending' : 'ascending' })); }}>
      Description <FontAwesomeIcon icon={faSort} />
    </button>
    <input type="text" placeholder="Search Description" className="inputcss"
      value={folderColumnSearch.Descirption}
      onChange={(e) => { setFolderColumnSearch(prev => ({...prev, Descirption: e.target.value})); setCurrentPage(1); }}
      onKeyDown={(e: any) => { if (e.key === 'Enter') e.preventDefault(); }}
    />
  </div>
</th>
                                <th style={{minWidth:'65px', maxWidth:'65px'}}>Action</th>
                            </tr>
                            </thead>
                            <tbody>
                            {currentData.map((item:any, index:any) => (
                                <React.Fragment key={item.userId}>
                                <tr>
                                    <td style={{minWidth:'55px', maxWidth:'55px'}}>
                                 <span className="indexdesign">
                            
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
                                    {/* {item.permission || ''} */}
                                    <span>Folder Deligation</span>
                                    </td>
                                    <td>
                                    {/* {item.Descirption || ''} */}
                                    <span>Can create folder(folder created by user will go for approval) and can add, view, update , download documents.</span>
                                    </td>
                                    <td style={{minWidth:'65px', maxWidth:'65px'}}>
                                    <img
                                        className={styles.deleteicon}
                                        src={require("../assets/del.png")}
                                        alt="Delete"
                                        onClick={(event)=>{
                                            handleDeleteUser(item.userId,item.groupName,item)
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
          </div>;
};

export default ManageFolderDeligation;