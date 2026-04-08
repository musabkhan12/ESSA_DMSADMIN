import React, { useEffect } from 'react'
import "bootstrap/dist/css/bootstrap.min.css";
import Select from "react-select";
import { useRef, useState } from "react";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";


interface ManageFolderPermissionProps {
    OthProps: { [key: string]: string };
    onReturnToMain: () => void;
  }


const ManageFolderPermission : React.FC<ManageFolderPermissionProps> = ({
    OthProps,
    onReturnToMain,
  }) =>{

    console.log(OthProps, "oth props");
    const sp: SPFI = getSP();

    const currentUserEmailRef = useRef('');
    const getcurrentuseremail = async()=>{
        const userdata = await sp.web.currentUser();
        currentUserEmailRef.current = userdata.Email;
        defaultValue();
    }


    const [users, setUsers] = React.useState<any[]>([]);
    console.log("Users Array", users);

    const [selectedPermission,setSelectedPermission]=useState([]);
    const [defaultUser,setDefaultUser]=useState<{
        userId:number,
        value: String,
        label: String,
      }[]>([]);

    console.log("Default user array",defaultUser)
  
    const permissionArray:{value:string,label:string}[]=[
      {value:"Admin",label:"Admin"},
      {value:"Contribute",label:"Contribute"},
      {value:"Edit",label:"Edit"},
      {value:"Read",label:"Read"},
      {value:"View",label:"View"}
    ];
  
    const handleUserSelectForPermission=(selectedUser:any)=>{
        console.log("Selected user for permission",selectedUser);
        setDefaultUser(selectedUser);
    }
  
    const handlePermissionSelect=(selectedPermission:any)=>{
        console.log("Selected Permission",selectedPermission)
        const obj={
            value:selectedPermission.value,label:selectedPermission.value
        }
        setSelectedPermission([obj]);
    }


    // Fetch users from SharePoint
  React.useEffect(() => {
    getcurrentuseremail();
    console.log(currentUserEmailRef.current ,"my current id")
    const fetchUsers = async () => {
      try {
        // start
        const siteContext = await sp.site.openWebById(OthProps.SiteID);
        const user0 = await siteContext.web.siteUsers();

        const combineUsersArray=user0.map((user)=>(
              {
              userId:user.Id,
              value: user.Title,
              label: user.Title,
              email: user.Email,
          }
        ))
        setUsers(combineUsersArray);
        console.log("Sub site users",combineUsersArray);
        // const user0 = await sp.web.siteUsers();
        // const [
        //   users,
        //   users1,
        //   users2,
        //   users3,
        //   users4,
        // ] = await Promise.all([
        //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Read`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Initiator`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Contribute`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Admin`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_View`).users(),
        // ]);
        // console.log(users, "users ", users1,users2,users3,users4);
        // const combineArray = [
        //   ...(users || []),
        //   ...(users1 || []),
        //   ...(users2 || []),
        //   ...(users3 || []),
        //   ...(users4 || []),
        // ];
        // setUsers(
        //   combineArray.map((user) => ( 
        //   {
        //     userId:user.Id,
        //     value: user.Title,
        //     label: user.Title,
        //     email: user.Email,
        //   }
        // ))
        // );
        // console.log("combineArray", combineArray);
        // end
      } catch (error) {
        console.error("Error fetching site users:", error);
      }
    };

    fetchUsers();
  }, []);

// fetch the initial data from the  DMSFolderPrivacy list
    const defaultValue=async()=>{
        try {
            const fetchData=await sp.web.lists.getByTitle("DMSFolderPrivacy").items.select("User","UserID","UserPermission").filter(`SiteName eq '${OthProps.SiteTitle}' and DocumentLibraryName eq '${OthProps.DocumentLibraryName}' and CurrentUser eq '${currentUserEmailRef.current}'`)();
            // const permission=fetchData[0].UserPermission;
            // const objectPermission={value:permission,label:permission};
            // setSelectedPermission([objectPermission]);

            // Initialize array to store the default users
            const arrayToStoreDefaultUser = fetchData.map((user) => ({
                userId: user.UserID,
                value: user.User,
                label: user.User,
                Permission:user.UserPermission
 
            }));
            // filter the data.
            const filteredData = arrayToStoreDefaultUser.filter(item => item.userId !== null && item.value !== null && item.label !== null);
            const permission=filteredData[0].Permission;
            const objectPermission={value:permission,label:permission};
            console.log("Permission Object",objectPermission);
            setSelectedPermission([objectPermission]);
            setDefaultUser(filteredData);
            console.log("Fetch data from DMSFolderPrivacy",fetchData);

        } catch (error) {
            console.log("Erroe fetching data from DMSFolderPrivacy",error);
        }
    }
    

  const handleCreate=async()=>{
    console.log("create called");
    console.log("selected User Array",defaultUser);
    console.log("selected permission",selectedPermission);

    try {
        const fetchData=await sp.web.lists.getByTitle("DMSFolderPrivacy").items.select("User","UserID","UserPermission","ID").filter(`SiteName eq '${OthProps.SiteTitle}' and DocumentLibraryName eq '${OthProps.DocumentLibraryName}' and CurrentUser eq '${currentUserEmailRef.current}'`)();

        console.log("fetched user for delete and create new user",fetchData);
        fetchData.forEach(async(user)=>{
            try {
                // console.log("User->",user.User)
                // await sp.web.lists.getByTitle("DMSFolderPrivacy").items.getById(user.ID).delete();
                // console.log(`Item with ID: ${user.ID} has been deleted successfully.`);
                if(user.User !== null){
                    // console.log("Skip");
                    await sp.web.lists.getByTitle("DMSFolderPrivacy").items.getById(user.ID).delete();
                    console.log(`Item with ID: ${user.ID} has been deleted successfully.`);
                }
            } 
            catch (error) {
                console.log("Error deleting item from DMSFolderPrivacy",error);
            }
        })

        const payloadForDMSFolderPrivacy={
            SiteName:OthProps.SiteTitle,
            DocumentLibraryName:OthProps.DocumentLibraryName,
            CurrentUser:currentUserEmailRef.current,
            IsModified:true,
            UserPermission:selectedPermission[0].value
        }
        defaultUser.forEach(async(user)=>{
            try {
                
                (payloadForDMSFolderPrivacy as any).User=user.value;
                (payloadForDMSFolderPrivacy as any).UserID=user.userId;
                console.log("Payload for DMSFolderPrivacy",payloadForDMSFolderPrivacy);
                const addedItem = await sp.web.lists.getByTitle("DMSFolderPrivacy").items.add(payloadForDMSFolderPrivacy);
                console.log("Item added to the list DMSFolderPrivacy after selected value ",addedItem);
            } catch (error) {
                console.log("error adding data to the DMSFolderPrivacy",error);
            }
            
        })
     

    } catch (error) {
        console.log("error creating data insode the handle create function",error);
    }

  }
  const hidemanageworkflow=()=>{
    onReturnToMain()
    
  }

  return (
    <div className="container mt-4 second">
    <div className="modal show d-block" tabIndex={-1}>
            <div style={{minWidth:"60%"}} className="modal-dialog">
                <div className="modal-content" style={{
                            
                            padding:"0px"
                            
                }}>
                    <div className=" marginleftcard" style={{
                        height: "auto",
                        width:"auto",
                    }}>
                    <h5 className="mb-3 " style={{
                        display:"block",
                          textAlign: "left",
    width: "200px"
                        // textAlign:"left200px"
                    }}>
                          <span className="crossicon" onClick={()=>hidemanageworkflow()}>&#10006;</span>
                        <strong>Manage Permission</strong>
                    </h5>
                    <div className="row mb-3 approvalheirarcystyle">
                            <div className="col-12 col-md-6">
                                <Select
                                    value={defaultUser}
                                    isMulti
                                    options={users}
                                    onChange={(selected: any) =>
                                    handleUserSelectForPermission(selected)
                                    }
                                    placeholder="Enter names or email addresses..."
                                    noOptionsMessage={() => "No User Found..."}
                                />
                            </div>
                            <div className="col-12 col-md-6" style={{
                            width:"auto"
                            }}>
                                <Select
                                    value={selectedPermission}
                                    options={permissionArray}
                                    onChange={(selected: any) =>
                                    handlePermissionSelect(selected)
                                    }
                                    placeholder="Select Permission"
                                    noOptionsMessage={() => "No Such Permission Find"}
                                />
                            </div>

                    </div>
                    </div>
                    <div className="modal-footer">
                            <button type="button" className="btn btn-primary" 
                            onClick={handleCreate}
                            >
                            Create
                            </button>
                            <button type="button" className="btn btn-secondary" 
                            //   onClick={toggleModal}
                            onClick={()=>onReturnToMain()}
                            >
                                Cancel{" "}
                            </button>
                    </div>
                </div>
            </div>
    </div>
</div>
  )
}

export default ManageFolderPermission