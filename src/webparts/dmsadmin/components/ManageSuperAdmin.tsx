import * as React from 'react';
import { getSP } from '../loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import "../../verticalSideBar/components/VerticalSidebar.scss"
import VerticalSideBar from '../../verticalSideBar/components/VerticalSideBar';
import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
import UserContext from '../../../GlobalContext/context';
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import "../../verticalSideBar/components/VerticalSidebar.scss"
import Provider from '../../../GlobalContext/provider';
import { useMediaQuery } from 'react-responsive';
import styles from './Form.module.scss'
import Swal from 'sweetalert2';
import Select from "react-select";

// import context from '../../../GlobalContext/context';
// import classNames from "classnames";
// import { useState, useEffect, useRef , useMemo } from "react";
// import JoditEditor from "jodit-react";
// import Jodit from 'jodit-react';
//let selectedUsersForPermission:any[];  // Aman 22/4/26
// let description:any;


export const ManageSuperAdmin = (props:any) => {
    const sp: SPFI = getSP();
    const { useHide }: any = React.useContext(UserContext);
    const elementRef = React.useRef<HTMLDivElement>(null);
    const isMobile = useMediaQuery({ query: '(max-width: 768px)' });
    const [selectedUser,setSelectedUser]=React.useState([]);
    const [refresh,setRefresh]=React.useState(false);
    const [activeComponent,setActiveComponent]=React.useState('');
    const [user,setUser]=React.useState<any[]>([]);
    const [selectedUsersForPermission, setSelectedUsersForPermission] = React.useState<any[]>([]);  // Aman 22/4/26
    // Aman 21/4/26 start
    const [userError, setUserError] = React.useState(false);
    // Aman 21/4/26 end 
    // const [description,setDescription]=React.useState('');
    console.log("selectedUser",selectedUser);
    console.log("props",props);

    React.useEffect(()=>{
            const fetchUserFromSelectedGroup=async()=>{
                try {
                    // const subsiteContext = await sp.site.openWebById(props.selectedEntityForPermission.SiteID);
                    // const usersFromSelectedGroups = await subsiteContext.web.siteGroups.getByName(`${props.selectedGropuForPermission.value}`).users();
                    // console.log("usersFromSelectedGroups",usersFromSelectedGroups);
                    const usersFromDMSSuperAdmin = await sp.web.siteGroups.getByName('DMSSuper_Admin').users();
                    setSelectedUser(usersFromDMSSuperAdmin);
                  } catch (error) {
                    console.log("error from getting the users from the groups after selecting the groups",error);
                  }
            }
            fetchUserFromSelectedGroup();
    },[refresh]);

    const handleDeleteUser=async(userId:any,UserTitle:any)=>{
        console.log("UserId",userId);
        try {

            // const subsitecontext=await sp.site.openWebById(props.selectedEntityForPermission.SiteID);
            // Get the group by name
            const group = await sp.web.siteGroups.getByName('DMSSuper_Admin');
            // Remove the user from the group using their userId
            // await group.users.removeById(userId);
            // console.log(`User with ID ${userId} has been removed from the super admin group`);
            // onRemove(UserTitle);
            // setRefresh(!refresh);
            confirmDelete(group,userId,UserTitle)
        } catch (error) {
            console.error("Error removing user from group: ", error);
        }
    }

    const handleToggleAddUsers=()=>{
      // Aman 21/4/26 start
        //selectedUsersForPermission = undefined;  
        setSelectedUsersForPermission([]);
        setUserError(false);
      // Aman 21/4/26 end
        setActiveComponent("AddUser");
    }
    
    const handleUsersSelect=(selectedUsers:any)=>{
        console.log("selectedUsers",selectedUsers);
        //selectedUsersForPermission=selectedUsers;
        setSelectedUsersForPermission(selectedUsers || []);
    }

    React.useEffect(()=>{
        const fetchUsers=async()=>{
            const user = await sp.web.siteUsers();
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
        fetchUsers();
        
       
    },[])

    // console.log("description",description);
    // React.useEffect(()=>{
    //      // Add Description
    //      const addDescription=()=>{
    //         const result: string = props.selectedGropuForPermission.value.split("_")[1];
    //         console.log("Description",result);
    //         switch (result) {
    //             case 'Admin':
    //                  setDescription("Full Control - Has full control.");
    //                  break;
    //             case 'Read':
    //                 setDescription("Read - Can view pages and download documents.");
    //                 break;
    //             case 'View':
    //                 setDescription("View - Can only view content.");
    //                 break;
    //             case 'Contribute':
    //                 setDescription("Contribute - Can view, add, update, and delete documents.");
    //                 break;
    //             case 'Initiator':
    //                 setDescription("Initiator - Can view, add, update and delete documents.");
    //                 break;
    //             case 'Approval':
    //                 setDescription("Approval - Can view, add, update and delete documents.");
    //                 break;
    //             case 'AllUsers':
    //                 setDescription("AllUsers - Can view, add, update and delete documents.");
    //                 break;
    //             default:
    //                 setDescription("Unknown role.");
    //         }
    //     }
    //     addDescription();

    // },[])
   

    const handleAddUsers=async()=>{
        console.log("selectedUsersForPermission",selectedUsersForPermission);
        // console.log("selectedGropuForPermission",props.selectedGropuForPermission.value);
        // console.log("selectedEntityForPermission",props.selectedEntityForPermission.value);

        if(selectedUsersForPermission === undefined || selectedUsersForPermission.length === 0){
          setUserError(true);  // Aman 21/4/26
          checkValidation();
          return;
        }

        // New Code for chcek that if user already exist or not
        // const usersFromDMSSuperAdmin = await sp.web.siteGroups.getByName('DMSSuper_Admin').users();
        // console.log("usersFromDMSSuperAdmin",usersFromDMSSuperAdmin);
        // const ids2 = usersFromDMSSuperAdmin.map(item => item.Id)
        // console.log("ids2",ids2);
        // const alReadyPresent=selectedUsersForPermission.filter(item => ids2.includes(Number(item.id)));
        // console.log("alReadyPresent",alReadyPresent);

        // if(alReadyPresent.length>0){
        //     alreadyPresent();
        //     return;
        // }

        // const subsiteContext = await sp.site.openWebById(props.selectedEntityForPermission.SiteID); 
        // selectedUsersForPermission.forEach(async(user:any)=>{
        //   try {
        //     const userObj = await sp.web.ensureUser(user.email);
        //     console.log("userObj",userObj);
        //     const users=await sp.web.siteGroups.getByName('DMSSuper_Admin').users.add(userObj.data.LoginName);
        //     console.log(`${user.email} added to the super admin group successfully.`,users);
        //   } catch (error) {
        //     console.error(`Failed to add ${user.email} to the group: `, error);
        //   }
        // })
        // onSuccess();
        // setRefresh(!refresh);
        // setActiveComponent('');

        // New Code start
        await Promise.all(selectedUsersForPermission.map(async (user: any) => {
            try {
              const userObj = await sp.web.ensureUser(user.email);
              console.log("userObj", userObj);
              const users = await sp.web.siteGroups.getByName('DMSSuper_Admin').users.add(userObj.data.LoginName);
              console.log(`${user.email} added to the super admin group successfully.`, users);
            } catch (error) {
              console.error(`Failed to add ${user.email} to the group: `, error);
            }
          }));
        //selectedUsersForPermission=undefined;
        setSelectedUsersForPermission([]); 
        //   End
        onSuccess();
        setActiveComponent('');
        setRefresh(!refresh);
        
      }
    
    const handleBackToTable=()=>{
      // Aman 21/4/26 start
        //selectedUsersForPermission = undefined;  
        setSelectedUsersForPermission([]); 
        setUserError(false);  
      // Aman 21/4/26 end
        setActiveComponent('');
    }
    const onSuccess=()=>{
        Swal.fire(`Added Successsfully`,"", "success");  // ritik 21/4/26 
    }
    const onRemove=(UserTitle:any)=>{
        Swal.fire(`${UserTitle} Removed Successsfully`,"", "success");
    }
    const checkValidation=()=>{
        Swal.fire("Please fill out the fields!", "All fields are required");
  }

  // Already present erro start
    // const alreadyPresent=()=>{
    // Swal.fire(`User Already Exist`, "Please Change the User", "warning");
    // }
// End

  // Added confirm popup start
  const confirmDelete=(group:any,userId:any,userTitle:any)=>{
    // Aman 21/4/26 start
    // Swal.fire({
    //   title: "Are you sure?",
    //   text: "You won't be able to revert this!",
    //   icon: "warning",
    //   showCancelButton: true,
    //   confirmButtonColor: "#3085d6",
    //   cancelButtonColor: "#d33",
    //   confirmButtonText: "Yes, Removed it!"
    Swal.fire({
  title: "Do you want to delete this request?",
  icon: "warning",
  showCancelButton: true,
  confirmButtonColor: "#3085d6",
  cancelButtonColor: "#d33",
  confirmButtonText: "OK",
  cancelButtonText: "Cancel"
  // Aman 21/4/26 end
    }).then(async(result) => {
      if (result.isConfirmed) {
      await group.users.removeById(userId);
      setRefresh(!refresh);
        Swal.fire({
          // ritik 21/4/26 start
          title: "Deleted successfully.",
          //text: `${userTitle} Suucessfuly Removed.`,
          // ritik 21/4/26 end
          icon: "success"
        });
      }
    });
  }
//   End


  // Code for filter and search start
 const [filters, setFilters] = React.useState({
    SNo: '',
    Title : '',
    // Title: '',
    Email: '',
    Modified: '',
    Status: '',
  
    SubmittedDate: ''
  });
  const [sortConfig, setSortConfig] = React.useState({ key: '', direction: 'ascending' });
  
  const handleFilterChange = (e: React.ChangeEvent<HTMLInputElement>, field: string) => {
    setFilters({
      ...filters,
      [field]: e.target.value,
    });
    console.log(filters , "filters filters")
  };
  
  const handleSortChange = (key: string) => {
    let direction = 'ascending';
    if (sortConfig.key === key && sortConfig.direction === 'ascending') {
      direction = 'descending';
    }
    setSortConfig({ key, direction });
  };
  
  const applyFiltersAndSorting = (data: any[]) => {
    const filteredData = data.filter((item, index) => {
      return (
        (filters.SNo === '' || String(index + 1).includes(filters.SNo)) &&
        (filters.Title === '' || 
          (item.Title && item.Title.toLowerCase().includes(filters.Title.toLowerCase()))) &&
        (filters.Email === '' || 
          (item.Email && item.Email.toLowerCase().includes(filters.Email.toLowerCase()))) &&
        (filters.Modified === '' || 
          (item.Editor.Title && item.Editor.Title.toLowerCase().includes(filters.Modified.toLowerCase()))) &&
        (filters.SubmittedDate === '' || 
          (item.Status && item.Status.toLowerCase().includes(filters.SubmittedDate.toLowerCase())))
      );
    });
  
    const naturalSort = (a: any, b: any) => {
      return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
    };
  
    const sortedData = filteredData.sort((a, b) => {
      if (sortConfig.key === 'SNo') {
        const aIndex = data.indexOf(a);
        const bIndex = data.indexOf(b);
        return sortConfig.direction === 'ascending' ? aIndex - bIndex : bIndex - aIndex;
      } else if (sortConfig.key) {
        const aValue = a[sortConfig.key] ? a[sortConfig.key].toLowerCase() : '';
        const bValue = b[sortConfig.key] ? b[sortConfig.key].toLowerCase() : '';
        return sortConfig.direction === 'ascending' ? naturalSort(aValue, bValue) : naturalSort(bValue, aValue);
      }
      return 0;
    });
  
    return sortedData;
  };
  
  const filteredUserData=applyFiltersAndSorting(selectedUser);
  // end
  
    // Add pagination start
    const [currentPage, setCurrentPage] = React.useState(1);
    const itemsPerPage = 10;
    const totalPages = Math.ceil(filteredUserData.length / itemsPerPage);
    
    const handlePageChange = (pageNumber: any) => {
      if (pageNumber > 0 && pageNumber <= totalPages) {
        setCurrentPage(pageNumber);
      }
    };
    
    const startIndex = (currentPage - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    const currentData = filteredUserData.slice(startIndex, endIndex);
  
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
  return (
<>
              {activeComponent === '' && (
                    <div className={styles.argform}>
                      <div className='row mt-minus30'>
                            <div className='col-md-7 '>
                            <div className='page-title fw-bold mb-0 font-20 d-none'>Manage Super Admin 
                            </div>
                            <div className='mb-2 mt-0 d-none'>
                            <span className='text-muted font-14' style={{
                                color:"Black"
                            }}>User From Super Admin Group Will Have Full Control 1.</span>
                        </div>
                            </div>
                            <div className='col-md-5'>
                            <div className='justify-content-end'>
                                <div className='padd-right1 mt-0'>
                                    
                            <button  style={{height:'40px', borderRadius:'4px', padding:'9px 10px'
                        
                            }}type="button" className='btn btn-primary' onClick={handleToggleAddUsers}>
                                Add User
                            </button>
                            </div></div></div>
                        </div>
                        <div>
                      
                      
                            {/* <a className={styles.backbuttonform}
                                onClick={props.onBack}
                            >
                                <img
                                className={styles.backimg}
                                />
                                <p className={styles.Addtext}>Back</p>
                            </a> */}
                           
                        
                        </div>
                       
                        <div style={{padding:'15px', marginTop:'20px'}} className={styles.container}>
                          <div className='col-md-12'>
                            <div className='page-title fw-bold mb-0 font-20'>Manage Super Admin
                            </div>
                            <div className='mb-2 mt-0'>
                            <span className='text-muted font-14' style={{
                                color:"Black"
                            }}>User From Super Admin Group Will Have Full Control</span>
                        </div>
                            </div>
                        <table className="mtbalenew">
                            <thead>
                            <tr>
                                <th style={{minWidth:'20px',maxWidth:'20px'}}>S.No.</th>
                                <th style={{minWidth:'80px',maxWidth:'80px'}}>User</th>
                                <th>Email</th>
                                <th style={{minWidth:'40px',maxWidth:'40px'}}>Action</th>
                            </tr>
                            </thead>
                            <tbody>
                            {currentData.map((item:any, index:any) => (
                                <React.Fragment key={item.Id}>
                                <tr>
                                    <td style={{minWidth:'20px',maxWidth:'20px'}}>
                                    <span className='indexdesign'> {index + 1}</span>
                                    </td>
                                    <td style={{minWidth:'80px',maxWidth:'80px'}}>
                                    {item.Title || ''}
                                    </td>
                                    <td >
                                    {item.Email || ''}
                                    </td>
                                    <td style={{minWidth:'40px',maxWidth:'40px'}}>
                                    <img
                                        className={styles.deleteicon}
                                        src={require("../assets/del.png")}
                                        alt="Delete"
                                        onClick={(event)=>{
                                            handleDeleteUser(item.Id,item.Title)
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
                    </div>
              )}
              {activeComponent === "AddUser" && 
                (
                <div className={styles.argform}>
                    <div className='row mt-minus30'>
                        <div className='col-md-7'>

                      
                
                    <div className='d-none'>Admin Panel &gt; Manage Super Admin &gt; Add Super Admin</div>
                    </div>
                    <div className='col-md-5'>
                    <div className='padd-right1 mt-0'>
                        <button style={{display:"inline-block",minWidth:"auto"}} type='button' onClick={handleBackToTable} className={styles.backbuttonform}>
                            Back
                        </button>
                    </div>
                    </div>
                    </div>
                    <div style={{
                      
                      position:"relative",
                      
                      marginTop:"20px",
                      padding:"16px",
                      border:"1px solid #ccc",
                      borderRadius:"8px",
                      background:"#fff",

                    }}>
                      {/* Rohit 21/4/26 */}
                    <div className='page-title fw-bold mb-3 font-20'>Add Super Admin</div>  

                        <p style={{
                            color:"Black",
                           
                        }}>Add Users <span className="text-danger">*</span></p>
                        <div style={{
                            gap:"30px",
                            display:"flex"
                        }}>
                            <div  style={{
                                width:"370px"
                            }}>
                              {/* Aman 21/4/26 start */}
                                {/* <Select
                                    isMulti
                                    options={user}
                                    onChange={(selected: any) =>
                                    handleUsersSelect(selected)
                                    }
                                    placeholder="Select User..."
                                    noOptionsMessage={() => "No User Found..."}
                                /> */}
                <Select
    isMulti
    options={user}
    value={selectedUsersForPermission} // Maps the state to UI
    onChange={(selected: any) => {
        handleUsersSelect(selected);
        if (selected && selected.length > 0) {
            setUserError(false);
        }
    }}
    placeholder="Select User..."
    noOptionsMessage={() => "No User Found..."}
    styles={{
        control: (base, state) => ({
            ...base,
            border: userError
                ? "2px solid #dc3545"
                : state.isFocused
                    ? "1px solid #86b7fe"
                    : base.border,
            backgroundColor: userError ? "#fff5f5" : "#fff",
            boxShadow: "none",
            padding: "2px"
        })
    }}
/>
                    {/* Aman 21/4/26 end  */}
                            </div>

                            <div>
                                <button type='button' style={{padding:'9px 10px', borderRadius:'4px'}} className='btn btn-primary' onClick={handleAddUsers}>
                                    Add
                                </button>
                            </div>
                        </div>
                        
                    </div>           
                </div>
                )
              } 
              </> 
           
  )
}
