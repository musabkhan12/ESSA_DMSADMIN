import * as React from 'react';
import { getSP } from '../loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import "../../verticalSideBar/components/VerticalSidebar.scss"
import VerticalSideBar from '../../verticalSideBar/components/VerticalSideBar';
import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
import UserContext from '../../../GlobalContext/context';
import "../../../CustomCss/mainCustom.scss";
import "../../verticalSideBar/components/VerticalSidebar.scss"
import Provider from '../../../GlobalContext/provider';
import { useMediaQuery } from 'react-responsive';
import styles from './Form.module.scss'
import Swal from 'sweetalert2';
import Select from "react-select";
import {  spfi } from '@pnp/sp';
import { SPFx } from "@pnp/sp";
import { WebPartContext } from '@microsoft/sp-webpart-base';
 
// import context from '../../../GlobalContext/context';
// import classNames from "classnames";
// import { useState, useEffect, useRef , useMemo } from "react";
// import JoditEditor from "jodit-react";
// import Jodit from 'jodit-react';
let selectedUsersForPermission:any[];
// let description:any;
 
//Aman 27/2/26
interface ManageSuperProps {
  sp: SPFI;
  context: WebPartContext;
}
export const ManageSuper: React.FC<ManageSuperProps> = ({ sp, context }) => {
    
    const { useHide }: any = React.useContext(UserContext);
    const elementRef = React.useRef<HTMLDivElement>(null);
    const isMobile = useMediaQuery({ query: '(max-width: 768px)' });
    const [selectedUser,setSelectedUser]=React.useState([]);
    const [refresh,setRefresh]=React.useState(false);
    const [activeComponent,setActiveComponent]=React.useState('');
    const [user,setUser]=React.useState<any[]>([]);
    const [essa,setEssa]=React.useState<any>(null);

    const [validationError, setValidationError] = React.useState(false);
 
 
 
    const [sites, setSites] = React.useState<any[]>([]);
const [selectedSite, setSelectedSite] = React.useState<any>(null);
 
 const [siteFilter, setSiteFilter] = React.useState<string>('All');
    // const [description,setDescription]=React.useState('');
    console.log("selectedUser",selectedUser);
    //Aman 27/2/26 Line removed of console log
 
    // React.useEffect(()=>{
    //         const fetchUserFromSelectedGroup=async()=>{
    //             try {
    //                 // const subsiteContext = await sp.site.openWebById(props.selectedEntityForPermission.SiteID);
    //                 // const usersFromSelectedGroups = await subsiteContext.web.siteGroups.getByName(`${props.selectedGropuForPermission.value}`).users();
    //                 // console.log("usersFromSelectedGroups",usersFromSelectedGroups);
    //                 const usersFromDMSSuperAdmin = await sp.web.siteGroups.getByName('DMSSuper_Admin').users();
    //                 setSelectedUser(usersFromDMSSuperAdmin);
    //               } catch (error) {
    //                 console.log("error from getting the users from the groups after selecting the groups",error);
    //               }
    //         }
    //         fetchUserFromSelectedGroup();
    // },[refresh]);
 
  // Addhyan - here i fetch all site MasterSiteCollection
  const Essa = spfi("https://officeindia.sharepoint.com/sites/Essa/").using(SPFx(context));
  React.useEffect(() => {
  const fetchAllSitesUsers = async () => {
    try {
      // 1. Master list se sab site URLs lao
      const items = await Essa.web.lists
        .getByTitle("MasterSiteCollection")
        .items
        .select("Id", "Title", "SiteURL")
        .getAll();
 
      console.log("Sites:", items);
 
      // 2. Har site ke liye users fetch karo
      const allUsersPromises = items.map(async (site) => {
        try {
          const spSite = spfi(site.SiteURL).using(SPFx(context));
 
          const users = await spSite.web.siteGroups
            .getByName("DMSSuper_Admin")
            .users();
            console.log(`Users from ${site.SiteURL}:`, users);
 
          // 3. Har user me site info add karo
          return users.map((u) => ({
            ...u,
            siteName: site.Title,
            siteURL: site.SiteURL,
          }));
 
        } catch (err) {
          console.log(`Error in site ${site.SiteURL}`, err);
          return []; // error aaye to empty return
        }
      });
 
      // 4. Sab promises resolve karo
      const results = await Promise.all(allUsersPromises);
 
      // 5. Flatten array (2D → 1D)
     const combinedUsers = results.flat();
 
console.log("Combined Users:", combinedUsers);
 
// ✅ Duplicate remove
const uniqueUsers = combinedUsers.filter(
  (user, index, self) =>
    index === self.findIndex(
      (u) => u.Email === user.Email && u.siteURL === user.siteURL
    )
);
 
console.log("Unique Users:", uniqueUsers);
 
// ✅ Final state me unique data set karo
setSelectedUser(uniqueUsers);
 
    } catch (error) {
      console.error("Error fetching users:", error);
    }
  };
 
  fetchAllSitesUsers();
}, [refresh]);
 
  
 
 
 
 
  
 
 
 
 
 
 
 
 
 
 
 
    // const handleDeleteUser=async(userId:any,UserTitle:any)=>{
    //     console.log("UserId",userId);
    //     try {
 
    //         // const subsitecontext=await sp.site.openWebById(props.selectedEntityForPermission.SiteID);
    //         // Get the group by name
    //         const group = await sp.web.siteGroups.getByName('DMSSuper_Admin');
    //         // Remove the user from the group using their userId
    //         // await group.users.removeById(userId);
    //         // console.log(`User with ID ${userId} has been removed from the super admin group`);
    //         // onRemove(UserTitle);
    //         // setRefresh(!refresh);
    //         confirmDelete(group,userId,UserTitle)
    //     } catch (error) {
    //         console.error("Error removing user from group: ", error);
    //     }
    // }
    const handleDeleteUser = async (userId : any, userTitle: any, siteURL: any) => {
  try {
    const spSite = spfi(siteURL).using(SPFx(context));
 
    const group = await spSite.web.siteGroups.getByName('DMSSuper_Admin');
 
    confirmDelete(group, userId, userTitle);
 
  } catch (error) {
    console.error("Delete error:", error);
  }
};
 
 
 
 
 
 
React.useEffect(() => {
  const fetchSites = async () => {
    try {
      const items = await Essa.web.lists
        .getByTitle("MasterSiteCollection")
        .items
        .select("Title", "SiteURL")
        .getAll();
 
      const formattedSites = items.map(site => ({
        label: site.Title,
        value: site.SiteURL
      }));
 
      setSites(formattedSites);
 
    } catch (error) {
      console.error("Error fetching sites:", error);
    }
  };
 
  fetchSites();
}, []);
 
    const handleToggleAddUsers=()=>{
        setActiveComponent("AddUser");
    }
    
    const handleUsersSelect=(selectedUsers:any)=>{
        console.log("selectedUsers",selectedUsers);
        selectedUsersForPermission=selectedUsers;
    }
 
    //
    
    React.useEffect(() => {
  const fetchUsersForSite = async () => {
    try {
      if (!selectedSite) {
        setUser([]); // clear dropdown
        return;
      }
 
      const spSite = spfi(selectedSite.value).using(SPFx(context));
 
      const users = await spSite.web.siteUsers();
 
      const formatted = users.map((u) => ({
        id: String(u.Id),
        value: u.Title,
        email: u.Email,
        label: u.Title
      }));
 
      console.log("Users for selected site:", formatted);
 
      setUser(formatted);
 
    } catch (error) {
      console.error("Error fetching users for site:", error);
    }
  };
 
  fetchUsersForSite();
}, [selectedSite]);
 
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
   
 
    // const handleAddUsers=async()=>{
    //     console.log("selectedUsersForPermission",selectedUsersForPermission);
    //     // console.log("selectedGropuForPermission",props.selectedGropuForPermission.value);
    //     // console.log("selectedEntityForPermission",props.selectedEntityForPermission.value);
 
    //     if(selectedUsersForPermission === undefined || selectedUsersForPermission.length === 0){
    //       checkValidation();
    //       return;
    //     }
 
    //     // New Code for chcek that if user already exist or not
    //     // const usersFromDMSSuperAdmin = await sp.web.siteGroups.getByName('DMSSuper_Admin').users();
    //     // console.log("usersFromDMSSuperAdmin",usersFromDMSSuperAdmin);
    //     // const ids2 = usersFromDMSSuperAdmin.map(item => item.Id)
    //     // console.log("ids2",ids2);
    //     // const alReadyPresent=selectedUsersForPermission.filter(item => ids2.includes(Number(item.id)));
    //     // console.log("alReadyPresent",alReadyPresent);
 
    //     // if(alReadyPresent.length>0){
    //     //     alreadyPresent();
    //     //     return;
    //     // }
 
    //     // const subsiteContext = await sp.site.openWebById(props.selectedEntityForPermission.SiteID);
    //     // selectedUsersForPermission.forEach(async(user:any)=>{
    //     //   try {
    //     //     const userObj = await sp.web.ensureUser(user.email);
    //     //     console.log("userObj",userObj);
    //     //     const users=await sp.web.siteGroups.getByName('DMSSuper_Admin').users.add(userObj.data.LoginName);
    //     //     console.log(`${user.email} added to the super admin group successfully.`,users);
    //     //   } catch (error) {
    //     //     console.error(`Failed to add ${user.email} to the group: `, error);
    //     //   }
    //     // })
    //     // onSuccess();
    //     // setRefresh(!refresh);
    //     // setActiveComponent('');
 
    //     // New Code start
    //     await Promise.all(selectedUsersForPermission.map(async (user: any) => {
    //         try {
    //           const userObj = await sp.web.ensureUser(user.email);
    //           console.log("userObj", userObj);
    //           const users = await sp.web.siteGroups.getByName('DMSSuper_Admin').users.add(userObj.data.LoginName);
    //           console.log(`${user.email} added to the super admin group successfully.`, users);
    //         } catch (error) {
    //           console.error(`Failed to add ${user.email} to the group: `, error);
    //         }
    //       }));
    //     selectedUsersForPermission=undefined;
    //     //   End
    //     onSuccess();
    //     setActiveComponent('');
    //     setRefresh(!refresh);
        
    //   }
const handleAddUsers = async () => {
 
  console.log("Selected Site:", selectedSite);
  console.log("Selected Users:", selectedUsersForPermission);
 
  // if (!selectedSite) {
  //   Swal.fire("Please select a Location!");
  //   setValidationError(true);   // Aman 8/04/26
  //   return;
  // }
 
  if (!selectedUsersForPermission || selectedUsersForPermission.length === 0 || !selectedSite) {
    setValidationError(true);   // Aman 8/04/26
    checkValidation();
    return;
  }
  setValidationError(false);   // Aman 8/04/26 
 
  try {
    const spSite = spfi(selectedSite.value).using(SPFx(context));
 
    // Get existing users in the group to check for duplicates
    const existingUsers = await spSite.web.siteGroups
      .getByName("DMSSuper_Admin")
      .users();
 
    const existingEmails = existingUsers.map(u => u.Email);
 
    const results = await Promise.all(
      selectedUsersForPermission.map(async (user: any) => {
        if (existingEmails.includes(user.email)) {
          console.log(`⚠️ ${user.email} already in group`);
          return { success: false, user: user.email, error: "Already in group" };
        }
 
        try {
          console.log("Adding user:", user);
 
          // 🔴 IMPORTANT FIX
          const userObj = await spSite.web.ensureUser(user.email || user.value);
 
          console.log("User ensured:", userObj);
 
          await spSite.web.siteGroups
            .getByName("DMSSuper_Admin")
            .users.add(userObj.data.LoginName);
 
          console.log(`✅ ${user.email} added successfully`);
          return { success: true, user: user.email };
 
        } catch (error) {
          console.error(`❌ Failed to add ${user.email}`, error);
          return { success: false, user: user.email, error };
        }
      })
    );
 
    const successes = results.filter(r => r.success);
    const failures = results.filter(r => !r.success);
 
    if (successes.length > 0) {
      Swal.fire("Users Added Successfully", `Added: ${successes.length}, Skipped (already in group): ${failures.filter(f => f.error === "Already in group").length}, Failed: ${failures.filter(f => f.error !== "Already in group").length}`, "success");
    } else if (failures.every(f => f.error === "Already in group")) {
      Swal.fire("Users Already in Group", "All selected users are already super admins on this site", "info");
    } else {
      Swal.fire("Failed to Add Users", `Failed: ${failures.length}`, "error");
    }
 
    selectedUsersForPermission = undefined;
 
    setActiveComponent('');
    setRefresh(!refresh);
 
  } catch (error) {
    console.error("❌ Add user error:", error);
    Swal.fire("Error", "An error occurred while adding users", "error");
  }
};
    
    const handleBackToTable=()=>{
        setActiveComponent('');
    }
    const onSuccess=()=>{
        Swal.fire(`Users Added Successsfully`,"", "success");
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
    Swal.fire({
      title: "Are you sure?",
      text: "You won't be able to revert this!",
      icon: "warning",
      showCancelButton: true,
      confirmButtonColor: "#3085d6",
      cancelButtonColor: "#d33",
      confirmButtonText: "Yes, Removed it!"
    }).then(async(result) => {
      if (result.isConfirmed) {
      await group.users.removeById(userId);
      setRefresh(!refresh);
        Swal.fire({
          title: "Removed!",
          text: `${userTitle} Suucessfuly Removed.`,
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
          (item.Editor && item.Editor.Title && item.Editor.Title.toLowerCase().includes(filters.Modified.toLowerCase()))) &&
        (filters.SubmittedDate === '' ||
          (item.Status && item.Status.toLowerCase().includes(filters.SubmittedDate.toLowerCase()))) &&
        (siteFilter === 'All' || item.siteName === siteFilter)
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
                            <div className='col-md-7'>
                            <div className='fw-bold mb-0 font-20 d-none'>Admin Panel &gt; Manage Location Admin
                            </div>
                            <div className='mb-1 mt-0 d-none'>
                            <span className='text-muted font-14' style={{
                                color:"Black", fontWeight:'500'
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
                       
                        <div style={{padding:'15px', marginTop:'25px', marginBottom:'30px'}} className={styles.container}>
                             <div className="d-flex align-items-center justify-content-between"> 
                              
                              <div style={{lineHeight:'1.3'}} className='page-title fw-bold mb-0 font-20'>Admin Panel &gt; Manage Location Admin
                                <div className='mb-2 mt-0'>
                            <span className='text-muted font-14' style={{
                                color:"Black", fontWeight:'500'
                            }}>User From Super Admin Group Will Have Full Control 1.</span>
                        </div>
                            </div>

                            
                        
                        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '15px' }}>
  <label style={{ fontWeight: 600, minWidth: '120px' }}>Location Filter:</label>
  <select
    className="form-select"
    style={{ width: '250px' }}
    value={siteFilter}
    onChange={(e) => {
      setSiteFilter(e.target.value);
      setCurrentPage(1);
    }}
  >
    <option value="All">All</option>
    {sites.map((site) => (
      <option key={site.value} value={site.label}>
        {site.label}
      </option>
    ))}
  </select>
</div>
</div> 
                        <table className="mtbalenew">
                            <thead>
                            <tr>
                                <th style={{minWidth:'40px',maxWidth:'40px'}}>S.No.</th>
                                <th style={{minWidth:'130px',maxWidth:'130px'}}>Site Location</th>
                                <th style={{minWidth:'100px',maxWidth:'100px'}}>User</th>
                                <th style={{minWidth:'200px',maxWidth:'200px'}}>Email</th>
                                
                                <th style={{minWidth:'40px',maxWidth:'40px'}}>Action</th>
                            </tr>
                            </thead>
                            <tbody>
                            {currentData.map((item:any, index:any) => (
                                <React.Fragment key={item.Id + item.siteURL}>
                                <tr>
                                    {/* <td style={{minWidth:'20px',maxWidth:'20px'}}>
                                    <span className='indexdesign'> {index + 1}</span>
                                    </td> */}
                                     <td style={{ minWidth: '40px', maxWidth: '40px' }}>
                                    <span className="indexdesign">
                                      {/* {index + 1} */}
                                      {(currentPage - 1) * itemsPerPage + index + 1}
                                    </span>
                                  </td>
                                  <td style={{minWidth:'130px',maxWidth:'130px'}}>
                                    {item.siteName || 'N/A'}
                                    </td>
                                     {/* <td style={{ minWidth: '55px', maxWidth: '55px' }}><span className="indexdesign">{index + 1}</span></td> */}
                                    <td style={{minWidth:'100px',maxWidth:'100px'}}>
                                    {item.Title || ''}
                                    </td>
                                    <td style={{minWidth:'200px',maxWidth:'200px'}}>
                                    {item.Email || ''}
                                    </td>
                                    
                                    <td style={{minWidth:'40px',maxWidth:'40px'}}>
                                    <img
                                        className={styles.deleteicon}
                                        src={require("../assets/del.png")}
                                        alt="Delete"
                                       onClick={() => handleDeleteUser(item.Id, item.Title, item.siteURL)}
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
                      
                      marginTop:"25px",
                      padding:"20px",
                      border:"1px solid #ccc",
                      borderRadius:"8px",
                      background:"#fff",
 
                    }}>
                      <div className="page-title fw-bold mb-2 font-20 mt-0">Admin Panel &gt; Manage Location Admin &gt; Add Location Admin</div>
                        <p style={{
                            color:"Black",
                           
                        }}>Add Users</p>
                        <div style={{
                            gap:"30px",
                            display:"flex"
                        }}>
                          <div style={{ width: "300px", marginBottom: "10px" }}>
  {/* <Select
    options={sites}
    value={selectedSite}
    onChange={(site: any) => setSelectedSite(site)}
    placeholder="Select Location..."
  /> */}
  <Select
    options={sites}
    value={selectedSite}
    onChange={(site: any) => {
      setSelectedSite(site);
      if(site) setValidationError(false);
    }}
    placeholder="Select Location..."
    styles={{
      control: (base) => ({
        ...base,
        // borderColor: validationError && !selectedSite ? "red" : base.borderColor,
        // '&:hover': { borderColor: validationError && !selectedSite ? "red" : base.borderColor }
      })
    }}
  />
</div>
                            <div  style={{
                                width:"370px"
                            }}>
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
    onChange={(selected: any) => {
      handleUsersSelect(selected);
      if(selected && selected.length > 0) setValidationError(false);
    }}
    placeholder="Select User..."
    noOptionsMessage={() => "No User Found..."}
    onKeyDown={(e: any) => {
      if (e.key === 'Enter') {
        e.preventDefault(); // Prevents the page from submitting/going back
      }
    }}
    styles={{
      control: (base) => ({
        ...base,
        // borderColor: validationError && (!selectedUsersForPermission || selectedUsersForPermission.length === 0) ? "red" : base.borderColor,
        // '&:hover': { borderColor: validationError && (!selectedUsersForPermission || selectedUsersForPermission.length === 0) ? "red" : base.borderColor }
      })
    }}
  />
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
