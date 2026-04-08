import * as React from 'react';
import { getSP } from '../loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import "../../verticalSideBar/components/VerticalSidebar.scss"
import VerticalSideBar from '../../verticalSideBar/components/VerticalSideBar';
import UserContext from '../../../GlobalContext/context';
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import "../../verticalSideBar/components/VerticalSidebar.scss"
import Provider from '../../../GlobalContext/provider';
import { useMediaQuery } from 'react-responsive';
import styles from './Form.module.scss'
import Swal from 'sweetalert2';
import Select from "react-select";
import { faSort } from '@fortawesome/free-solid-svg-icons';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';

// import context from '../../../GlobalContext/context';
// import classNames from "classnames";
// import { useState, useEffect, useRef , useMemo } from "react";
// import JoditEditor from "jodit-react";
// import Jodit from 'jodit-react';
let selectedUsersForPermission:any[];
// let description:any;

export const ManagePermission = (props:any) => {
    const sp: SPFI = getSP();
    const { useHide }: any = React.useContext(UserContext);
    const elementRef = React.useRef<HTMLDivElement>(null);
    const isMobile = useMediaQuery({ query: '(max-width: 768px)' });
    const [selectedUser,setSelectedUser]=React.useState([]);
    const [refresh,setRefresh]=React.useState(false);
    const [activeComponent,setActiveComponent]=React.useState('');
    const [user,setUser]=React.useState<any[]>([]);
    const [description,setDescription]=React.useState('');
    console.log("selectedUser",selectedUser);
    console.log("props",props);

    React.useEffect(()=>{
            const fetchUserFromSelectedGroup=async()=>{
                try {
                    const subsiteContext = await sp.site.openWebById(props.selectedEntityForPermission.SiteID);
                    const usersFromSelectedGroups = await subsiteContext.web.siteGroups.getByName(`${props.selectedGropuForPermission.value}`).users();
                    console.log("usersFromSelectedGroups",usersFromSelectedGroups);
                    setSelectedUser(usersFromSelectedGroups);
                  } catch (error) {
                    console.log("error from getting the users from the groups after selecting the groups",error);
                  }
            }
            fetchUserFromSelectedGroup();
    },[refresh]);

    const handleDeleteUser=async(userId:any,UserTitle:any)=>{
        console.log("UserId",userId);
        try {

            const subsitecontext=await sp.site.openWebById(props.selectedEntityForPermission.SiteID);
            // Get the group by name
            const group = await subsitecontext.web.siteGroups.getByName(props.selectedGropuForPermission.value);
            // Remove the user from the group using their userId
            // await group.users.removeById(userId);
            // console.log(`User with ID ${userId} has been removed from the group '${props.selectedGropuForPermission.value}'`);
            // onRemove(UserTitle);
            // setRefresh(!refresh);

            confirmDelete(group,userId,UserTitle);
        } catch (error) {
            console.error("Error removing user from group: ", error);
        }
    }

    const handleToggleAddUsers=()=>{
        setActiveComponent("AddUser");
    }
    
    const handleUsersSelect=(selectedUsers:any)=>{
        console.log("selectedUsers",selectedUsers);
        selectedUsersForPermission=selectedUsers;
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
    React.useEffect(()=>{
         // Add Description
         const addDescription=()=>{
            const result: string = props.selectedGropuForPermission.value.split("_")[1];
            console.log("Description",result);
            switch (result) {
                case 'Admin':
                     setDescription("Full Control - Has full control.");
                     break;
                case 'Read':
                    setDescription("Read - Can view pages and download documents.");
                    break;
                case 'View':
                    setDescription("View - Can only view content.");
                    break;
                case 'Contribute':
                    setDescription("Contribute - Can view, add, update and download documents.");
                    break;
                case 'Initiator':
                    setDescription("Initiator - Can view, add, update and download documents.");
                    break;
                case 'Approval':
                    setDescription("Approval - Can view, add, update and download documents.");
                    break;
                case 'AllUsers':
                    setDescription("AllUsers - Can view, add, update and download documents.");
                    break;
                default:
                    setDescription("Unknown role.");
            }
        }
        addDescription();

    },[])
   

    const handleAddUsers=async()=>{
        console.log("selectedUsersForPermission",selectedUsersForPermission);
        console.log("selectedGropuForPermission",props.selectedGropuForPermission.value);
        console.log("selectedEntityForPermission",props.selectedEntityForPermission.value);

        if(selectedUsersForPermission === undefined || selectedUsersForPermission.length === 0){
          checkValidation();
          return;
        }
        const subsiteContext = await sp.site.openWebById(props.selectedEntityForPermission.SiteID); 
        selectedUsersForPermission.forEach(async(user:any)=>{
          try {
            const userObj = await sp.web.ensureUser(user.email);
            console.log("userObj",userObj);
            const users=await subsiteContext.web.siteGroups.getByName(`${props.selectedGropuForPermission.value}`).users.add(userObj.data.LoginName);
            console.log(`${user.email} added to the group successfully.`,users);
          } catch (error) {
            console.error(`Failed to add ${user.email} to the group: `, error);
          }
        })
        onSuccess();
        setRefresh(!refresh);
        setActiveComponent('');
        
      }
    
    const handleBackToTable=()=>{
        setActiveComponent('');
    }
    const onSuccess=()=>{
        Swal.fire(`Users Added Successsfully`,"", "success");
    }
    // const onRemove=(UserTitle:any)=>{
    //     Swal.fire(`${UserTitle} Removed Successsfully`,"", "success");
    // }
    const checkValidation=()=>{
        Swal.fire("Please fill out the fields!", "All fields are required");
  }
// Added conform delete start
  const confirmDelete=(group:any,userId:any,userName:any)=>{
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
  
        // to refresh the user table
        setRefresh(!refresh);
        Swal.fire({
          title: "Removed!",
          text: `${userName} Suucessfuly removed.`,
          icon: "success"
        });
      }
    });
  }
//end

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
   <div className={styles.maincont}>
              {activeComponent === '' && (
                    <div className={styles.argform}>
                        <div className='row'>
                        <div className='col-md-7 pt-0'>
                        <div className='page-title fw-bold mb-0 font-20'>{props.selectedEntityForPermission?.value} &gt; {props.selectedGropuForPermission?.value}
                       
                        </div>
                        <div>
                            <span className='text-muted font-14'>{description}</span>
                        </div>
                        </div>
                        <div className='col-sm-5'>
                          <div className='padd-right1 mt-0'>
                        <div className={styles.actions}>
                            <a className={styles.backbuttonform}
                                onClick={props.onBack}
                            >
                                <img
                                className={styles.backimg}
                                />
                                <p className={styles.Addtext}>Back</p>
                            </a>
                         
                        </div>
                        </div>
                        </div></div>
                       
                        <div style={{padding:'15px', marginTop:'15px'}} className={styles.container}>
                        <table className='mtbalenew'>

                            <thead>
                            <tr>
                                {/* <th className={styles.serialno}>S.No.</th> */}
                                <th style={{  minWidth: '55px', maxWidth: '55px' }}>
                    
                                  <div className="">
                                    <span>S.No.</span>
                                    <span onClick={() => handleSortChange('SNo')}>
                                      <FontAwesomeIcon icon={faSort} />
                                    </span>
                                  </div>
                                  {/* <div className="bd-highlight">
                                    <input
                                      type="text"
                                      placeholder="index"
                                      onChange={(e) => handleFilterChange(e, 'SNo')}
                                      className="inputcss"
                                      style={{ width: '100%' }}          
                                    />
                                  </div> */}
                                </th>
                                {/* <th className={styles.tabledept}>User</th> */}
                                <th  >
                                  <div >
                                  <div  >
                                    <span >User</span> &nbsp;&nbsp;
                                    <span className="Sorting" onClick={() => handleSortChange('Title')}>
                                      <FontAwesomeIcon icon={faSort} /> 
                                    </span>
                                  </div>
                                  {/* <div className=" bd-highlight">
                                    <input 
                                      type="text" 
                                      placeholder="Filter by User" 
                                      onChange={(e) => handleFilterChange(e, 'Title')}
                                      className='inputcss' 
                                      style={{ width: '100%' }} 
                                    />
                                  </div> */}
                                  </div>
                                </th>
                                {/* <th  className={styles.tabledept}>Email</th> */}
                                <th  >
                                  <div >
                                  <div >
                                    <span >Email</span> &nbsp;&nbsp;
                                    <span className="Sorting" onClick={() => handleSortChange('Title')}>
                                      <FontAwesomeIcon icon={faSort} /> 
                                    </span>
                                  </div>
                                  {/* <div className=" bd-highlight">
                                    <input 
                                      type="text" 
                                      placeholder="Filter by Email" 
                                      onChange={(e) => handleFilterChange(e, 'Email')}
                                      className='inputcss' 
                                      style={{ width: '100%' }} 
                                    />
                                  </div> */}
                                  </div>
                                </th>
                                <th style={{minWidth:'65px', maxWidth:'65px'}}>Action</th>
                            </tr>
                            </thead>
                            <tbody>
                            {currentData.map((item:any, index:any) => (
                                <React.Fragment key={item.Id}>
                                <tr >
                                    <td style={{minWidth:'55px', textAlign:'center', maxWidth:'55px'}} >
                                  <span style={{marginLeft:'45px'}} className='indexdesign'>
                                    {(currentPage - 1) * itemsPerPage + index + 1}
                                    </span>
                                    </td>
                                    <td >
                                    {item.Title || ''}
                                    </td>
                                    <td >
                                    {item.Email || ''}
                                    </td>
                                    <td style={{minWidth:'65px', maxWidth:'65px'}}>
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
                    <div className={styles.title}>  {props.selectedEntityForPermission.value} &gt; {props.selectedGropuForPermission.value}</div>
                    <div>
                        <button type='button' onClick={handleBackToTable}>
                            Back
                        </button>
                    </div>
                    <div style={{
                      width:"fit-content",
                      position:"relative",
                      marginLeft:"50px",
                      marginTop:"50px",
                      padding:"20px",
                      border:"2px solid #7fc4de",
                      borderRadius:"20px",
                      background:"#fff",

                    }}>
                        <p style={{
                            color:"Black",
                            marginBottom:"20px",
                            marginLeft:"160px"
                        }}>Add Users</p>
                        <div style={{
                            gap:"30px",
                            display:"flex"
                        }}>
                            <div  style={{
                                width:"220px"
                            }}>
                                <Select
                                    isMulti
                                    options={user}
                                    onChange={(selected: any) =>
                                    handleUsersSelect(selected)
                                    }
                                    placeholder="Select User..."
                                    noOptionsMessage={() => "No User Found..."}
                                />
                            </div>

                            <div>
                                <button type='button' onClick={handleAddUsers}>
                                    Add
                                </button>
                            </div>
                        </div>
                        
                    </div>           
                </div>
                )
              } 
              </div>
 </>
            
           
  )
}
// export const ManagePermission = (props:any) => {
//     const sp: SPFI = getSP();
//     const { useHide }: any = React.useContext(UserContext);
//     const elementRef = React.useRef<HTMLDivElement>(null);
//     const isMobile = useMediaQuery({ query: '(max-width: 768px)' });
//     const [selectedUser,setSelectedUser]=React.useState([]);
//     const [refresh,setRefresh]=React.useState(false);
//     const [activeComponent,setActiveComponent]=React.useState('');
//     const [user,setUser]=React.useState<any[]>([]);
//     const [description,setDescription]=React.useState('');
//     console.log("selectedUser",selectedUser);
//     console.log("props",props);

//     React.useEffect(()=>{
//             const fetchUserFromSelectedGroup=async()=>{
//                 try {
//                     const subsiteContext = await sp.site.openWebById(props.selectedEntityForPermission.SiteID);
//                     const usersFromSelectedGroups = await subsiteContext.web.siteGroups.getByName(`${props.selectedGropuForPermission.value}`).users();
//                     console.log("usersFromSelectedGroups",usersFromSelectedGroups);
//                     setSelectedUser(usersFromSelectedGroups);
//                   } catch (error) {
//                     console.log("error from getting the users from the groups after selecting the groups",error);
//                   }
//             }
//             fetchUserFromSelectedGroup();
//     },[refresh]);

//     const handleDeleteUser=async(userId:any,UserTitle:any)=>{
//         console.log("UserId",userId);
//         try {

//             const subsitecontext=await sp.site.openWebById(props.selectedEntityForPermission.SiteID);
//             // Get the group by name
//             const group = await subsitecontext.web.siteGroups.getByName(props.selectedGropuForPermission.value);
//             // Remove the user from the group using their userId
//             await group.users.removeById(userId);
//             console.log(`User with ID ${userId} has been removed from the group '${props.selectedGropuForPermission.value}'`);
//             onRemove(UserTitle);
//             setRefresh(!refresh);
//         } catch (error) {
//             console.error("Error removing user from group: ", error);
//         }
//     }

//     const handleToggleAddUsers=()=>{
//         setActiveComponent("AddUser");
//     }
    
//     const handleUsersSelect=(selectedUsers:any)=>{
//         console.log("selectedUsers",selectedUsers);
//         selectedUsersForPermission=selectedUsers;
//     }

//     React.useEffect(()=>{
//         const fetchUsers=async()=>{
//             const user = await sp.web.siteUsers();
//             const groups = await sp.web.siteGroups();
//             console.log("groups fetch from the site",groups);
//             const usersArray2=groups.map((user)=>(
//                 {
//                   id:String(user.Id),
//                   value: user.Title,
//                   email: user.Title,
//                   label:user.Title,
//                   loginName:user.LoginName
//                 }
//           ))
//           console.log("site users",usersArray2);
//             console.log("users fetch from the site",user);
//               const usersArray=user.map((user)=>(
//                     {
//                       id:String(user.Id),
//                       value: user.Title,
//                       email: user.Email,
//                       label:user.Title,
//                       loginName:user.LoginName
//                     }
//               ))
//               console.log("site users",usersArray);
//               setUser(usersArray);
//         }
//         fetchUsers();
        
       
//     },[])

//     // console.log("description",description);
//     React.useEffect(()=>{
//          // Add Description
//          const addDescription=()=>{
//             const result: string = props.selectedGropuForPermission.value.split("_")[1];
//             console.log("Description",result);
//             switch (result) {
//                 case 'Admin':
//                      setDescription("Full Control - Users Can Create and Update Folder,Upload , View , Share & Delete Files.");
//                      break;
//                 case 'Read':
//                     setDescription("Read - Users Can Uplaod and View Files.");
//                     break;
//                 case 'View':
//                     setDescription("View - Users Can only View Files.");
//                     break;
//                 case 'Contribute':
//                     setDescription("Contribute - Users Can Upload , View , Share Files.");
//                     break;
//                 case 'Initiator':
//                     setDescription("Initiator - Users Can Upload , View , Share Files.");
//                     break;
//                 case 'Approval':
//                     setDescription("Approval - Users Will Approve Files");
//                     break;
//                 case 'AllUsers':
//                     setDescription("AllUsers - All Users");
//                     break;
//                 default:
//                     setDescription("Unknown role.");
//             }
//         }
//         addDescription();

//     },[])
   

//     const handleAddUsers=async()=>{
//         console.log("selectedUsersForPermission",selectedUsersForPermission);
//         console.log("selectedGropuForPermission",props.selectedGropuForPermission.value);
//         console.log("selectedEntityForPermission",props.selectedEntityForPermission.value);

//         if(selectedUsersForPermission === undefined || selectedUsersForPermission.length === 0){
//           checkValidation();
//           return;
//         }
//         const subsiteContext = await sp.site.openWebById(props.selectedEntityForPermission.SiteID); 
//         selectedUsersForPermission.forEach(async(user:any)=>{
//           try {
//             const userObj = await sp.web.ensureUser(user.email);
//             console.log("userObj",userObj);
//             const users=await subsiteContext.web.siteGroups.getByName(`${props.selectedGropuForPermission.value}`).users.add(userObj.data.LoginName);
//             console.log(`${user.email} added to the group successfully.`,users);
//           } catch (error) {
//             console.error(`Failed to add ${user.email} to the group: `, error);
//           }
//         })
//         onSuccess();
//         setRefresh(!refresh);
//         setActiveComponent('');
        
//       }
    
//     const handleBackToTable=()=>{
//         setActiveComponent('');
//     }
//     const onSuccess=()=>{
//         Swal.fire(`Users Added Successsfully`,"", "success");
//     }
//     const onRemove=(UserTitle:any)=>{
//         Swal.fire(`${UserTitle} Removed Successsfully`,"", "success");
//     }
//     const checkValidation=()=>{
//         Swal.fire("Please fill out the fields!", "All fields are required");
//   }

//   return (
//   <>
//               {activeComponent === '' && (
//                     <div className={styles.argform}>
//                         <div className='row'>
//                         <div className='col-md-7'>
//                         <div className={styles.title}>{props.selectedEntityForPermission.value} &gt; {props.selectedGropuForPermission.value}
//                         <div>
//                             <span className='text-muted font-14' style={{
//                                 color:"Black"
//                             }}>{description}</span>
//                         </div>
//                         </div>
//                         </div>
//                         {/* <div style={{
//                             display:"flex"
//                         }}>
//                             <button type="button" onClick={handleToggleAddUsers}>
//                                 Add User
//                             </button>
//                         </div> */}
//                           <div className='col-md-5'>
//                         <div className='padd-right'>
//                             <a className={styles.backbuttonform}
//                                 onClick={props.onBack}
//                             >
//                                 <img
//                                 className={styles.backimg}
//                                 />
//                                 <p className={styles.Addtext}>Back</p>
//                             </a>
//                             <button type="button" className='btn btn-primary' onClick={handleToggleAddUsers}>
//                                 Add User 2
//                             </button>
//                         </div></div>
//                         </div>

//                         <div className={styles.container}>
//                         <table className={styles["event-table"]}>

//                             <thead>
//                             <tr>
//                                 <th className={styles.serialno}>S.No.</th>
//                                 <th className={styles.tabledept}>User</th>
//                                 <th  className={styles.tabledept}>Email</th>
//                                 <th className={styles.editdeleteicons}>Action</th>
//                             </tr>
//                             </thead>
//                             <tbody>
//                             {selectedUser.map((item:any, index:any) => (
//                                 <React.Fragment key={item.Id}>
//                                 <tr className={styles.tabledata}>
//                                     <td className={styles.serialno}>
//                                     &nbsp; &nbsp; {index + 1}
//                                     </td>
//                                     <td className={styles.tabledept}>
//                                     {item.Title || ''}
//                                     </td>
//                                     <td className={styles.tabledept}>
//                                     {item.Email || ''}
//                                     </td>
//                                     <td className={styles.editdeleteicons}>
//                                     <img
//                                         className={styles.deleteicon}
//                                         src={require("../assets/delete.png")}
//                                         alt="Delete"
//                                         onClick={(event)=>{
//                                             handleDeleteUser(item.Id,item.Title)
//                                         }}
//                                     />
//                                     </td>
//                                 </tr>
//                                 </React.Fragment>
//                             ))}
//                         </tbody>
//                         </table>
//                         </div>
//                     </div>
//               )}
//               {activeComponent === "AddUser" && 
//                 (
//                 <div className={styles.argform}>
//                     <div className='row'>
//                         <div className='col-md-7'>

                     
//                     <div className={styles.title}>  {props.selectedEntityForPermission.value} &gt; {props.selectedGropuForPermission.value}</div>
//                   </div>
//                   <div className='col-md-5'>
                    
                  
//                     <div className='d-flex justify-content-end padd-right'>
//                         <button type='button' className='btn backbuttonform' onClick={handleBackToTable}>
//                             Back
//                         </button>
//                     </div>
//                     </div>
//                     </div>
//                     <div style={{
                   
//                       position:"relative",
                    
//                       marginTop:"10px",
//                       padding:"20px",
//                       border:"2px solid #7fc4de",
//                       borderRadius:"10px",
//                       background:"#fff",

//                     }}>
//                         <p style={{
//                             color:"Black",
                        
                     
//                         }}>Add Users</p>
//                         <div style={{
//                             gap:"30px",
//                             display:"flex"
//                         }}>
//                             <div  style={{
//                                 width:"370px"
//                             }}>
//                                 <Select
//                                     isMulti
//                                     options={user}
//                                     onChange={(selected: any) =>
//                                     handleUsersSelect(selected)
//                                     }
//                                     placeholder="Select User..."
//                                     noOptionsMessage={() => "No User Found..."}
//                                 />
//                             </div>

//                             <div>
//                                 <button type='button' className='btn btn-primary' onClick={handleAddUsers}>
//                                     Add
//                                 </button>
//                             </div>
//                         </div>
                        
//                     </div>           
//                 </div>
//                 )
//               } 
//              </>
//   )
// }
