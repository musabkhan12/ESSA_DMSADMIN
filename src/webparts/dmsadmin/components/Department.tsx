import * as React from 'react';
import { getSP } from '../loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
// // import "../../verticalSideBar/components/VerticalSidebar.scss"
import VerticalSideBar from '../../verticalSideBar/components/VerticalSideBar';
import UserContext from '../../../GlobalContext/context';

import Provider from '../../../GlobalContext/provider';
import { useMediaQuery } from 'react-responsive';
// import context from '../../../GlobalContext/context';

// import classNames from "classnames";
import styles from './Form.module.scss'
// import { useState, useEffect, useRef , useMemo } from "react";


// import JoditEditor from "jodit-react";
// import Jodit from 'jodit-react';
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
// // import "../../verticalSideBar/components/VerticalSidebar.scss"
import CreateDepartment from './CreateDepartment';
import Swal from 'sweetalert2';
import { format } from '@fluentui/react';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faSort } from '@fortawesome/free-solid-svg-icons';

// const Department = () => {

//   const sp: SPFI = getSP();
interface DepartmentProps {
  sp: SPFI;
}

const Department: React.FC<DepartmentProps> = ({ sp }) => {
  console.log(sp, 'sp');
  const { useHide }: any = React.useContext(UserContext);
  console.log('This function is called only once', useHide);
  const elementRef = React.useRef<HTMLDivElement>(null);

  const [departmentDetails,setdivisionDetails]=React.useState<any[]>([]);
  const [refresh,setRefresh]=React.useState(false);
  console.log("Fetched Entity",departmentDetails);
 
 
  React.useEffect(()=>{
        async function fetchData(){
            console.log("Fetchin Entity");
            const department = await sp.web.lists
            .getByTitle('DepartmentMasterList')
            .items.select("Title","Active","Created","UniqueId","Author/Title","Editor/Title","Id").expand("Author","Editor")();
            setdivisionDetails(department);
            console.log("Fetched Entity",department);
        }

        fetchData();
  // },[])
  // Ritik 25-03-2026
   },[refresh])

  
  // Media query to check if the screen width is less than 768px
  const isMobile = useMediaQuery({ query: '(max-width: 768px)' });

    
  
    const [showFirstDiv, setShowFirstDiv] = React.useState(true);
    const [currentDepartmentId, setCurrentDepartmentId] = React.useState<number | null>(null);
    const [currentJobTitle, setCurrentJobTitle] = React.useState('');
    const [currentIsActive, setCurrentIsActive] = React.useState('');
    
    const dynamicHeading=currentDepartmentId? "Edit Department" : "Create Department";

    const handleButtonClickShow = () => {
      setShowFirstDiv(false);
      setCurrentDepartmentId(null);
      setCurrentJobTitle('');
      setCurrentIsActive('');
    };
  
    const handleBackButtonClick = () => {
      // Show the first div and hide the second div when the back button is clicked.
      setShowFirstDiv(true);
      setCurrentDepartmentId(null);
      setCurrentJobTitle('');
      setCurrentIsActive('');
      setRefresh(!refresh)
    };

    const handleEditClick=(department:any)=>{
      console.log("Department",department)
      setShowFirstDiv(false);
      setCurrentDepartmentId(department.Id);
      setCurrentJobTitle(department.Title);
      setCurrentIsActive(department.Active);    
  }

    // Ritik 25-03-2026 added handleFormSubmit function We dont need to refresh the page after submitting
  const handleFormSubmit = () => {
    setRefresh(prev => !prev);
    setShowFirstDiv(true);
  };
  // end here
    // New Code start 
  // Added delete funcyionality.
  const confirmDelete=(Id:any,department:any)=>{
    Swal.fire({
      title: "Are you sure?",
      text: "You won't be able to revert this!",
      icon: "warning",
      showCancelButton: true,
      confirmButtonColor: "#3085d6",
      cancelButtonColor: "#d33",
      confirmButtonText: "Yes, Removed it!"
    }).then(async(result:any) => {
      if (result.isConfirmed) {
      const deletedData=await sp.web.lists.getByTitle("DepartmentMasterList").items.getById(Id).delete();
      console.log("department deleted successfully",deletedData);
      setRefresh(!refresh);
        Swal.fire({
          title: "Removed!",
          text: `${department} Suucessfuly Removed.`,
          icon: "success"
        });
      }
    });
  }
  const handleDeleteDepartment=async(ID:any,departmentName:String)=>{
    console.log("ID",ID);
    try {
      confirmDelete(ID,departmentName)
    } catch (error) {
      console.log("Error in deleting the Department",error);
    }
  }
  // End

  // Code for filter and search start
  const [filters, setFilters] = React.useState({
    SNo: '',
    Title : '',
    // Title: '',
    CurrentUser: '',
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
        (filters.CurrentUser === '' || 
          (item.Author.Title && item.Author.Title.toLowerCase().includes(filters.CurrentUser.toLowerCase()))) &&
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
  
  const filteredDepartmentData=applyFiltersAndSorting(departmentDetails);
  // end

    // Add pagination start
    const [currentPage, setCurrentPage] = React.useState(1);
    const itemsPerPage = 10;
    const totalPages = Math.ceil(filteredDepartmentData.length / itemsPerPage);
    
    const handlePageChange = (pageNumber: any) => {
      if (pageNumber > 0 && pageNumber <= totalPages) {
        setCurrentPage(pageNumber);
      }
    };
    
    const startIndex = (currentPage - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    const currentData = filteredDepartmentData.slice(startIndex, endIndex);
  
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
    }
        
  return (
 <div>
{showFirstDiv ? (
        <div className={styles.argform}>
          <div className='row'>
            <div className='col-md-7 pt-0'>
              

            <div className='page-title fw-bold mb-1 font-20'>Department</div>
            </div>
            <div className='col-md-5'>       
              <div className='padd-right1 mt-0'>
                   <div className={styles.actions}>
              {/* <a className={styles.backbuttonform}>
                <img
                  className={styles.backimg}
                //   src={require("../assets/left.png")}
                />
                <p className={styles.Addtext}>Back</p>
              </a> */}
              <div 
                onClick={handleButtonClickShow}
                className={styles.addbuttonargform}
              >
                {/* <img
                  className={styles.addimg}
                  src={require("../assets/plus.png")}
                /> */}
                <p className={styles.Addtext}>Create New</p>
              </div>
            </div>
            </div>
            </div>
          </div>
          <div className='mt-3'>
          {/* <div className={styles.container}>
            <table className={styles["event-table"]}>

              <thead>
                <tr>
                  <th className={styles.serialno}>S.No.</th>
                  <th className={styles.tabledept}>Title</th>
                  <th className={styles.tabledept}>IsActive</th>
                  <th className={styles.tabledept}>Created At</th>
                  <th className={styles.tabledept}>Created By</th>
                  <th className={styles.tabledept}>Modified By</th>
                  <th className={styles.editdeleteicons}>Action</th>
                </tr>
              </thead>
              <tbody>
                {departmentDetails.map((item, index) => (
                    <React.Fragment key={item.UniqueId}>
                    <tr className={styles.tabledata}>
                        <td className={styles.serialno}>
                        &nbsp; &nbsp; {index + 1}
                        </td>
                        <td className={styles.tabledept}>
                        {item.Title || 'No Title'}
                        </td>
                        <td className={styles.tabledept}>
                        {item.Active === 'Yes' ? 'Active' : 'Inactive'}
                        </td>
                        <td className={styles.tabledept}>
                        {item.Created || 'No Date'}
                        </td>
                        <td className={styles.tabledept}>
                        {item.Author.Title || 'No Author'}
                        </td>
                        <td className={styles.tabledept}>
                        {item.Editor.Title || 'No Author'}
                        </td>
                        <td className={styles.editdeleteicons}>
                        <img
                            className={styles.editdicon}
                            src={require("../assets/edit.svg")}
                            alt="Edit"
                            onClick={() => handleEditClick(item)}
                        />
                        <img
                            className={styles.deleteicon}
                            src={require("../assets/delete.png")}
                            alt="Delete"
                        />
                        </td>
                    </tr>
                    </React.Fragment>
                ))}
            </tbody>
            </table>
          </div> */}
                 <div style={{padding:'15px'}} className={styles.container}>
                 <table className="mtbalenew">

              <thead>
                <tr>
                  {/* <th className={styles.serialno}>S.No.</th> */}
                  <th  style={{ borderBottomLeftRadius: '0px', minWidth: '50px', maxWidth: '50px', borderTopLeftRadius: '0px' }}>
                    
                    <div className="">
                      <span>S.No.</span>&nbsp;
                      {/* <span onClick={() => handleSortChange('SNo')}>
                        <FontAwesomeIcon icon={faSort} />
                      </span> */}
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
                  {/* <th className={styles.tabledept}>Title</th> */}
                  <th >
                   
                    <div className="" >
                      <span >Title</span> &nbsp;
                      <span className="Sorting" onClick={() => handleSortChange('Title')}>
                        <FontAwesomeIcon icon={faSort} /> 
                      </span>
                    </div>
                    {/* <div className=" bd-highlight">
                      <input 
                        type="text" 
                        placeholder="Filter by Title" 
                        onChange={(e) => handleFilterChange(e, 'Title')}
                        className='inputcss' 
                        style={{ width: '100%' }} 
                      />
                    </div> */}
                  
                  </th>
                  <th style={{minWidth: '70px', maxWidth: '70px' }}>
                  <div className="" >
                      <span >Status</span> 
                  
                    </div>
                    {/* <div className="d-flex flex-column bd-highlight "> </div> */}
                    </th>

                <th>
                  <div className="" >
                      <span >Created At</span> 
                  
                    </div>
                    {/* <div className="d-flex flex-column bd-highlight "> </div> */}
                    </th>
                  {/* <th className={styles.tabledept}>Created By</th> */}
                  <th style={{minWidth:'80px', maxWidth:'80px'}} >
                    
                      <div className="" >
                        <span >	Created By</span>  
                        <span 
                          onClick={() => handleSortChange('Title')}>
                              <FontAwesomeIcon icon={faSort} /> 
                        </span>
                      </div>
                        {/* <div className=" bd-highlight">
                          <input 
                            type="text" 
                            placeholder="Filter by User" 
                            onChange={(e) => handleFilterChange(e, 'CurrentUser')}
                            className='inputcss' 
                            style={{ width: '100%' }} />
                        </div> */}
                      
                  </th>
                  {/* <th className={styles.tabledept}>Modified By</th> */}
                  <th style={{minWidth:'80px', maxWidth:'80px'}}>
                  
                      <div className="" >
                        <span >	Modified By</span>  
                        <span 
                          onClick={() => handleSortChange('Title')}>
                              <FontAwesomeIcon icon={faSort} /> 
                        </span>
                      </div>
                        {/* <div className=" bd-highlight">
                          <input 
                            type="text" 
                            placeholder="Filter by User" 
                            onChange={(e) => handleFilterChange(e, 'Modified')}
                            className='inputcss' 
                            style={{ width: '100%' }} />
                        </div> */}
                     
                  </th>
                  <th style={{ borderBottomLeftRadius: '0px', minWidth: '70px', maxWidth: '70px', borderTopLeftRadius: '0px' }}>
                    
                    <div className="" >
                        <span >Action</span> 
                    
                      </div>
                      {/* <div className="d-flex flex-column bd-highlight "> </div> */}
                      </th>
                </tr>
              </thead>
              <tbody>
                {currentData.map((item, index) => (
                    <React.Fragment key={item.UniqueId}>
                    <tr className={styles.tabledata}>
                    <td style={{ borderBottomLeftRadius: '0px', minWidth: '50px', maxWidth: '50px', borderTopLeftRadius: '0px' }}>
                    <span className='indexdesign'>
                        {(currentPage - 1) * itemsPerPage + index + 1}
                        </span>
                        </td>
                        <td title={item.Title}>
                        {item.Title || 'No Title'}
                        </td>
                        <td style={{minWidth: '70px', maxWidth: '70px' }}>
                          <div className='stausbg newsta'>
                        {item.Active === "Yes" ? 'Active' : 'Inactive'}
                        </div>
                        </td>
                       
                        <td >
                        {/* {item.Created || 'No Date'} */}
                        {format(item.Created, 'MMM dd, yyyy') || 'No Date'}
                        </td>
                        <td style={{minWidth:'80px', maxWidth:'80px'}}>
                        {item.Author.Title || 'No Author'}
                        </td>
                        <td style={{minWidth:'80px', maxWidth:'80px'}}>
                        {item.Editor.Title || 'No Author'}
                        </td>
                        <td style={{ borderBottomLeftRadius: '0px', minWidth: '70px', maxWidth: '70px', borderTopLeftRadius: '0px' }}>
                        <img
                            className='editicon1'
                            src={require("../assets/edit.png")}
                            alt="Edit"
                            onClick={() => handleEditClick(item)}
                        />
                        <img
                            className='editicon12'
                            src={require("../assets/del.png")}
                            alt="Delete"
                            onClick={() => handleDeleteDepartment(item.Id,item.Title)}
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
        </div>
      ) : (
        <div className={styles.argform}>
      <div className='row'>
        <div className='col-md-7'>        
                    <div className='page-title fw-bold mb-1 font-20'>{dynamicHeading}</div>
                    </div>
                    <div className='col-md-5'>
                      <div className='padd-right1 mt-0'>
            <div className={styles.actions}>
              <a style={{marginBottom:"20px"}}
                className={styles.backbuttonform}
                onClick={handleBackButtonClick}
              >
                <img
                  className={styles.backimg}
                //   src={require("../assets/left.png")}
                />
                <p className={styles.Addtext}>Back</p>
              </a>
            </div>
            </div>
            </div>
          </div>
          <CreateDepartment
              sp={sp}
              currentId={currentDepartmentId}
              currentJobTitle={currentJobTitle}
              currentIsActive={currentIsActive}
              onCancel={() => setShowFirstDiv(true)} 
              onSubmit={handleFormSubmit} // Ritik 25-03-2026
          />

        </div>

      )}
 </div>
                
             
  );
};


export default Department;
