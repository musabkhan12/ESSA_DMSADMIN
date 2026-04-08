import * as React from 'react';
import { getSP } from '../loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import "../../verticalSideBar/components/VerticalSidebar.scss"
import VerticalSideBar from '../../verticalSideBar/components/VerticalSideBar';
import UserContext from '../../../GlobalContext/context';
import CreateEntityMapping from './CreateEntityMapping';
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
import "../../verticalSideBar/components/VerticalSidebar.scss"
// import EntityMapping from '../components/CreateMapping';
import Swal from 'sweetalert2';
import { format } from 'date-fns';

import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import {  faEllipsisV, faFileExport, faSort , faExclamation , faListSquares
} from '@fortawesome/free-solid-svg-icons';
//  const EntityMapping = () => {

//   const sp: SPFI = getSP();
interface EntityMappingProps {
  sp: SPFI;
}

const EntityMapping: React.FC<EntityMappingProps> = ({ sp }) => {
  console.log(sp, 'sp');
  const { useHide }: any = React.useContext(UserContext);
  console.log('This function is called only once', useHide);
  const elementRef = React.useRef<HTMLDivElement>(null);

  const [entityDetails,setEntityDetails]=React.useState<any[]>([]);
  console.log("Fetched Entity",entityDetails);
  // console.log("Devision Title",entityDetails[0].Devisionlookup.Title);
 
 // New code start
const [refresh,setRefresh]=React.useState(false);
// end
  React.useEffect(()=>{
        async function fetchData(){
            console.log("Fetching Entity");
            // const entity = await sp.web.lists
            // .getByTitle('EntityDivisionDepartmentMappingMasterList')
            // .items.select("Entitylookup/Title","Entitylookup/Active","Departmentlookup/Active","Devisionlookup/Devision","Devisionlookup/Active","Id").expand("Entitylookup","Departmentlookup","Devisionlookup")();
            // const entity = await sp.web.lists.getByTitle('EntityDivisionDepartmentMappingMasterList').items
            // .select("*").expand("Entitylookup", "Departmentlookup", "Devisionlookup")();
            const entity = await sp.web.lists
          .getByTitle("EntityDivisionDepartmentMappingMasterList")
          .items.select(
            "Entitylookup/Title",
            "Entitylookup/Active",
            "Devisionlookup/Title",
            "Departmentlookup/Title",
            "Devisionlookup/Active",
            "Departmentlookup/Active",
            "Id",
            "UniqueId",
            "Created",
            "Author/Title"
          )
          .expand("Entitylookup", "Devisionlookup", "Departmentlookup","Author").filter(`Devisionlookup/Title ne ${null}`).orderBy("Modified", false)();
          console.log("entity",entity);
            setEntityDetails(entity);
            // console.log("Fetched Entity",entity[0].Devisionlookup.Title);
        }

        fetchData();
  },[refresh])

  
  // Media query to check if the screen width is less than 768px
  const isMobile = useMediaQuery({ query: '(max-width: 768px)' });

    
  
    const [showFirstDiv, setShowFirstDiv] = React.useState(true);

    const [currentEntityId, setCurrentEntityId] = React.useState<number | null>(null);
    const [currentEntity, setCurrentEntity] = React.useState('');
    const [currentDevision, setCurrentDevision] = React.useState('');
    const [currentDepartment, setCurrentDepartment] = React.useState('');
  
    const handleButtonClickShow = () => {
      setShowFirstDiv(false);
      setCurrentEntityId(null);
      setCurrentEntity('');
      setCurrentDevision('');
      setCurrentDepartment('');
      setRefresh(!refresh)
    };
  
    const handleBackButtonClick = () => {
      // Show the first div and hide the second div when the back button is clicked.
      setShowFirstDiv(true);
      setCurrentEntityId(null);
      setCurrentEntity('');
      setCurrentDevision('');
      setCurrentDepartment('');
    };


    const handleEditClick=(item:any)=>{
      console.log("entity",item)
      setShowFirstDiv(false);
      setCurrentEntityId(item.Id);
      setCurrentEntity(item.Entitylookup?.Title);
      setCurrentDevision(item.Devisionlookup?.Title);
      setCurrentDepartment(item.Departmentlookup?.Title);
  }


   // Code for filter and search start
   const [filters, setFilters] = React.useState({
    SNo: '',
    Title : '',
    // Title: '',
    CurrentUser: '',
    Division: '',
    Department: '',

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
          (item.Entitylookup.Title && item.Entitylookup.Title.toLowerCase().includes(filters.Title.toLowerCase()))) &&
        (filters.CurrentUser === '' || 
          (item.Author.Title && item.Author.Title.toLowerCase().includes(filters.CurrentUser.toLowerCase()))) &&
        (filters.Division === '' || 
          (item.Devisionlookup.Title && item.Devisionlookup.Title.toLowerCase().includes(filters.Division.toLowerCase()))) &&
        (filters.Department === '' || 
          (item.Departmentlookup?.Title && item.Departmentlookup?.Title.toLowerCase().includes(filters.Department.toLowerCase())))
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
  
  const filteredMappedData=applyFiltersAndSorting(entityDetails);
  // end

      // Add pagination start
      const [currentPage, setCurrentPage] = React.useState(1);
      const itemsPerPage = 10;
      const totalPages = Math.ceil(filteredMappedData.length / itemsPerPage);
      
      const handlePageChange = (pageNumber: any) => {
        if (pageNumber > 0 && pageNumber <= totalPages) {
          setCurrentPage(pageNumber);
        }
      };
      
      const startIndex = (currentPage - 1) * itemsPerPage;
      const endIndex = startIndex + itemsPerPage;
      const currentData = filteredMappedData.slice(startIndex, endIndex);
    
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
                {showFirstDiv ? (
        // <div className={styles.argform}>
        //   <header>
        //     <div className={styles.title}>Mapped Entity</div>
        //     <div className={styles.actions}>
        //       {/* <a className={styles.backbuttonform}>
        //         <img
        //           className={styles.backimg}
        //         //   src={require("../assets/left.png")}
        //         />
        //         <p className={styles.Addtext}>Back</p>
        //       </a> */}
        //       <a
        //         onClick={handleButtonClickShow}
        //         className={styles.addbuttonargform}
        //       >
        //         <img
        //           className={styles.addimg}
        //           src={require("../assets/plus.png")}
        //         />
        //         <p className={styles.Addtext}>Create New</p>
        //       </a>
        //     </div>
        //   </header>
        //   <div className={styles.container}>
        //     <table className={styles["event-table"]}>

        //       <thead>
        //         <tr>
        //           <th className={styles.serialno}>S.No.</th>
        //           <th className={styles.tabledept}>Entity</th>
        //           <th  className={styles.tabledept}>Devision</th>
        //           <th className={styles.tabledept}>Department</th>
        //           <th className={styles.tabledept}>IsActive</th>
        //           <th className={styles.tabledept}>Created At</th>
        //           <th className={styles.tabledept}>Created By</th>
        //           <th className={styles.editdeleteicons}>Action</th>
        //         </tr>
        //       </thead>
        //       <tbody>
        //         {entityDetails.map((item, index) => (
        //             <React.Fragment key={item.UniqueId}>
        //             <tr className={styles.tabledata}>
        //                 <td className={styles.serialno}>
        //                 &nbsp; &nbsp; {index + 1}
        //                 </td>
        //                 <td className={styles.tabledept}>
        //                 {item.Entitylookup?.Title || 'No Title'}
        //                 </td>
        //                 <td className={styles.tabledept}>
        //                 {item.Devisionlookup?.Title || ''}
        //                 </td>
        //                 <td className={styles.tabledept} title={item.SiteURL}>
        //                 {item.Departmentlookup?.Title || ''}
        //                 </td>
        //                 <td className={styles.tablename}>
        //                 {item.Entitylookup?.Active === "Yes" ? 'Active' : 'Inactive'}
        //                 </td>
        //                 <td className={styles.tabledept}>
        //                 {item.Created || 'No Date'}
        //                 </td>
        //                 <td className={styles.tabledept}>
        //                 {item.Author?.Title || 'No Author'}
        //                 </td>
        //                 <td className={styles.editdeleteicons}>
        //                 <img
        //                     className={styles.editdicon}
        //                     src={require("../assets/edit.svg")}
        //                     alt="Edit"
        //                     onClick={() => handleEditClick(item)}
        //                 />
        //                 {/* <img
        //                     className={styles.deleteicon}
        //                     src={require("../assets/delete.png")}
        //                     alt="Delete"
        //                 /> */}
        //                 </td>
        //             </tr>
        //             </React.Fragment>
        //         ))}
        //     </tbody>
        //     </table>
        //     <div className={styles.pagination}>
        //       <a href="#">1</a>
        //       <a href="#">2</a>
        //       <a href="#">3</a>
        //       <a href="#">4</a>
        //       <a href="#">5</a>
        //     </div>
        //   </div>
        // </div>
        <div className={styles.argform}>
                <div className='row'>
                  <div className='col-sm-7 pt-0'>
          <div className='page-title fw-bold mb-1 font-20'>Mapped Entity</div>
           </div>
           <div className='col-sm-5 pt-0'>
            <div className='padd-right1 mt-0'>
          <div className={styles.actions}>
           
            <div
              onClick={handleButtonClickShow}
              className={styles.addbuttonargform}
            >
            
              <p className={styles.Addtext}>Create New</p>
            </div>
          </div>
          </div>
          </div>
        </div>
        <div style={{padding:'15px', marginTop:'20px'}} className={styles.container}>
                 <table className="mtbalenew">

            <thead>
              <tr>
                {/* <th className={styles.serialno}>S.No.</th> */}
                <th  style={{ borderBottomLeftRadius: '0px', minWidth: '50px', maxWidth: '50px', borderTopLeftRadius: '0px' }}>
                    
                    <div className="">
                      <span>S.No.</span>
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
                {/* <th className={styles.tabledept}>Entity</th> */}
                <th >
                    
                    <div className="" >
                      <span >Entity</span> &nbsp;
                      <span className="Sorting" onClick={() => handleSortChange('Title')}>
                        <FontAwesomeIcon icon={faSort} /> 
                      </span>
                    </div>
                    {/* <div className=" bd-highlight">
                      <input 
                        type="text" 
                        placeholder="Filter by Entity" 
                        onChange={(e) => handleFilterChange(e, 'Title')}
                        className='inputcss' 
                        style={{ width: '100%' }} 
                      />
                    </div> */}
                    
                </th>
                {/* <th  className={styles.tabledept}>Devision</th> */}
                <th>
                   
                    <div className="" >
                      <span >Division</span> &nbsp;
                      <span className="Sorting" onClick={() => handleSortChange('Title')}>
                        <FontAwesomeIcon icon={faSort} /> 
                      </span>
                    </div>
                    {/* <div className=" bd-highlight">
                      <input 
                        type="text" 
                        placeholder="Filter by Division" 
                        onChange={(e) => handleFilterChange(e, 'Division')}
                        className='inputcss' 
                        style={{ width: '100%' }} 
                      />
                    </div> */}
                 
                </th>
                {/* <th className={styles.tabledept}>Department</th> */}
                <th>
                  
                    <div className="" >
                      <span >Department</span> &nbsp; 
                      <span className="Sorting" onClick={() => handleSortChange('Title')}>
                        <FontAwesomeIcon icon={faSort} /> 
                      </span>
                    </div>
                    {/* <div className=" bd-highlight">
                      <input 
                        type="text" 
                        placeholder="Filter by Department" 
                        onChange={(e) => handleFilterChange(e, 'Department')}
                        className='inputcss' 
                        style={{ width: '100%' }} 
                      />
                    </div> */}
                    
                </th>
                <th style={{minWidth: '70px', maxWidth: '70px' }}>
                  <div className="" >
                      <span >IsActive</span> 
                  
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
                <th>
                  
                      <div className="" >
                        <span >	Created By</span>  &nbsp;
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
                  <tr>
                  <td style={{ borderBottomLeftRadius: '0px', minWidth: '50px', maxWidth: '50px', borderTopLeftRadius: '0px' }}>
                  <span className='indexdesign'>
                      {(currentPage - 1) * itemsPerPage + index + 1}
                      </span>
                      </td>
                      <td>
                      {item.Entitylookup?.Title || 'No Title'}
                      </td>
                      <td>
                      {item.Devisionlookup?.Title || ''}
                      </td>
                      <td  title={item.SiteURL}>
                      {item.Departmentlookup?.Title || ''}
                      </td>
                      <td style={{minWidth: '70px', maxWidth: '70px' }}>
                          <div className='stausbg newsta'>
                          {item.Entitylookup?.Active === "Yes" ? 'Active' : 'Inactive'}
                        </div>
                        </td>
                     
                      <td >
                      {/* {item.Created || 'No Date'} */}
                      {format(new Date(item.Created), 'MMM dd, yyyy') || 'No Date'}
                      </td>
                      <td>
                      {item.Author?.Title || 'No Author'}
                      </td>
                      <td style={{ borderBottomLeftRadius: '0px', minWidth: '70px', maxWidth: '70px', borderTopLeftRadius: '0px' }}>
                        <img
                          className='editicon1'
                          src={require("../assets/edit.png")}
                          alt="Edit"
                          onClick={() => handleEditClick(item)}
                      />
                      {/* <img
                          className={styles.deleteicon}
                          src={require("../assets/delete.png")}
                          alt="Delete"
                      /> */}
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
          {/* <div className={styles.pagination}>
            <a href="#">1</a>
            <a href="#">2</a>
            <a href="#">3</a>
            <a href="#">4</a>
            <a href="#">5</a>
          </div> */}
        </div>
      </div>
      ) : (
        <div className={styles.argform}>
          <div className='row' style={{marginBottom:"20px"}}>
          <div className="col-md-7">
            <div className="page-title fw-bold mb-1 font-20">Create Mapping</div>
            </div>
            <div className="col-md-5">
              <div className='padd-right1 mt-0'>
              <div className={styles.actions}>
              <a style={{marginBottom:'20px'}}
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
          {/* <EntityMapping
            currentId={currentEntityId}
            currentEntity={currentEntity}
            currentDevision={currentDevision}
            currentDepartment={currentDepartment}
            // onCancel={() => setShowFirstDiv(true)}
            onCancel={() =>{ 
              setShowFirstDiv(true)
              setRefresh(!refresh)
            }}
          /> */}
           { <CreateEntityMapping
            sp={sp}
            currentId={currentEntityId}
            currentEntity={currentEntity}
            currentDevision={currentDevision}
            currentDepartment={currentDepartment}
            onCancel={() =>{ 
              setShowFirstDiv(true)
              setRefresh(!refresh)
            }}
          /> }
        </div>
        </div>
      )}
                </div>
</>
              
              
  );
};

export default EntityMapping;

