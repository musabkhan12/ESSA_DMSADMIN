import * as React from 'react';
import { getSP } from '../loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
// import "../../verticalSideBar/components/VerticalSidebar.scss"
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
// import "../../verticalSideBar/components/VerticalSidebar.scss"
import CreateEntity from './CreateMaster';
import { format } from '@fluentui/react';
import { faSort } from '@fortawesome/free-solid-svg-icons';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import Swal from 'sweetalert2';
const Entity = () => {

  const sp: SPFI = getSP();
  console.log(sp, 'sp');
  const { useHide }: any = React.useContext(UserContext);
  console.log('This function is called only once', useHide);
  const elementRef = React.useRef<HTMLDivElement>(null);
  const [refresh,setRefresh]=React.useState(false);
  const [entityDetails,setEntityDetails]=React.useState<any[]>([]);
  console.log("Fetched Entity",entityDetails);
 
 
  React.useEffect(()=>{
        async function fetchData(){
            console.log("Fetchin Entity");
            const entity = await sp.web.lists
            .getByTitle('MasterSiteURL')
            .items.select("SiteURL","Title","Active","Created","Description","UniqueId","Author/Title","Id","SiteID").expand("Author").orderBy("Modified", false)();
            setEntityDetails(entity);
            console.log("Fetched Entity",entity);
        }

        fetchData();
  },[refresh])

  
  // Media query to check if the screen width is less than 768px
  const isMobile = useMediaQuery({ query: '(max-width: 768px)' });

    
  
    const [showFirstDiv, setShowFirstDiv] = React.useState(true);
    const [currentEntityId, setCurrentEntityId] = React.useState<number | null>(null);
    const [currentJobTitle, setCurrentJobTitle] = React.useState('');
    const [currentIsActive, setCurrentIsActive] = React.useState('');
    const [CurrentisExternal, setCurrentIsExternal] = React.useState('');

    const handleButtonClickShow = () => {
      setShowFirstDiv(false);
      setCurrentEntityId(null);
      setCurrentJobTitle('');
      setCurrentIsActive('');
      setRefresh(!refresh)
    };
  
    const handleBackButtonClick = () => {
      // Show the first div and hide the second div when the back button is clicked.
      setShowFirstDiv(true);
      setCurrentEntityId(null);
      setCurrentJobTitle('');
      setCurrentIsActive('');
    };

    const handleEditClick=(entity:any)=>{
        console.log("entity",entity)
        setShowFirstDiv(false);
        setCurrentEntityId(entity.Id);
        setCurrentJobTitle(entity.Title);
        setCurrentIsActive(entity.Active);
    }
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
          (filters.Status === '' || 
            (item.Modified && item.Modified.toLowerCase().includes(filters.Status.toLowerCase()))) &&
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
    const filteredEntityData=applyFiltersAndSorting(entityDetails);
        // Code for pagination start
        const [currentPage, setCurrentPage] = React.useState(1);
        const itemsPerPage = 10;
        const totalPages = Math.ceil(filteredEntityData.length / itemsPerPage);
        
        const handlePageChange = (pageNumber: any) => {
          if (pageNumber > 0 && pageNumber <= totalPages) {
            setCurrentPage(pageNumber);
          }
        };
    const startIndex = (currentPage - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    const currentData = filteredEntityData.slice(startIndex, endIndex);
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

    const handleDeleteEntity=async(item:any)=>{
      console.log("Entity Item",item);
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
             
              try {
                // Delete the subsite
            const subsite=await sp.site.openWebById(item.SiteID);
            const deletedItem=await subsite.web.delete()
            console.log("deletedItem of subsite",deletedItem);
            // Delete the list by its name
             const itemDeletedList=await sp.web.lists.getByTitle(`DMS${item.Title}FileMaster`).delete();
             console.log("itemDeletedList",itemDeletedList);

            const getItemFromEntityDivisionDepartmentMppingList=await sp.web.lists.getByTitle("EntityDivisionDepartmentMappingMasterList").items.select("*","Entitylookup/Title").filter(`Entitylookup/Title eq '${item.Title}'`).expand('Entitylookup')();
            // const getItemFromEntityDivisionDepartmentMppingList = await sp.web.lists
            // .getByTitle("EntityDivisionDepartmentMappingMasterList")
            // .items.select(
            //   "Entitylookup/Title",
            //   "Entitylookup/Active",
            //   "Devisionlookup/Title",
            //   "Departmentlookup/Title",
            //   "Devisionlookup/Active",
            //   "Departmentlookup/Active",
            //   "Id",
            //   "UniqueId",
            //   "Created",
            //   "Author/Title"
            // )
            // .expand("Entitylookup", "Devisionlookup", "Departmentlookup","Author")();
            console.log("getItemFromEntityDivisionDepartmentMppingList",getItemFromEntityDivisionDepartmentMppingList);

            const getItemsFromMasterSiteUrl=await sp.web.lists.getByTitle("MasterSiteURL").items.select("*").filter(`Title eq '${item.Title}'`)();
            console.log("getItemsFromMasterSiteUrl",getItemsFromMasterSiteUrl);

            if(getItemsFromMasterSiteUrl.length > 0){
              for(const item of getItemsFromMasterSiteUrl){
                try {
                  const deletedData=await sp.web.lists.getByTitle("MasterSiteURL").items.getById(item.ID).delete()
                  console.log("Item deleted from dmsfoldermaster list",deletedData);
                } catch (error) {
                  console.log("Error in deleting the item from dmsfoldermasterlist",error);
                }
              }
            }

            const getItemsFromDMSFolderMaster=await sp.web.lists.getByTitle("DMSFolderMaster").items.select("*").filter(`SiteTitle eq '${item.Title}'`)();
            console.log("getItemsFromDMSFolderMaster",getItemsFromDMSFolderMaster);

            if(getItemsFromDMSFolderMaster.length > 0){
              for(const item of getItemsFromDMSFolderMaster  ){
                try {
                  const deletedData=await sp.web.lists.getByTitle("DMSFolderMaster").items.getById(item.ID).delete()
                  console.log("Item deleted from dmsfoldermaster list",deletedData);
                } catch (error) {
                  console.log("Error in deleting the item from dmsfoldermasterlist",error);
                }
                
              }
            }

            const getItemsFromDMSFolderPrivacy=await sp.web.lists.getByTitle("DMSFolderPrivacy").items.select("*").filter(`SiteName eq '${item.Title}'`)();
            console.log("getItemsFromDMSFolderPrivacy",getItemsFromDMSFolderPrivacy);

            const getItemsFromDMSPreviewFormMaster=await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("*").filter(`SiteName eq '${item.Title}'`)();
            console.log("getItemsFromDMSPreviewFormMaster",getItemsFromDMSPreviewFormMaster);

            const getItemsFromDMSFolderPermissionMaster=await sp.web.lists.getByTitle("DMSFolderPermissionMaster").items.select("*").filter(`SiteName eq '${item.Title}'`)();
            console.log("getItemsFromDMSFolderPermissionMaster",getItemsFromDMSFolderPermissionMaster);

            setRefresh(!refresh);
              Swal.fire({
                title: "Removed!",
                text: `${item.Title} Suucessfuly Removed.`,
                icon: "success"
              });
              } catch (error) {
                console.log("Error in deleting the subsite",error);
              }
            
            }
          });
    }

  return (
<div>
{showFirstDiv ? (
        <div className={styles.argform}>
          <div className='row'>
            <div className='col-md-7 pt-0'>
            <div className='page-title fw-bold mb-1 font-20'>Site Collection - (Location)</div>
            </div>
            <div className='col-md-5'>
            <div className="padd-right1 mt-0">
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
          <div className='mt-3'>
          {/* <div className={styles.container}>
            <table className={styles["event-table"]}>

              <thead>
                <tr>
                  <th className={styles.serialno}>S.No.</th>
                  <th className={styles.tabledept}>Title</th>
                  <th  className={styles.tabledept}>Description</th>
                  <th className={styles.tabledept}>URL</th>
                  <th className={styles.tabledept}>Status</th>
                  <th className={styles.tabledept}>Created At</th>
                  <th className={styles.tabledept}>Created By</th>
                  <th className={styles.editdeleteicons}>Action</th>
                </tr>
              </thead>
              <tbody>
                {entityDetails.map((item, index) => (
                    <React.Fragment key={item.UniqueId}>
                    <tr className={styles.tabledata}>
                        <td className={styles.serialno}>
                        &nbsp; &nbsp; {index + 1}
                        </td>
                        <td className={styles.tabledept}>
                        {item.Title || 'No Title'}
                        </td>
                        <td className={styles.tabledept}>
                        {item.Description || 'No Description'}
                        </td>
                        <td className={styles.tabledept} title={item.SiteURL}>
                        {item.SiteURL || 'No URL'}
                        </td>
                        <td className={styles.tablename}>
                        {item.Active === "Yes" ? 'Active' : 'Inactive'}
                        </td>
                        <td className={styles.tabledept}>
                        {item.Created || 'No Date'}
                        </td>
                        <td className={styles.tabledept}>
                        {item.Author.Title || 'No Author'}
                        </td>
                        <td className={styles.editdeleteicons}>
                        <img
                            className={styles.editdicon}
                            src={require("../assets/edit.svg")}
                            alt="Edit"
                            onClick={() => handleEditClick(item)}
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
                  <th  style={{ borderBottomLeftRadius: '0px', minWidth: '55px', maxWidth: '55px', borderTopLeftRadius: '0px' }}>
 
                    <div className="pb-0">
                      <span>S.No.</span>&nbsp;
                      {/*<span onClick={() => handleSortChange('SNo')}>
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
                  <th className="" >
                  
                    <div className="pb-0" >
                      <span >site collection</span> &nbsp;
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
                  {/* <th  className={styles.tabledept}>Description</th> */}
                  <th  >
               
                    <div className=" pb-0" >
                      <span >Description</span> 
                  
                    </div>
                    {/* <div className="d-flex flex-column bd-highlight "> </div> */}
                  </th>
                  <th >
                  <div className="pb-0" >
                      <span >URL</span> 
                  
                    </div>
                    {/* <div className="d-flex flex-column bd-highlight "> </div> */}
                    
                    </th>
                  <th style={{minWidth: '70px', maxWidth: '70px' }}>
                  <div className=" pb-0" >
                      <span >Status</span> 
                  
                    </div>
                    {/* <div className="d-flex flex-column bd-highlight "> </div> */}
                    </th>
                  <th >
                  <div className="pb-0" >
                      <span >Created At</span> 
                  
                    </div>
                    {/* <div className="d-flex flex-column bd-highlight "> </div> */}
                    </th>
                  {/* <th className={styles.tabledept}>Created By</th> */}
                  <th  >
                  
                      <div className=" pb-0" >
                        <span >	Created By</span> &nbsp; 
                        <span 
                          onClick={() => handleSortChange('Entity')}>
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
                    
                  <div className=" pb-0" >
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
                        <td style={{ borderBottomLeftRadius: '0px', textAlign:'center', minWidth: '55px', maxWidth: '55px', borderTopLeftRadius: '0px' }}>
                          <span className='indexdesign'>
                        {(currentPage - 1) * itemsPerPage + index + 1}
                        </span>                        </td>
                        <td className="">
                        {item.Title || 'No Title'}
                        </td>
                        <td  title={item.Description}
                        >
                        {item.Description || 'No Description'}
                        </td>
                        <td  title={item.SiteURL}>
                        {item.SiteURL || 'No URL'}
                        </td>
                        <td style={{minWidth: '70px', maxWidth: '70px' }}>
                          <div className='stausbg newsta'>
                        {item.Active === "Yes" ? 'Active' : 'Inactive'}
                        </div>
                        </td>
                        <td >
                    
                        {/* {format(new Date(item.Created), 'MMM dd, yyyy') || 'No Date'} */}
                        {format((item.Created), 'MMM dd, yyyy') || 'No Date'}
                        
                        </td>
                        <td >
                        {item.Author.Title || 'No Author'}
                        </td>
                        <td style={{ borderBottomLeftRadius: '0px', minWidth: '70px', maxWidth: '70px', borderTopLeftRadius: '0px' }}>
                        <img
                            className='editicon12'
                            src={require("../assets/edit.png")}
                            alt="Edit"
                            onClick={() => handleEditClick(item)}
                        />
                         <img
                            className={styles.deleteicon}
                            src={require("../assets/del.png")}
                            alt="Delete"
                            onClick={() => handleDeleteEntity(item)}
                        />
                        </td>
                    </tr>
                    </React.Fragment>
                ))}
            </tbody>
            </table>
            {/* Pagination Controls */}
            <Pagination
                  currentPage={currentPage}
                  totalPages={totalPages}
                  handlePageChange={handlePageChange}
  
            />
            {/* <div className={styles.pagination}>
            
              {[...Array(totalPages)].map((_, pageIndex) => (
            <a
              key={pageIndex}
              href="#"
              className={currentPage === pageIndex + 1 ? styles.active : ""}
              onClick={(e) => {
                e.preventDefault();
                handlePageClick(pageIndex + 1);
              }}
            >
              {pageIndex + 1}
            </a>
          ))}
          </div> */}
          </div>
          </div>
        </div>
      ) : (
        <div className={styles.argform}>
          <div style={{marginBottom:"20px"}} className='row'>
            <div className='col-md-7'>
            <div className='page-title fw-bold mb-1 font-20'>Create Site Collection - (Location)</div>
            </div>
            <div className='col-md-5'>
             <div className='padd-right1 mt-0'>
            <div className={styles.actions}>
              <div
                className={styles.backbuttonform}
                onClick={handleBackButtonClick}
              >
                <img
                  className={styles.backimg}
                //   src={require("../assets/left.png")}
                />
                <p className={styles.Addtext}>Back</p>
              </div>
            </div>
            </div>
            </div>
          </div>
          <CreateEntity
                currentId={currentEntityId}
                currentJobTitle={currentJobTitle}
                currentIsActive={currentIsActive}
                IsExternal={CurrentisExternal}
                onCancel={() =>{ 
                  setShowFirstDiv(true)
                  setRefresh(!refresh)
                }}
                
          />
        </div>

      )}
</div>
              
              
  );
};

export default Entity;
