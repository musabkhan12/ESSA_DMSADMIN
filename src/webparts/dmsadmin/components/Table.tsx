// @ts-ignore
import * as React from "react";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
// import { faEdit, faTrashAlt } from '@fortawesome/free-regular-svg-icons';
// import { IMediaMasterProps } from './IMediaMasterProps';
import "bootstrap/dist/css/bootstrap.min.css";
// import "bootstrap//dist/"
import "../../../CustomCss/mainCustom.scss";
// import "../../verticalSideBar/components/VerticalSidebar.scss";
// import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";
import UserContext from "../../../GlobalContext/context";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import {  faEllipsisV, faFileExport, faSort , faExclamation , faListSquares
} from '@fortawesome/free-solid-svg-icons';
// import { useState , useEffect } from "react";
// import Provider from "../../../GlobalContext/provider";
import { useMediaQuery } from "react-responsive";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/sites"
import "@pnp/sp/presets/all"
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";

// import "../../verticalSideBar/components/VerticalSidebar.scss";
import "./dmscss";
import { useState , useRef , useEffect} from "react";
import "./MediaMaster.module.scss"
import "./mediamaster.scss"
import "./CustomTable.scss"

import Swal from 'sweetalert2';
import * as XLSX from 'xlsx';
// import moment from 'moment';
// import { title } from "process";
let currentsiteID = ""
interface CreateFolderProps {
  Currentbuttonclick: { [key: string]: string };
  onReturnToMain: () => void;
}

const Table: React.FC<CreateFolderProps> = ({Currentbuttonclick , onReturnToMain }) => {
  console.log(Currentbuttonclick , "Currentbuttonclick")
  const sp: SPFI = getSP();
  console.log(sp, "sp");

  const { useHide }: any = React.useContext(UserContext);
  const elementRef = React.useRef<HTMLDivElement>(null);
  const [isSidebarOpen, setIsSidebarOpen] = React.useState(true);

  interface IListItem {
    ID: number;
    Title: string;
    // Add other properties as needed from your list
  }
   const [mediaData, setmediaData] = useState<IListItem[]>([]); 
  
  // const [mediaData, setmediaData] = React.useState([]);
  const isMobile = useMediaQuery({ query: '(max-width: 768px)' });
  const [filters, setFilters] = React.useState({
    SNo: '',
    Title : '',
    FileName: '',
    CurrentUser: '',
    Modified: '',
    Status: '',

    SubmittedDate: ''
  });
  const [sortConfig, setSortConfig] = React.useState({ key: '', direction: 'ascending' });

  React.useEffect(() => {
    console.log("This function is called only once", useHide);

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


  React.useEffect(() => {
    console.log("This function is called only once", useHide);

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
/////////////////// DMS Code start / ////////////////////////////////////
  
  console.log("This is current side ID",currentsiteID)
  const currentUserEmailRef = useRef('');
  useEffect(() => {
     getcurrentuseremail()
     ApiCall()
}, []);
 const getcurrentuseremail = async()=>{
  const userdata = await sp.web.currentUser();
  currentUserEmailRef.current = userdata.Email;
  console.log(currentUserEmailRef.current, "currentuser")
 }

 const ApiCall = async () => {
  if (Currentbuttonclick.buttonclickis === 'Myrequest') {
    console.log(Currentbuttonclick , "Currentbuttonclick")
    console.log(typeof Currentbuttonclick , "typeof Currentbuttonclick")
    // alert(Currentbuttonclick.buttonclickis)
    // do something
        // Fetch the list of active lists
        const FilesItems = await sp.web.lists
        .getByTitle("MasterSiteURL")
        .items.select("Title", "SiteID", "FileMasterList", "Active")
        .filter(`Active eq 'Yes'`)();
    
      console.log("Active Sites List Names", FilesItems);
    
      FilesItems.forEach(async (fileItem) => {
  
          const filesData = await sp.web.lists
            .getByTitle(`${fileItem.FileMasterList}`)
            .items.select("ID" , "FileName", "FileUID", "FileSize", "FileVersion" ,"Status" , "SiteID" , "CurrentUser"
              , "Modified"
            )
            .filter(
              `CurrentUser eq '${currentUserEmailRef.current}'`
            )();
    
          console.log(`Files of Current user ${fileItem.FileMasterList}`, filesData);
          setmediaData((prevMediaData) => [...prevMediaData, ...filesData]); 
          // setmediaData(filesData)
          console.log(filesData , "filesData")
  });
  } else {
    console.log(Currentbuttonclick , "Currentbuttonclick")
    console.log(typeof Currentbuttonclick , "typeof Currentbuttonclick")
    // alert(Currentbuttonclick)
    // do something else
        // Fetch the list of active lists
        const FilesItems = await sp.web.lists
        .getByTitle("MasterSiteURL")
        .items.select("Title", "SiteID", "FileMasterList", "Active")
        .filter(`Active eq 'Yes'`)();
    
      console.log("Active Sites List Names", FilesItems);
    
      FilesItems.forEach(async (fileItem) => {
  
          const filesData = await sp.web.lists
            .getByTitle(`${fileItem.FileMasterList}`)
            .items.select("ID" , "FileName", "FileUID", "FileSize", "FileVersion" ,"Status" , "SiteID" , "CurrentUser"
              , "Modified" , "IsFavourite"
            )
            .filter(
              `IsFavourite eq 1 and CurrentUser eq '${currentUserEmailRef.current}'`
            )();
    
          console.log(`Files of Current user ${fileItem.FileMasterList}`, filesData);
          setmediaData((prevMediaData) => [...prevMediaData, ...filesData]); 
          // setmediaData(filesData)
          console.log(filesData , "filesData")
  });
  }


console.log(mediaData , "statelistData")
 }
const headers = [
  { label: 'S.No.', key: 'ID', style: { width: '5%' } },
  { label: 'Title', key: 'Title', style: { width: '20%' } },
  { label: 'Image', key: 'mediaandNewsBannerImage', type: 'image', style: { width: '10%' } },
  { label: 'Description', key: 'Description', style: { width: '50%' } },
  { label: 'Date', key: 'SubmittedDate', style: { width: '15%' } },
  { label: 'Action', key: 'Action', style: { width: '15%' } },

];
const handleFilterChange = (e: React.ChangeEvent<HTMLInputElement>, field: string) => {
  setFilters({
    ...filters,
    [field]: e.target.value,
  });
  console.log(filters , "filters filters")
};
console.log(filters , "filters filters")
const handleSortChange = (key: string) => {
  let direction = 'ascending';
  if (sortConfig.key === key && sortConfig.direction === 'ascending') {
    direction = 'descending';
  }
  setSortConfig({ key, direction });
};
const applyFiltersAndSorting = (data: any[]) => {
  debugger;
  // Filter data
  const filteredData = data.filter((item, index) => {
    return (
      (filters.Title === '' || item.FileName.toLowerCase().indexOf(filters.Title.toLowerCase()) !== -1) &&
       (filters.Title === '' || item.FileName.toLowerCase().includes(filters.Title.toLowerCase())) &&
       (filters.CurrentUser === '' || item.CurrentUser.toLowerCase().includes(filters.CurrentUser.toLowerCase())) &&
      (filters.Status === '' || item.Modified.toLowerCase().includes(filters?.Status?.toLowerCase())) &&
      (filters.SubmittedDate === '' || item.Status.toLowerCase().includes(filters.SubmittedDate.toLowerCase()))
    );
  });

  // Natural sort function for alphanumeric values
  const naturalSort = (a:any, b:any) => {
    return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
  };

  // Sort data
  const sortedData = filteredData.sort((a, b) => {
    if (sortConfig.key === 'SNo') {
      // Sort by index
      const aIndex = data.indexOf(a);
      const bIndex = data.indexOf(b);
      return sortConfig.direction === 'ascending' ? aIndex - bIndex : bIndex - aIndex;
    } else if (sortConfig.key) {
      // Sort by other keys
      const aValue = a[sortConfig.key] ? a[sortConfig.key].toLowerCase() : '';
      const bValue = b[sortConfig.key] ? b[sortConfig.key].toLowerCase() : '';

      return sortConfig.direction === 'ascending' ? naturalSort(aValue, bValue) : naturalSort(bValue, aValue);
    }
    return 0;
  });

  return sortedData;
};

const filteredAnnouncementData = applyFiltersAndSorting(mediaData);

const [currentPage, setCurrentPage] = React.useState(1);
const itemsPerPage = 10;
const totalPages = Math.ceil(filteredAnnouncementData.length / itemsPerPage);

const handlePageChange = (pageNumber: any) => {
  if (pageNumber > 0 && pageNumber <= totalPages) {
    setCurrentPage(pageNumber);
  }
};

const startIndex = (currentPage - 1) * itemsPerPage;
const endIndex = startIndex + itemsPerPage;
const currentData = filteredAnnouncementData.slice(startIndex, endIndex);


//#region Download exl file 
const handleExportClick = () => {
  console.log(currentData,'currentData');
  
  const exportData = currentData.map((item, index) => ({

    'S.No.': startIndex + index + 1,
    'FileName': item.FileName,
    'SubmittedBy': item.CurrentUser,
    'Modified': item.Modified,
    'Status': item.Status,

  }));

  exportToExcel(exportData, 'MeadiaGallery');
};
const exportToExcel = (data: any[], fileName: string) => {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  XLSX.writeFile(workbook, `${fileName}.xlsx`);
};

const [isOpen, setIsOpen] = React.useState(false);
const toggleDropdown = () => {
  setIsOpen(!isOpen);
};
const Editmedia = (id: any) => {
  debugger
  //  setUseId(id)

  // window.location.href = `${siteUrl}/SitePages/MediaGalleryForm.aspx`;
}
//#endregion

//#region 
const Deletemedia = (id: any) => {
  Swal.fire({
    title: "Are you sure?",
    text: "You won't be able to revert this!",
    icon: "warning",
    showCancelButton: true,
    confirmButtonColor: "#3085d6",
    cancelButtonColor: "#d33",
    confirmButtonText: "Yes, delete it!"
  }).then((result) => {
    if (result.isConfirmed) {
      // const DeleteRes = DeletemediaAPI(sp, id)
      ApiCall()
      Swal.fire({
        title: "Deleted!",
        text: "Item has been deleted.",
        icon: "success"
      });

    }
  })
}
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
  const visiblePages = [];
  const limit = Math.min(pageLimit, totalPages);
  for (let i = 0; i < limit; i++) {
    visiblePages.push(adjustedStartPage + i);
  }


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
     {visiblePages.map((pageNumber:any) => (
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



  return (
<>
{/* <button onClick={onReturnToMain}>Back To Main Component</button> */}
<table className="mtable table-centered table-nowrap table-borderless mb-0">
                        <thead className="tablehead">
                          <tr>
                            <th className="tableheadtitle" style={{ borderBottomLeftRadius: '10px', minWidth: '50px', maxWidth: '50px', borderTopLeftRadius: '10px' }}>
 
                              <div className="d-flex pb-0"
                                >
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
                            <th className="tableheadtitle" ><div className="d-flex flex-column bd-highlight ">
                              <div className="d-flex pb-0" >
                                <span >Document Name</span>  
                                {/* <span className="Sorting" onClick={() => handleSortChange('Title')}><FontAwesomeIcon icon={faSort} /> </span> */}
                                </div>
                              {/* <div className=" bd-highlight">
                                <input type="text" placeholder="Filter by Title" onChange={(e) => handleFilterChange(e, 'Title')}
                                  className='inputcss' style={{ width: '100%' }} />
                              </div> */}
                            </div></th>
                            <th className="tableheadtitle" >
                              <div className="d-flex flex-column bd-highlight ">
                                <div className="d-flex pb-0" >
                                  <span >	Submiited By</span>  
                                  {/* <span onClick={() => handleSortChange('Entity')}><FontAwesomeIcon icon={faSort} /> </span></div>
                                <div className=" bd-highlight">
                                  <input type="text" placeholder="Filter by Entity" onChange={(e) => handleFilterChange(e, 'CurrentUser')}
                                    className='inputcss' style={{ width: '100%' }} /> */}
                                </div>
                              </div>
                            </th>
                            <th className="tableheadtitle" >
                              <div className="d-flex flex-column bd-highlight ">
                                <div className="d-flex pb-0" >
                                  <span >Modified date</span> 
                                   {/* <span onClick={() => handleSortChange('Status')}><FontAwesomeIcon icon={faSort} /> </span> */}
                                   </div>
                                {/* <div className=" bd-highlight">
                                  <input type="text" placeholder="Filter by Status" onChange={(e) => handleFilterChange(e, 'Status')}
                                    className='inputcss' style={{ width: '100%' }} />
                                </div> */}
                              </div>
                            </th>
                            <th className="tableheadtitle"  style={{ minWidth: '80px', maxWidth: '80px' }}><div className="d-flex flex-column bd-highlight ">
                              <div className="d-flex pb-0" >
                                <span >Status</span> 
                                 {/* <span onClick={() => handleSortChange('SubmittedDate')}><FontAwesomeIcon icon={faSort} /> </span> */}
                                </div>
                              {/* <div className=" bd-highlight">
                                <input type="text" placeholder="Filter by Date" onChange={(e) => handleFilterChange(e, 'SubmittedDate')}
                                  className='inputcss' style={{ width: '100%' }} />
                              </div> */}
                            </div></th>
                            <th className="tableheadtitle"  style={{ textAlign: 'center' }}>
                              <div className="d-flex flex-column bd-highlight pb-0">
                                <div className="d-flex  pb-0" >  <span >Action</span> 
                               <div className="dropdown">
                                  <FontAwesomeIcon icon={faEllipsisV} onClick={toggleDropdown} size='xl'/>
                                </div> 
                                </div>
                                <div className=" bd-highlight">   <div id="myDropdown" className={`dropdown-content ${isOpen ? 'show' : ''}`}>
                                  <div onClick={handleExportClick} className="" >
                                    <FontAwesomeIcon icon={faFileExport} />  Export
                                  </div>
                                </div></div>
 
                              </div>
                         
                            </th>
                          </tr>
                        </thead>
                        <tbody style={{ maxHeight: '5000px' }}>
                          {currentData.length > 0 ? currentData.map((item, index) => {
                            console.log(item , "item >>>>>>>>")
                            // const ImageUrl = item.Image == undefined || item.Image == null ? "" : JSON.parse(item.Image);
                            return (
 
                              <tr key={index} className="tablerow">
                                <td style={{ minWidth: '50px', maxWidth: '50px' }}>{index + 1}</td>
                                <td>{item.FileName}</td>
                                <td>{item.CurrentUser}</td>
                                <td>{item.Modified}</td>
                                <td >{item.Status} </td>
                                 <td style={{ minWidth: '50px', maxWidth: '50px' }} className="ng-binding">
                                  <div className="d-flex  pb-2" style={{ justifyContent: 'space-around' }}>
                                    <span >
                                       {/* <a className="action-icon text-primary" onClick={() => Editmedia(item.ID)}>
                                  <FontAwesomeIcon icon={faEdit} /> 
                                    </a> */}
                                    
                                    </span>  <span >
                                      <a className="action-icon text-danger" onClick={() => Deletemedia(item.ID)}>
                                           <FontAwesomeIcon className="Exclamationamark" icon={faExclamation} /> 
                                       {/* <FontAwesomeIcon style={{color: "black"}} icon={faListSquares}/> */}
                                      </a></span></div>
                                     
                                </td> 
                              </tr>
                            )
                          }) : ""
                          }
                        </tbody>
                      </table>
                  <Pagination
                  currentPage={currentPage}
                  totalPages={totalPages}
                  handlePageChange={handlePageChange}
  
                  />
</>
         
            
  );
};



// const Table = () => {
//   return (
//     <Provider>
//       <ArgPoc />
//     </Provider>
//   );
// };

export default Table;
