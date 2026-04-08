import * as React from 'react';
import { getSP  , getGraphClient} from '../loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import Select from 'react-select'
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import VerticalSideBar from '../../verticalSideBar/components/VerticalSideBar';
import UserContext from '../../../GlobalContext/context';
import Provider from '../../../GlobalContext/provider';
import { useMediaQuery } from 'react-responsive';
import styles from './Form.module.scss'
import { format } from '@fluentui/react';
import { faSort } from '@fortawesome/free-solid-svg-icons';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import Swal from 'sweetalert2';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import "@pnp/sp/site-groups"
import "@pnp/sp/folders"; 
import "@pnp/sp/webs"; 
import "./CreateFoldercss";
let currentUserEmail:any = '';


interface Sitemainprops {
  context: WebPartContext;
}
let folderpath:any = '';
let DocumentLibraryName:any = '';
const Site : React.FC<Sitemainprops> = ({ context }) => {
  console.log("context of site" , context);


  async function fetchGroups() {
    const gp = await getGraphClient(context);
      debugger
    console.log("graph client" , gp)
    console.log("fetching groups");
  const groups = await gp.api("/groups")
            .filter("securityEnabled eq true") // only security groups
            .top(999) // max 999 per request (pagination required for more)
            .get();

        console.log(groups.value); 
        debugger
        return groups.value;
}
  const sp: SPFI = getSP();
  const { useHide }: any = React.useContext(UserContext);
  const elementRef = React.useRef<HTMLDivElement>(null);
  const [refresh, setRefresh] = React.useState(false);
  
  // State for active sites from MasterSiteURL
  const [activeSites, setActiveSites] = React.useState<any[]>([]);
  const [activeSitesdepartment, setActiveSitesDepartment] = React.useState<any[]>([]);
  
  // State for DMSFolderMaster data where IsSite="Yes"
  const [dmsFolderData, setDmsFolderData] = React.useState<any[]>([]);
  
  // Form state
  const [showFirstDiv, setShowFirstDiv] = React.useState(true);
  const [selectedSite, setSelectedSite] = React.useState('');
  const [selectedSitedepartment, setSelectedSitedepartment] = React.useState('');
  const [siteName, setSiteName] = React.useState('');
  const [subsiteName, setSubSiteName] = React.useState('');
  const [viewMode, setViewMode] = React.useState("list"); 
const [DocumentLibraryName, setDocumentLibraryName] = React.useState('');
const [folderpath, setFolderpath] = React.useState('');
const [siteID , setSiteID] = React.useState('');
  // Media query to check if the screen width is less than 768px
  const isMobile = useMediaQuery({ query: '(max-width: 768px)' });
 React.useEffect(() => {
  const selectedSitedepartment2 = activeSitesdepartment.find((site) => site.DocumentLibraryName === selectedSitedepartment);
  if (selectedSitedepartment2) {
    setDocumentLibraryName(selectedSitedepartment2.DocumentLibraryName);
    console.log("folder path to add folder" , selectedSitedepartment2.FolderPath);
    setFolderpath(selectedSitedepartment2.FolderPath);
  }
}, [selectedSitedepartment]);
  // Fetch active sites from MasterSiteURL
  React.useEffect(() => {
    sp.web.currentUser().then(user => {
  const userEmail = user.Email;
  console.log(userEmail);
//   alert(userEmail);
  currentUserEmail = userEmail;
});
fetchGroups();
    async function fetchActiveSites() {
      try {
        console.log("Fetching Active Sites");
        const sites = await sp.web.lists
          .getByTitle('MasterSiteURL')
          .items.select("SiteURL", "Title", "Active", "Created", "Description", "UniqueId", "Author/Title", "Id", "SiteID")
          .expand("Author")
          .filter("Active eq 'Yes'")
          .orderBy("Modified", false)();
       
        setActiveSites(sites);
   
        console.log("Fetched Active Sites", sites);
      } catch (error) {
        console.error("Error fetching active sites:", error);
      }
    }

    fetchActiveSites();
  }, [refresh]);
  
  const getdataofdepartment = async (e: any) => {
   setSelectedSite(e.target.value)
   
 const sitedepartments = await sp.web.lists
        .getByTitle('DMSFolderMaster').items.select("SiteTitle" , "DocumentLibraryName" , "IsLibrary" , "IsActive" ,"IsSite" , "FolderPath")
        .filter("IsLibrary eq 1 and IsActive eq 1 and IsSite eq 'Yes' and SiteTitle eq '" + e.target.value + "'")()
     setActiveSitesDepartment(sitedepartments);
  }
  // Fetch DMSFolderMaster data where IsSite="Yes"
  React.useEffect(() => {
    async function fetchDmsFolderData() {
      try {
        console.log("Fetching DMSFolderMaster data");
        const data = await sp.web.lists
          .getByTitle('DMSFolderMaster')
          .items.select("Title", "IsSite", "Created", "Author/Title", "Id" ,"DocumentLibraryName" , "SiteTitle" , "ParentFolderId" , "FolderName" , "IsSubsite") 
          .expand("Author")
       .filter("(IsSite eq 'Yes' or IsSubsite eq 'Yes') and IsActive eq 1")
          .orderBy("SiteTitle", true)();
        
        setDmsFolderData(data);
        console.log("Fetched DMSFolderMaster data", data);
      } catch (error) {
        console.error("Error fetching DMSFolderMaster data:", error);
      }
    }

    fetchDmsFolderData();
  }, [refresh]);

  const handleButtonClickShow = () => {
    setShowFirstDiv(false);
  };

  const handleBackButtonClick = () => {
    setViewMode("list")
    setShowFirstDiv(true);
    setSelectedSite('');
    setSiteName('');
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    //  alert("Site creation logic goes here" + selectedSite + " Site Name: " + siteName);
    // Validate form
    if (!selectedSite || !siteName) {
      Swal.fire({
        title: "Error!",
        text: "Please select a site and enter a site name",
        icon: "error"
      });
      return;
    }
    
    try {
// Example payload for site creation, using selectedSite and siteName from state
const payloadForFolderMaster: any = {
  SiteTitle: selectedSite,
  IsProcessRelated: 'No'
};
        // alert("DocumentLibraryName" + siteName + "Folderpath" + `https://officeindia.sharepoint.com/sites/Intranetdemos/${selectedSite}/${siteName}`);
        (payloadForFolderMaster as any).DocumentLibraryName=siteName;
          //  (payloadForFolderMaster as any).FolderPath=`/sites/IntranetUAT/${OthProps.Entity}/${folderName}`;
          //  (payloadForFolderMaster as any).FolderPath=`/sites/AlRostmanispfx2/${OthProps.Entity}/${folderName}`;
           (payloadForFolderMaster as any).FolderPath=`/sites/Intranetdemos/${selectedSite}/${siteName}`;
          //  (payloadForFolderMaster as any).FolderPath=`/sites/AlRostmani/${OthProps.Entity}/${folderName}`;
          (payloadForFolderMaster as any).IsLibrary=true;
          (payloadForFolderMaster as any).IsActive=false;
     
            (payloadForFolderMaster as any).IsPrivate=false;
          
          
            (payloadForFolderMaster as any).IsFolderDeligation=false;
          
       
            (payloadForFolderMaster as any).External=false;
            
                (payloadForFolderMaster as any).CurrentUser= currentUserEmail;
                (payloadForFolderMaster as any).IsSite='Yes';
            (payloadForFolderMaster as any).SiteURL=`https://officeindia.sharepoint.com/sites/Intranetdemos/${selectedSite}/${siteName}`;
              const addedItem = await sp.web.lists.getByTitle("DMSFolderMaster").items.add(payloadForFolderMaster);
        console.log("Item added successfully in the DMSFolderMaster", addedItem);
// Your logic to create the site goes here

 const payloadForFolderMaster2: any = {
  SiteName: selectedSite,
  DocumentLibraryName : siteName,
  IsDocumentLibrary : true,
  IsPrivate : false,
};
       const addedItem2 = await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.add(payloadForFolderMaster2);
console.log("Creating site with:", { selectedSite, siteName, payloadForFolderMaster });

// After successful creation
Swal.fire({
  title: "Success!",
  text: "Site created successfully",
  icon: "success"
});

// Reset form and go back to list view
setSelectedSite('');
setSiteName('');
setShowFirstDiv(true);
setRefresh(!refresh);
    } catch (error) {
      console.error("Error creating site:", error);
      Swal.fire({
        title: "Error!",
        text: "Failed to create site",
        icon: "error"
      });
    }
  };

  const handleSubmit2 = async (e: React.FormEvent) => {
    e.preventDefault();
    if(!selectedSite || !selectedSitedepartment ) {
      Swal.fire({
        title: "Error!",
        text: "Please select a site, department and enter a subsite name",
        icon: "error"
      });
    }
    try {
        
              const {web} =await sp.site.openWebById(siteID);
             if(!web){
              // alert("web not found");
             }
              // alert("web"+ web);
              // alert("siteID"+ siteID);
              // alert("folderpath"+ folderpath);
              // alert("subsiteName"+ subsiteName);
              console.log("subsiteName which is folder name  "+ subsiteName)
              console.log("folder whihc i am creating   "+ `${folderpath}/${subsiteName.trim()}`)
         
              const folderAddResult = await web.folders.addUsingPath(`${folderpath}/${subsiteName.trim()}`);
              debugger
//               if(!folderAddResult){
//                 alert("folder not created");
//               }
//               console.log("Folder created successfully -",folderAddResult);
          const payloadForFolderMaster: any = {
  SiteTitle: selectedSite,
  IsProcessRelated: 'No'
};
        // alert("DocumentLibraryName" + siteName + "Folderpath" + `https://officeindia.sharepoint.com/sites/Intranetdemos/${selectedSite}/${siteName}`);
        (payloadForFolderMaster as any).DocumentLibraryName=selectedSitedepartment;
          //  (payloadForFolderMaster as any).FolderPath=`/sites/IntranetUAT/${OthProps.Entity}/${folderName}`;
          //  (payloadForFolderMaster as any).FolderPath=`/sites/AlRostmanispfx2/${OthProps.Entity}/${folderName}`;
           (payloadForFolderMaster as any).FolderPath=`/sites/Intranetdemos/${selectedSite}/${selectedSitedepartment}/${subsiteName}`;
          //  (payloadForFolderMaster as any).FolderPath=`/sites/AlRostmani/${OthProps.Entity}/${folderName}`;
          (payloadForFolderMaster as any).IsLibrary=false;
          (payloadForFolderMaster as any).IsFolder=true;
           ( payloadForFolderMaster as any).FolderName= subsiteName;
            (payloadForFolderMaster as any).IsPrivate=false;
            (payloadForFolderMaster as any).IsSubsite='Yes';
            
          
            (payloadForFolderMaster as any).IsFolderDeligation=false;
          
       
            (payloadForFolderMaster as any).External=false;
            
                (payloadForFolderMaster as any).CurrentUser= currentUserEmail;
                (payloadForFolderMaster as any).IsSite='No';
                (payloadForFolderMaster as any).IsActive=true;
            (payloadForFolderMaster as any).SiteURL=`https://officeindia.sharepoint.com/sites/Intranetdemos/${selectedSite}/${siteName}`;
              const addedItem = await sp.web.lists.getByTitle("DMSFolderMaster").items.add(payloadForFolderMaster);
          
               
const createGroups = async (baseName: string) => {
  try {
    // Group suffixes
    const suffixes = ["Admin", "Read", "Contribute", "Approval"];

    for (const suffix of suffixes) {
      const groupName = `${baseName}_${suffix}`;

      // Check if group exists already
      let exists = false;
      try {
        await sp.web.siteGroups.getByName(groupName)();
        exists = true;
        console.log(`✅ Group already exists: ${groupName}`);
      } catch {
        exists = false;
      }

      if (!exists) {
        await sp.web.siteGroups.add({
          Title: groupName,
          Description: `Group for ${groupName}`,
          OnlyAllowMembersViewMembership: false,
          AllowMembersEditMembership: false,
          AllowRequestToJoinLeave: false,
          AutoAcceptRequestToJoinLeave: false,
        });
        console.log(`🎉 Created group: ${groupName}`);
      }
    }
  } catch (err: any) {
    console.error("❌ Error creating groups:", err.message);
  }
};

// Example usage
createGroups(subsiteName); 

              if(folderAddResult){
               Swal.fire({
                title: "Success!",  
                text: "Subsite created successfully",
                icon: "success"
              });
              }
        
    }catch(error) { 
      console.error("Error creating subsite:", error);
    }
  }

  // Filter and sort functionality for the table
const [filters, setFilters] = React.useState({
  SNo: '',
  SiteTitle: '',
  SiteCollection: '',
  FolderName: '',
  IsSite: '',
  CurrentUser: '',
  Created: '',
});

  const [sortConfig, setSortConfig] = React.useState({ key: '', direction: 'ascending' });

  const handleFilterChange = (e: React.ChangeEvent<HTMLInputElement>, field: string) => {
    setFilters({
      ...filters,
      [field]: e.target.value,
    });
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
     (filters.SiteTitle === '' || (item.DocumentLibraryName && item.DocumentLibraryName.toLowerCase().includes(filters.SiteTitle.toLowerCase()))) &&
    (filters.SiteCollection === '' || (item.SiteTitle && item.SiteTitle.toLowerCase().includes(filters.SiteCollection.toLowerCase()))) &&
    (filters.IsSite === '' || (item.IsSite && item.IsSite.toLowerCase().includes(filters.IsSite.toLowerCase()))) &&
    (filters.CurrentUser === '' || (item.Author && item.Author.Title && item.Author.Title.toLowerCase().includes(filters.CurrentUser.toLowerCase()))) &&
    (filters.Created === '' || (item.Created && item.Created.toLowerCase().includes(filters.Created.toLowerCase())))
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

  const filteredDmsData = applyFiltersAndSorting(dmsFolderData);
  
  // Pagination
  const [currentPage, setCurrentPage] = React.useState(1);
  const itemsPerPage = 10;
  const totalPages = Math.ceil(filteredDmsData.length / itemsPerPage);
  
  const handlePageChange = (pageNumber: any) => {
    if (pageNumber > 0 && pageNumber <= totalPages) {
      setCurrentPage(pageNumber);
    }
  };
  
  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const currentData = filteredDmsData.slice(startIndex, endIndex);

  interface PaginationProps {
    currentPage: number;
    totalPages: any;
    handlePageChange: any;
  }
  
  const Pagination = ({ currentPage, totalPages, handlePageChange }: PaginationProps) => {
    const pageLimit = 5;
    const startPage = Math.max(1, currentPage - Math.floor(pageLimit / 2));
    const endPage = Math.min(totalPages, startPage + pageLimit - 1);
    const adjustedStartPage = Math.max(1, Math.min(startPage, totalPages - pageLimit + 1));
    
    const visiblePages = Array.from(
      { length: Math.min(pageLimit, totalPages) },
      (_, index) => adjustedStartPage + index
    );

    return (
      <nav className="pagination-container">
        <ul className="pagination">
          <li className={`page-item ${currentPage === 1 ? 'disabled' : ''}`}>
            <a
              className="page-link PreviousPage"
              onClick={() => handlePageChange(currentPage - 1)}
              aria-label="Previous"
            >
              «
            </a>
          </li>
          
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
    <div>

      {viewMode === "list" && (
  <div className={styles.argform}>
          <div className='row'>
            <div className='col-md-7 pt-0'>
              <div className='page-title fw-bold mb-1 font-20'>Site - (Department) </div>
            </div>
            <div className='col-md-5'>
              <div className="padd-right1 mt-0">
                {/* <div
                  onClick={handleButtonClickShow}
                  className={styles.addbuttonargform}
                >
                  <p className={styles.Addtext}>Create New</p>
                </div> */}
                <Select
  classNamePrefix="custom-select"
  options={[
    { value: "createSite", label: "Create New Site - (Department)" },
    { value: "createSubSite", label: "Create New SubSite - (Section)" },
  ]}
  defaultValue={{ value: "createSubSite", label: "Create New Site/SubSite" }}
  isSearchable={false}
  isClearable={false}
  onChange={(selectedOption) => setViewMode(selectedOption?.value || "list")}
  styles={{
    control: (provided, state) => ({
      ...provided,
      borderColor: "#7fc4de", // border color
      boxShadow: state.isFocused ? "0 0 0 1px #7fc4de" : "none",
      "&:hover": {
        borderColor: "#7fc4de",
      },
    }),
    option: (provided, state) => ({
      ...provided,
      backgroundColor: state.isFocused
        ? "#7fc4de" // hover
        : state.isSelected
        ? "#7fc4de" // when clicked/selected
        : "white", // default (no background)
      color: state.isFocused || state.isSelected ? "white" : "black",
      "&:active": {
        backgroundColor: "#7fc4de",
      },
    }),
    singleValue: (provided) => ({
      ...provided,
      color: "black", // selected text stays black
    }),
    dropdownIndicator: (provided) => ({
      ...provided,
      color: "#7fc4de", // dropdown arrow color
    }),
    indicatorSeparator: () => ({
      display: "none", // hide separator line
    }),
  }}
/>

              </div>
            </div>
          </div>
          <div className='mt-3'>
            <div style={{padding:'15px'}} className={styles.container}>
              <table className="mtbalenew">
              <thead>
  <tr>
    <th style={{ borderBottomLeftRadius: '0px', minWidth: '55px', maxWidth: '55px', borderTopLeftRadius: '0px' }}>
      <div className="pb-0">
        <span>S.No.</span>
        <br />
        <input
          type="text"
          placeholder="Search"
          value={filters.SNo}
          onChange={(e) => handleFilterChange(e, 'SNo')}
          className="form-control form-control-sm"
        />
      </div>
    </th>
  <th>
      <div className="pb-0">
        <span>Site Collection - (Location)</span>
        <br />
        <input
          type="text"
          placeholder="Search"
          value={filters.SiteCollection || ''}
          onChange={(e) => handleFilterChange(e, 'SiteCollection')}
          className="form-control form-control-sm"
        />
      </div>
    </th>
    <th>
      <div className="pb-0">
        <span>Site - (Department)</span> &nbsp;
        {/* <span className="Sorting" onClick={() => handleSortChange('Title')}>
          <FontAwesomeIcon icon={faSort} />
        </span> */}
        <br />
        <input
          type="text"
          placeholder="Search"
          value={filters.SiteTitle}
          onChange={(e) => handleFilterChange(e, 'SiteTitle')}
          className="form-control form-control-sm"
        />
      </div>
    </th>
    <th>
      <div className="pb-0">
        <span>SubSite - (Section)</span> &nbsp;
        {/* <span className="Sorting" onClick={() => handleSortChange('Title')}>
          <FontAwesomeIcon icon={faSort} />
        </span> */}
        <br />
        <input
          type="text"
          placeholder="Search"
          value={filters.FolderName}
          onChange={(e) => handleFilterChange(e, 'FolderName')}
          className="form-control form-control-sm"
        />
      </div>
    </th>

  
{/* 
    <th>
      <div className="pb-0">
        <span>IsActive</span>
        <br />
        <input
          type="text"
          placeholder="Search"
          value={filters.IsSite || ''}
          onChange={(e) => handleFilterChange(e, 'IsSite')}
          className="form-control form-control-sm"
        />
      </div>
    </th> */}

    <th>
      <div className="pb-0">
        <span>Created At</span>
        <br />
        <input
          type="text"
          placeholder="Search"
          value={filters.Created}
          onChange={(e) => handleFilterChange(e, 'Created')}
          className="form-control form-control-sm"
        />
      </div>
    </th>

    <th>
      <div className="pb-0">
        <span>Created By</span> &nbsp;
        <span onClick={() => handleSortChange('Author')}>
          <FontAwesomeIcon icon={faSort} />
        </span>
        <br />
        <input
          type="text"
          placeholder="Search"
          value={filters.CurrentUser}
          onChange={(e) => handleFilterChange(e, 'CurrentUser')}
          className="form-control form-control-sm"
        />
      </div>
    </th>
  </tr>
</thead>

                <tbody>
                  {currentData.map((item, index) => (
                    <React.Fragment key={item.Id}>
                      <tr>
                        <td style={{ borderBottomLeftRadius: '0px', textAlign:'center', minWidth: '55px', maxWidth: '55px', borderTopLeftRadius: '0px' }}>
                          <span className='indexdesign'>
                            {(currentPage - 1) * itemsPerPage + index + 1}
                          </span>
                        </td>
                           <td className="">
                          {item.SiteTitle || 'No Title'}
                        </td>
                        <td className="">
                          {item.DocumentLibraryName || 'No Title'}
                        </td>

                        <td className="">
                          {item.FolderName || '---'}
                        </td>
                     
                        {/* <td>
                          {item.IsSite || 'No Site Info'}
                        </td> */}
                        <td>
                          {item.Created}
                        </td>
                        <td>
                          {item.Author?.Title || 'No Author'}
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
            </div>
          </div>
        </div>
    )}

    {
      viewMode === "createSite" && (
        <div className={styles.argform}>
          <div style={{marginBottom:"20px"}} className='row'>
            <div className='col-md-7'>
              <div className='page-title fw-bold mb-1 font-20'>Create Site </div>
            </div>
            <div className='col-md-5'>
              <div className='padd-right1 mt-0'>
                <div className={styles.actions}>
                  <div
                    className={styles.backbuttonform}
                    onClick={handleBackButtonClick}
                  >
                    <p className={styles.Addtext}>Back</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
          
          {/* New Site Creation Form */}
          <div >
            <form onSubmit={handleSubmit}>
              <div className="row mb-3">
                <div className="col-md-6">
                  <label htmlFor="siteDropdown" className="form-label">Select Site Collection - (Location)</label>
                  <select 
                    className="form-select" 
                    id="siteDropdown"
                    value={selectedSite}
                    onChange={(e) => setSelectedSite(e.target.value)}
                    required
                  >
                    <option value="">Select a Site Collection - (Location)</option>
                    {activeSites.map((site) => (
                      <option key={site.Id} value={site.Title}>
                        {site.Title}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
              
              <div className="row mb-3">
                <div className="col-md-6">
                  <label htmlFor="siteName" className="form-label">Site Name - (Department)</label>
                  <input 
                    type="text" 
                    className="form-control" 
                    id="siteName"
                    value={siteName}
                    onChange={(e) => setSiteName(e.target.value)}
                    placeholder="Enter site name - (Department)"
                    required
                  />
                </div>
              </div>
              
              <div className="row">
                <div className="col-md-12">
                  <button type="submit" style={{padding : "7px" ,   margin : "20px 20px 20px 0px"}} className={`btn btn-primary`}>
                    Create Site - (Department)
                  </button>
                  <button 
                    type="button" 
                    className={`btn btn-secondary`}
                    onClick={handleBackButtonClick}
                  >
                    Cancel
                  </button>
                </div>
              </div>
            </form>
          </div>
          {/* <div className={styles.DmsAdminForm}>
  <div className={styles.formcontainer}>
    <form onSubmit={handleSubmit}>
      
      <div>
   <div className="row mb-3">
        <div className="col-md-6">
          <label htmlFor="siteDropdown" className={styles.label}>
            Select Site Collection - (Location)
            <span style={{ color: "red", fontWeight: "bold" }}> *</span>
          </label>
          <select
            className={styles.inputform1}
            id="siteDropdown"
            value={selectedSite}
            onChange={(e) => setSelectedSite(e.target.value)}
            required
          >
            <option value="">Select a Site Collection - (Location)</option>
            {activeSites.map((site) => (
              <option key={site.Id} value={site.Title}>
                {site.Title}
              </option>
            ))}
          </select>
        </div>
      </div>

      <div className="row mb-3">
        <div className="col-md-6">
          <label htmlFor="siteName" className={styles.label}>
            Site Name - (Department)
            <span style={{ color: "red", fontWeight: "bold" }}> *</span>
          </label>
          <input
            type="text"
            className={styles.inputform1}
            id="siteName"
            value={siteName}
            onChange={(e) => setSiteName(e.target.value)}
            placeholder="Enter site name"
            required
          />
        </div>
      </div>

      </div>
   
      <div className="mt-2 text-center mb-2">
        <button type="submit" className={styles.backbuttonform1}>
          <p className={styles.Addtext}>Create Site - (Department)</p>
        </button>
        <button
          type="button"
          style={{ marginBottom: "15px" }}
          className={styles.addbuttonargform1}
          onClick={handleBackButtonClick}
        >
          <p className={styles.Addtext}>Cancel</p>
        </button>
      </div>
    </form>
  </div>
</div> */}



        </div>
      )
    }
    {
      viewMode === "createSubSite" && (
        
    <div className={styles.argform}>
          <div style={{marginBottom:"20px"}} className='row'>
            <div className='col-md-7'>
              <div className='page-title fw-bold mb-1 font-20'>Create SubSite </div>
            </div>
            <div className='col-md-5'>
              <div className='padd-right1 mt-0'>
                <div className={styles.actions}>
                  <div
                    className={styles.backbuttonform}
                    onClick={handleBackButtonClick}
                  >
                    <p className={styles.Addtext}>Back</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
          
          {/* New Site Creation Form */}
          <div >
            <form onSubmit={handleSubmit2}>
              <div className="row mb-3">
                <div className="col-md-6">
                  <label htmlFor="siteDropdown" className="form-label">Select Site Collection - (Location)</label>
                  <select 
                    className="form-select" 
                    id="siteDropdown"
                    value={selectedSite}
                     onChange={(e:any) => {
    const selectedSite = activeSites.find((site) => site.Title === e.target.value);
    if (selectedSite) {
      setSiteID(selectedSite.SiteID);
      
    }
    getdataofdepartment(e);
  }}
                    required
                  >
                    <option value="">Select a Site Collection - (Location)</option>
                    {activeSites.map((site) => (
                      <option key={site.Id} value={site.Title}>
                        {site.Title}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
              
              <div className="row mb-3">
                  <div className="col-md-6">
                  <label htmlFor="siteDropdown" className="form-label">Select Site Name - (Department)</label>
                 <select
  className="form-select"
  id="siteDropdown"
  value={selectedSitedepartment}
  onChange={(e) => setSelectedSitedepartment(e.target.value)}
  required
>
  <option value="">Select a Site - (Department)</option>
  {activeSitesdepartment.map((site) => (
    <option key={site.Id} value={site.DocumentLibraryName}>
      {site.DocumentLibraryName}
    </option>
  ))}
</select>
                </div>
              </div>

              <div className="row mb-3">
                <div className="col-md-6">
                  <label htmlFor="subsiteName" className="form-label">SubSite Name - (Section)</label>
                  <input 
                    type="text" 
                    className="form-control" 
                    id="subsiteName"
                    value={subsiteName}
                    onChange={(e) => setSubSiteName(e.target.value)}
                    placeholder="Enter site name - (Section)"
                    required
                  />
                </div>
              </div>
              
              <div className="row">
                <div className="col-md-12">
                  <button type="submit" style={{padding : "7px" ,   margin : "20px 20px 20px 0px"}} className={`btn btn-primary`}>
                    Create SubSite
                  </button>
                  <button 
                    type="button" 
                    className={`btn btn-secondary`}
                    onClick={handleBackButtonClick}
                  >
                    Cancel
                  </button>
                </div>
              </div>
            </form>
          </div>
        </div>
            )
    }
      {/* {showFirstDiv ? (
        <div className={styles.argform}>
          <div className='row'>
            <div className='col-md-7 pt-0'>
              <div className='page-title fw-bold mb-1 font-20'>Site - (Department) </div>
            </div>
            <div className='col-md-5'>
              <div className="padd-right1 mt-0">
                
                   <Select
 
      classNamePrefix="custom-select"
      options={[{ value: 'Create New Site', label: 'Create New Site' } , { value: 'Create New SubSite', label: 'Create New SubSite' }]}
      defaultValue={{ value: 'Create New SubSite', label: 'Create New SubSite' }}
      isSearchable={false}
      isClearable={false}
      onChange={(selectedOption) => {
        if (selectedOption?.value === 'Create New Site') {
          handleButtonClickShow();
        } 
      }}
    />
              </div>
            </div>
          </div>
          <div className='mt-3'>
            <div style={{padding:'15px'}} className={styles.container}>
              <table className="mtbalenew">
              <thead>
  <tr>
    <th style={{ borderBottomLeftRadius: '0px', minWidth: '55px', maxWidth: '55px', borderTopLeftRadius: '0px' }}>
      <div className="pb-0">
        <span>S.No.</span>
        <br />
        <input
          type="text"
          placeholder="Search"
          value={filters.SNo}
          onChange={(e) => handleFilterChange(e, 'SNo')}
          className="form-control form-control-sm"
        />
      </div>
    </th>
  <th>
      <div className="pb-0">
        <span>Site Collection - (Location)</span>
        <br />
        <input
          type="text"
          placeholder="Search"
          value={filters.SiteCollection || ''}
          onChange={(e) => handleFilterChange(e, 'SiteCollection')}
          className="form-control form-control-sm"
        />
      </div>
    </th>
    <th>
      <div className="pb-0">
        <span>Site - (Department)</span> &nbsp;
        <span className="Sorting" onClick={() => handleSortChange('Title')}>
          <FontAwesomeIcon icon={faSort} />
        </span>
        <br />
        <input
          type="text"
          placeholder="Search"
          value={filters.SiteTitle}
          onChange={(e) => handleFilterChange(e, 'SiteTitle')}
          className="form-control form-control-sm"
        />
      </div>
    </th>

  

    <th>
      <div className="pb-0">
        <span>IsActive</span>
        <br />
        <input
          type="text"
          placeholder="Search"
          value={filters.IsSite || ''}
          onChange={(e) => handleFilterChange(e, 'IsSite')}
          className="form-control form-control-sm"
        />
      </div>
    </th>

    <th>
      <div className="pb-0">
        <span>Created At</span>
        <br />
        <input
          type="text"
          placeholder="Search"
          value={filters.Created}
          onChange={(e) => handleFilterChange(e, 'Created')}
          className="form-control form-control-sm"
        />
      </div>
    </th>

    <th>
      <div className="pb-0">
        <span>Created By</span> &nbsp;
        <span onClick={() => handleSortChange('Author')}>
          <FontAwesomeIcon icon={faSort} />
        </span>
        <br />
        <input
          type="text"
          placeholder="Search"
          value={filters.CurrentUser}
          onChange={(e) => handleFilterChange(e, 'CurrentUser')}
          className="form-control form-control-sm"
        />
      </div>
    </th>
  </tr>
</thead>

                <tbody>
                  {currentData.map((item, index) => (
                    <React.Fragment key={item.Id}>
                      <tr>
                        <td style={{ borderBottomLeftRadius: '0px', textAlign:'center', minWidth: '55px', maxWidth: '55px', borderTopLeftRadius: '0px' }}>
                          <span className='indexdesign'>
                            {(currentPage - 1) * itemsPerPage + index + 1}
                          </span>
                        </td>
                           <td className="">
                          {item.SiteTitle || 'No Title'}
                        </td>
                        <td className="">
                          {item.DocumentLibraryName || 'No Title'}
                        </td>
                     
                        <td>
                          {item.IsSite || 'No Site Info'}
                        </td>
                        <td>
                          {item.Created}
                        </td>
                        <td>
                          {item.Author?.Title || 'No Author'}
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
      ) : 
      
      (
        <div className={styles.argform}>
          <div style={{marginBottom:"20px"}} className='row'>
            <div className='col-md-7'>
              <div className='page-title fw-bold mb-1 font-20'>Create Site </div>
            </div>
            <div className='col-md-5'>
              <div className='padd-right1 mt-0'>
                <div className={styles.actions}>
                  <div
                    className={styles.backbuttonform}
                    onClick={handleBackButtonClick}
                  >
                    <p className={styles.Addtext}>Back</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
          
        
          <div >
            <form onSubmit={handleSubmit}>
              <div className="row mb-3">
                <div className="col-md-6">
                  <label htmlFor="siteDropdown" className="form-label">Select Site Collection - (Location)</label>
                  <select 
                    className="form-select" 
                    id="siteDropdown"
                    value={selectedSite}
                    onChange={(e) => setSelectedSite(e.target.value)}
                    required
                  >
                    <option value="">Select a Site Collection - (Location)</option>
                    {activeSites.map((site) => (
                      <option key={site.Id} value={site.Title}>
                        {site.Title}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
              
              <div className="row mb-3">
                <div className="col-md-6">
                  <label htmlFor="siteName" className="form-label">Site Name - (Department)</label>
                  <input 
                    type="text" 
                    className="form-control" 
                    id="siteName"
                    value={siteName}
                    onChange={(e) => setSiteName(e.target.value)}
                    placeholder="Enter site name"
                    required
                  />
                </div>
              </div>
              
              <div className="row">
                <div className="col-md-12">
                  <button type="submit" className={`btn btn-primary`}>
                    Create Site - (Department)
                  </button>
                  <button 
                    type="button" 
                    className={`btn btn-secondary`}
                    onClick={handleBackButtonClick}
                  >
                    Cancel
                  </button>
                </div>
              </div>
            </form>
          </div>
        </div>
      )
      
      } */}
    </div>
  );
};

export default Site;