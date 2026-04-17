// import * as React from "react";
// import { spfi, SPFI } from "@pnp/sp";
// import { SPFx } from "@pnp/sp";
// import { WebPartContext } from "@microsoft/sp-webpart-base";
// import Swal from "sweetalert2";
// import Select from "react-select";
// import "bootstrap/dist/css/bootstrap.min.css";
// import "../../../CustomCss/mainCustom.scss";

// interface Props {
//   context: WebPartContext;
// }

// const GROUPS = ["ESSA Owners", "ESSA Members", "ESSA Visitors"];

// const ManageEssaCoreGroups: React.FC<Props> = ({ context }) => {
//   const sp: SPFI = React.useMemo(() => spfi().using(SPFx(context)), [context]);

//   const [selectedGroup, setSelectedGroup] = React.useState("ESSA Owners");
//   const [groupUsers, setGroupUsers] = React.useState<any[]>([]);
//   const [principalOptions, setPrincipalOptions] = React.useState<any[]>([]);
//   const [selectedPrincipals, setSelectedPrincipals] = React.useState<any[]>([]);

//   // Pagination
//   const [currentPage, setCurrentPage] = React.useState(1);
//   const itemsPerPage = 10;

//   // Load users + groups for dropdown
//   React.useEffect(() => {
//     loadPrincipals();
//   }, []);

//   const loadPrincipals = async () => {
//     const users = await sp.web.siteUsers();
//     const groups = await sp.web.siteGroups();

//     const userOptions = users.map((u) => ({
//       label: `${u.Title} (${u.Email})`,
//       value: u.LoginName,
//       email: u.Email,
//     }));

//     const groupOptions = groups.map((g) => ({
//       label: g.Title,
//       value: g.LoginName,
//     }));

//     setPrincipalOptions([...userOptions, ...groupOptions]);
//   };

//   // Load selected group users
//   React.useEffect(() => {
//     loadGroupUsers();
//   }, [selectedGroup]);

//   const loadGroupUsers = async () => {
//     const users = await sp.web.siteGroups.getByName(selectedGroup).users();
//     setGroupUsers(users);
//     setCurrentPage(1);
//   };

//   // Add users/groups
//   const handleAdd = async () => {
//     if (!selectedPrincipals || selectedPrincipals.length === 0) {
//       Swal.fire("Select user/group first");
//       return;
//     }

//     for (const p of selectedPrincipals) {
//       await sp.web.siteGroups.getByName(selectedGroup).users.add(p.value);
//     }

//     Swal.fire("Added successfully", "", "success");
//     setSelectedPrincipals([]);
//     loadGroupUsers();
//   };

//   // Remove user
//   const removeUser = async (loginName: string, name: string) => {
//     const confirm = await Swal.fire({
//       title: `Remove ${name}?`,
//       showCancelButton: true,
//     });

//     if (confirm.isConfirmed) {
//       await sp.web.siteGroups
//         .getByName(selectedGroup)
//         .users.removeByLoginName(loginName);

//       loadGroupUsers();
//     }
//   };

//   // Pagination logic
//   const totalPages = Math.ceil(groupUsers.length / itemsPerPage);
//   const startIndex = (currentPage - 1) * itemsPerPage;
//   const currentData = groupUsers.slice(startIndex, startIndex + itemsPerPage);

//   return (
//     <div className="argform">
//       <div className="page-title fw-bold font-20 mb-3">
//         Admin Panel &gt; Manage Core Groups
//       </div>

//       {/* Top Controls */}
//       <div className="d-flex gap-3 align-items-center mb-3">
//         <label className="fw-bold">Select Group:</label>

//         <select
//           className="form-select w-25"
//           value={selectedGroup}
//           onChange={(e) => setSelectedGroup(e.target.value)}
//         >
//           {GROUPS.map((g) => (
//             <option key={g}>{g}</option>
//           ))}
//         </select>

//         <div style={{ width: "420px" }}>
//           <Select
//             isMulti
//             options={principalOptions}
//             value={selectedPrincipals}
//             onChange={(val: any) => setSelectedPrincipals(val)}
//             placeholder="Search user or group..."
//             noOptionsMessage={() => "No result found"}
//           />
//         </div>

//         <button className="btn btn-success" onClick={handleAdd}>
//           Add
//         </button>
//       </div>

//       {/* Table */}
//       <table className="mtbalenew">
//         <thead>
//           <tr>
//             <th style={{ width: "60px" }}>S.No</th>
//             <th>User</th>
//             <th>Email</th>
//             <th style={{ width: "80px" }}>Action</th>
//           </tr>
//         </thead>
//         <tbody>
//           {currentData.map((u, i) => (
//             <tr key={u.Id}>
//               <td>{startIndex + i + 1}</td>
//               <td>{u.Title}</td>
//               <td>{u.Email}</td>
//               <td>
//                 <img
//                   src={require("../assets/del.png")}
//                   className="action-icon"
//                   onClick={() => removeUser(u.LoginName, u.Title)}
//                 />
//               </td>
//             </tr>
//           ))}
//         </tbody>
//       </table>

//       {/* Pagination */}
//       <div className="d-flex justify-content-center mt-3">
//         <button
//           className="btn btn-light me-2"
//           disabled={currentPage === 1}
//           onClick={() => setCurrentPage(currentPage - 1)}
//         >
//           «
//         </button>
//         <span className="align-self-center">
//           Page {currentPage} of {totalPages}
//         </span>
//         <button
//           className="btn btn-light ms-2"
//           disabled={currentPage === totalPages}
//           onClick={() => setCurrentPage(currentPage + 1)}
//         >
//           »
//         </button>
//       </div>
//     </div>
//   );
// };

// export default ManageEssaCoreGroups;
import * as React from "react";
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import Swal from "sweetalert2";
import Select from "react-select";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faSort } from "@fortawesome/free-solid-svg-icons";
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import "../components/BasicForm.module.scss"

interface Props {
  context: WebPartContext;
}

// srs 17/4/26
interface PaginationProps {
  currentPage: number;
  totalPages: number;
  handlePageChange: (pageNumber: number) => void;
}

const GROUPS = ["ESSA Owners", "ESSA Members", "ESSA Visitors"];

const ManageEssaCoreGroups: React.FC<Props> = ({ context }) => {
  const sp: SPFI = React.useMemo(() => spfi().using(SPFx(context)), [context]);

  const [selectedGroup, setSelectedGroup] = React.useState("ESSA Owners");
  const [groupUsers, setGroupUsers] = React.useState<any[]>([]);
  const [principalOptions, setPrincipalOptions] = React.useState<any[]>([]);
  const [selectedPrincipals, setSelectedPrincipals] = React.useState<any[]>([]);

  // Filters & Sorting
  const [filters, setFilters] = React.useState<any>({
    Title: "",
    Email: "",
  });
  const [sortConfig, setSortConfig] = React.useState({ key: "", direction: "ascending" });

  // Pagination
  const [currentPage, setCurrentPage] = React.useState(1);
  const itemsPerPage = 10;

  // Load dropdown users + groups
  React.useEffect(() => {
    loadPrincipals();
  }, []);

  const loadPrincipals = async () => {
    const users = await sp.web.siteUsers();
    const groups = await sp.web.siteGroups();

    const options = [
      ...users.map((u) => ({
        label: `${u.Title} (${u.Email})`,
        value: u.LoginName,
      })),
      ...groups.map((g) => ({
        label: `${g.Title} (Group)`,
        value: g.LoginName,
      })),
    ];

    setPrincipalOptions(options);
  };

  // Load group users
  React.useEffect(() => {
    loadGroupUsers();
  }, [selectedGroup]);

  const loadGroupUsers = async () => {
    const users = await sp.web.siteGroups.getByName(selectedGroup).users();
    setGroupUsers(users);
    setCurrentPage(1);
  };

  // Add
  // const handleAdd = async () => {
  //   for (const p of selectedPrincipals) {
  //     await sp.web.siteGroups.getByName(selectedGroup).users.add(p.value);
  //   }
  //   Swal.fire("Added!", "", "success");
  //   setSelectedPrincipals([]);
  //   loadGroupUsers();
  // };


// srs 17/4/26
  const handleAdd = async () => {
  // 1. Validation Check
  // Check if Group is default/empty and if any Principals are selected
  if (!selectedGroup || selectedGroup === "Select Group" || selectedPrincipals.length === 0) {
    Swal.fire({
      html: `
        <div style="padding: 10px; text-align: center; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
          <h2 style="
            color: #595959; 
            font-size: 28px; 
            font-weight: 600; 
            margin-bottom: 20px;
          ">Please fill out the fields!</h2>
          
          <p style="
            color: #545454; 
            font-size: 18px; 
            font-weight: 400; 
            margin-bottom: 25px;
          ">Please select a group and at least one user/group to add.</p>
          
          <button id="val-ok-button" style="
            background-color: #87CEEB; 
            color: white; 
            border: none; 
            padding: 10px 40px; 
            font-size: 18px; 
            border-radius: 5px; 
            cursor: pointer;
            font-weight: 500;
          ">OK</button>
        </div>
      `,
      showConfirmButton: false,
      width: '500px',
      didOpen: () => {
        const popup = Swal.getPopup();
        if (popup) popup.style.borderRadius = '5px';
        
        const btn = document.getElementById('val-ok-button');
        if (btn) btn.addEventListener('click', () => Swal.close());
      }
    });
    return; // Stop execution
  }

  // 2. Process Addition
  try {
    // Using Promise.all is faster than a standard for-loop for multiple additions
    const addPromises = selectedPrincipals.map(p => 
      sp.web.siteGroups.getByName(selectedGroup).users.add(p.value)
    );
    
    await Promise.all(addPromises);

    Swal.fire({
      title: "Added!",
      text: "Users have been successfully added to the group.",
      icon: "success",
      confirmButtonColor: "rgb(44, 153, 66)"
    });

    setSelectedPrincipals([]);
    loadGroupUsers();
  } catch (error) {
    console.error("Error adding users:", error);
    Swal.fire("Error", "Failed to add users to the group. Check permissions.", "error");
  }
};

  // Delete
  const removeUser = async (login: string, name: string) => {
    const confirm = await Swal.fire({ title: `Remove ${name}?`, showCancelButton: true });
    if (confirm.isConfirmed) {
      await sp.web.siteGroups.getByName(selectedGroup).users.removeByLoginName(login);
      loadGroupUsers();
    }
  };

  // Filtering
  const handleFilterChange = (e: any, field: string) => {
    setFilters({ ...filters, [field]: e.target.value.toLowerCase() });
  };

  // Sorting
  const handleSort = (key: string) => {
    let direction = "ascending";
    if (sortConfig.key === key && sortConfig.direction === "ascending") {
      direction = "descending";
    }
    setSortConfig({ key, direction });
  };

  const processedData = [...groupUsers]
    .filter(
      (u) =>
        (u.Title || "").toLowerCase().includes(filters.Title) &&
        (u.Email || "").toLowerCase().includes(filters.Email)
    )
    .sort((a, b) => {
      if (!sortConfig.key) return 0;
      const aVal = (a[sortConfig.key] || "").toString().toLowerCase();
      const bVal = (b[sortConfig.key] || "").toString().toLowerCase();
      if (aVal < bVal) return sortConfig.direction === "ascending" ? -1 : 1;
      if (aVal > bVal) return sortConfig.direction === "ascending" ? 1 : -1;
      return 0;
    });

  // Pagination
  const totalPages = Math.ceil(processedData.length / itemsPerPage);
  const start = (currentPage - 1) * itemsPerPage;
  const currentData = processedData.slice(start, start + itemsPerPage);

  // srs 17/4/26
  const handlePageChange = (pageNumber: number) => {
    if (pageNumber > 0 && pageNumber <= totalPages) {
      setCurrentPage(pageNumber);
    }
  };

  // NEW: The Pagination Component from your reference
  const Pagination = ({ currentPage, totalPages, handlePageChange }: PaginationProps) => {
    const pageLimit = 5; 
    const startPage = Math.max(1, currentPage - Math.floor(pageLimit / 2));
    const adjustedStartPage = Math.max(1, Math.min(startPage, totalPages - pageLimit + 1));
    const visiblePages = Array.from(
      { length: Math.min(pageLimit, totalPages) },
      (_, index) => adjustedStartPage + index
    );

    return (
      <nav className="pagination-container">
        <ul className="pagination">
          <li className={`page-item ${currentPage === 1 ? 'disabled' : ''}`}>
            <a className="page-link PreviousPage" onClick={() => handlePageChange(currentPage - 1)} aria-label="Previous">
              «
            </a>
          </li>
          {visiblePages.map((pageNumber) => (
            <li key={pageNumber} className={`page-item ${currentPage === pageNumber ? 'active' : ''}`}>
              <a className="page-link" onClick={() => handlePageChange(pageNumber)}>
                {pageNumber}
              </a>
            </li>
          ))}
          <li className={`page-item ${currentPage === totalPages ? 'disabled' : ''}`}>
            <a className="page-link NextPage" onClick={() => handlePageChange(currentPage + 1)} aria-label="Next">
              »
            </a>
          </li>
        </ul>
      </nav>
    );
  };
  // srs 17/4/26 end
 
  return (
    <div className="argform">
   
 <div style={{marginTop:'78px'}} className="card card-body">
 <div className="page-title fw-bold font-20 mb-3">
        Admin Panel &gt; Manage Core Groups
      </div>
         <div className="d-flex gap-3 align-items-center mb-3">
        <label className="fw-bold">Select Group:</label>
        <select
          className="form-select w-25"
          value={selectedGroup}
          onChange={(e) => setSelectedGroup(e.target.value)}
        >
          {GROUPS.map((g) => (
            <option key={g}>{g}</option>
          ))}
        </select>

        {/* <div style={{ width: "420px" }}> */}
           <label className="fw-bold">Select User: <span className="text-danger">*</span></label>
          <Select
            isMulti
            options={principalOptions}
            value={selectedPrincipals}
            onChange={(val: any) => setSelectedPrincipals(val)}
            placeholder="Search user or group..."
            onKeyDown={(e: any) => {
      if (e.key === 'Enter') {
        e.preventDefault(); // Prevents the page from submitting/going back
      }
    }}
          />
        {/* </div> */}

        <button type = "button" style={{backgroundColor: 'rgb(44, 153, 66)', borderColor:'rgb(44, 153, 66)'}} className="btn btn-success" onClick={handleAdd}>
          Add
        </button>
      </div>
 </div>
      {/* Table */}
      <table className="mtbalenew mt-2">
        <thead>
          <tr>
            <th style={{minWidth:'70px',maxWidth:'70px'}}>S.No</th>
            <th style={{minWidth:'250px',maxWidth:'250px'}}>
              <span onClick={() => handleSort("Title")}>
                User <FontAwesomeIcon icon={faSort} />
              </span>
              <input
                className="inputcss"
                placeholder="Search User"
                onChange={(e) => handleFilterChange(e, "Title")}
                // srs 17/4/26
                onKeyDown={(e: any) => {
      if (e.key === 'Enter') {
        e.preventDefault(); // Prevents the page from submitting/going back
      }
    }}
              />
            </th>
            <th style={{minWidth:'250px',maxWidth:'250px'}}>
              <span onClick={() => handleSort("Email")}>
                Email <FontAwesomeIcon icon={faSort} />
              </span>
              <input
                className="inputcss"
                placeholder="Search Email"
                onChange={(e) => handleFilterChange(e, "Email")}
                //srs 17/4/26
                onKeyDown={(e: any) => {
      if (e.key === 'Enter') {
        e.preventDefault(); // Prevents the page from submitting/going back
      }
    }}
              />
            </th>
            <th style={{minWidth:'70px',maxWidth:'70px'}}>Action</th>
          </tr>
        </thead>
        <tbody>
          {currentData.map((u, i) => (
            <tr key={u.Id}>
              <td style={{minWidth:'70px',maxWidth:'70px'}}>
                <div className="d-flex align-items-center justify-content-center">
                    <span className="indexdesign" style={{marginLeft:'0px'}}> {start + i + 1}</span></div>
               </td>
              <td style={{minWidth:'250px',maxWidth:'250px'}}>{u.Title}</td>
              <td style={{minWidth:'250px',maxWidth:'250px'}}>{u.Email}</td>
              <td style={{minWidth:'70px',maxWidth:'70px'}}>
                <img
                  src={require("../assets/del.png")}
                  className="action-icon"
                  onClick={() => removeUser(u.LoginName, u.Title)}
                />
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      {/* Pagination */}
      {/* <div className="pagination-container">
        <button
          className="btn btn-light me-2"
          disabled={currentPage === 1}
          onClick={() => setCurrentPage(currentPage - 1)}
        >
          «
        </button>
        <span className="align-self-center">
          Page {currentPage} of {totalPages}
        </span>
        <button
          className="btn btn-light ms-2"
          disabled={currentPage === totalPages}
          onClick={() => setCurrentPage(currentPage + 1)}
        >
          »
        </button>
      </div> */}
      {/* // srs 17/4/26 */}
      <Pagination 
        currentPage={currentPage} 
        totalPages={totalPages} 
        handlePageChange={handlePageChange} 
      />
      {/* // srs 17/4/26 */}
    
    </div>
  );
};

export default ManageEssaCoreGroups;