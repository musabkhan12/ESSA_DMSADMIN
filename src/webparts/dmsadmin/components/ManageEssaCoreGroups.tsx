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

interface Props {
  context: WebPartContext;
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
  const handleAdd = async () => {
    for (const p of selectedPrincipals) {
      await sp.web.siteGroups.getByName(selectedGroup).users.add(p.value);
    }
    Swal.fire("Added!", "", "success");
    setSelectedPrincipals([]);
    loadGroupUsers();
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

        <div style={{ width: "420px" }}>
          <Select
            isMulti
            options={principalOptions}
            value={selectedPrincipals}
            onChange={(val: any) => setSelectedPrincipals(val)}
            placeholder="Search user or group..."
          />
        </div>

        <button style={{backgroundColor: 'rgb(44, 153, 66)', borderColor:'rgb(44, 153, 66)'}} className="btn btn-success" onClick={handleAdd}>
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
      <div className="d-flex justify-content-center mt-3">
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
      </div>
    </div>
  );
};

export default ManageEssaCoreGroups;