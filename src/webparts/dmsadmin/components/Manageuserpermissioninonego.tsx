// import * as React from "react";
// import { PrimaryButton, DefaultButton } from "@fluentui/react";
// import { Modal } from "@fluentui/react/lib/Modal";
// import { DetailsList, IColumn } from "@fluentui/react/lib/DetailsList";
// import { Checkbox } from "@fluentui/react/lib/Checkbox";
// import { Spinner } from "@fluentui/react/lib/Spinner";
// import { SPFI } from "@pnp/sp";
// import { getSP } from "../loc/pnpjsConfig";
// import './Manageuserpermissioninonego.css'
// import "@pnp/sp/webs";
// import "@pnp/sp/folders";
// import "@pnp/sp/files";
// import "@pnp/sp/sites"
// import "@pnp/sp/presets/all"
// import "@pnp/sp/site-groups";
// export default function UserPermissionManager() {
//   const [users, setUsers] = React.useState<any[]>([]);
//   const [groups, setGroups] = React.useState<any[]>([]);
//   const [selectedUser, setSelectedUser] = React.useState<any>(null);
//   const [userGroups, setUserGroups] = React.useState<string[]>([]);
//   const [isModalOpen, setIsModalOpen] = React.useState(false);
//   const [loading, setLoading] = React.useState<boolean>(true);
//   const [searchQuery, setSearchQuery] = React.useState("");

//   const sp: SPFI = getSP();

//   React.useEffect(() => {
//     loadData();
//   }, []);

//   const loadData = async () => {
//     try {
//       setLoading(true);

//       const siteGroups = await sp.web.siteGroups.expand("Roles/RoleDefinitionBindings")();
//       console.log("Fetched Groups:", siteGroups);
//       // Get role assignments for each group
// const groupsWithPermissions = await Promise.all(
//   siteGroups.map(async (group) => {
//     try {
//      const roleAssignment: any = await sp.web.roleAssignments.filter(`PrincipalId eq ${group.Id}`)
//   .expand("Member,RoleDefinitionBindings")();

// return {
//   ...group,
//   Permissions: roleAssignment.RoleDefinitionBindings.map(
//     (r: any) => r.Name
//   ),
// };
//     } catch (e) {
//       console.error(`Error fetching permissions for group: ${group.Title}`, e);
//       return { ...group, Permissions: [] };
//     }
//   })
// );

// console.log("Groups with Permissions:", groupsWithPermissions);
// setGroups(groupsWithPermissions);
   

//       let userMap: any = {};
//       // Fetch users from all groups in parallel
//       await Promise.all(
//         siteGroups.map(async (g) => {
//           const grpUsers = await sp.web.siteGroups.getById(g.Id).users();
//           grpUsers.forEach((u) => {
//             if (!userMap[u.LoginName])
//               userMap[u.LoginName] = { ...u, Groups: [] };
//             userMap[u.LoginName].Groups.push(g.Title);
//           });
//         })
//       );

//       setUsers(Object.values(userMap));
//     } catch (err) {
//       console.error("Error loading user/group data", err);
//     } finally {
//       setLoading(false);
//     }
//   };

//   const openManageModal = (user: any) => {
//     setSelectedUser(user);
//     setUserGroups(user.Groups);
//     setIsModalOpen(true);
//   };

//   const toggleGroup = (group: string, checked?: boolean) => {
//     if (checked) {
//       setUserGroups((prev) => [...prev, group]);
//     } else {
//       setUserGroups((prev) => prev.filter((g) => g !== group));
//     }
//   };

//   const saveChanges = async () => {
//     if (!selectedUser) return;
//     try {
//       setLoading(true);

//       const oldGroups = selectedUser.Groups;
//       const added = userGroups.filter((g) => !oldGroups.includes(g));
//       const removed = oldGroups.filter((g: any) => !userGroups.includes(g));

//       // Perform add/remove operations
//       await Promise.all(
//         added.map((g) =>
//           sp.web.siteGroups.getByName(g).users.add(selectedUser.LoginName)
//         )
//       );
//       await Promise.all(
//         removed.map((g:any) =>
//           sp.web.siteGroups.getByName(g).users.removeByLoginName(
//             selectedUser.LoginName
//           )
//         )
//       );

//       setIsModalOpen(false);
//       loadData(); // Refresh
//     } catch (err) {
//       console.error("Error updating permissions", err);
//     } finally {
//       setLoading(false);
//     }
//   };

//   // const columns: IColumn[] = [
//   //   { key: "user", name: "User", fieldName: "Title", minWidth: 40 },
//   //   { key: "email", name: "Email", fieldName: "Email", minWidth: 40 },
//   //   {
//   //     key: "groups",
//   //     name: "Groups",
//   //     fieldName: "Groups",
//   //     minWidth: 350,
//   //     onRender: (item: any) => item.Groups.join(", "),
//   //   },
//   //   {
//   //     key: "action",
//   //     name: "Action",
//   //     minWidth: 100,
//   //     onRender: (item: any) => (
//   //       <PrimaryButton text="Manage" onClick={() => openManageModal(item)} />
//   //     ),
//   //   },
//   // ];
// const columns: IColumn[] = [
//   {
//     key: "user",
//     name: "User",
//     fieldName: "Title",
//     minWidth: 120,
//     maxWidth: 150,
//     isMultiline: false
//   },
//   {
//     key: "email",
//     name: "Email",
//     fieldName: "Email",
//     minWidth: 200,
//     maxWidth: 250,
//     isMultiline: false
//   },
//   {
//     key: "groups",
//     name: "Groups",
//     fieldName: "Groups",
//     minWidth: 400,
//     isMultiline: true, // allows wrapping for long group names
//     onRender: (item: any) => item.Groups.join(", "),
//   },
//   {
//     key: "action",
//     name: "Action",
//     minWidth: 100,
//     maxWidth: 120,
//     onRender: (item: any) => (
//       <PrimaryButton text="Manage" style={{backgroundColor:"#7fc4de" , border:"none"}} onClick={() => openManageModal(item)} />
//     ),
//   },
// ];

//   return (
//     <div>
//       <h2>User Permissions</h2>

//       {loading ? (
       
//          <div>
//                    <img
//                           src={require("../assets/ESSAROLLER.gif")}
//                           className="alignrightl"
//                           alt="Loading..."
//                         />
//                         <span id="loader">Loading user permissions...</span>
//          </div>
    
//       ) : (
//         <DetailsList items={users} columns={columns} />
//       )}

//       {/* <Modal isOpen={isModalOpen} onDismiss={() => setIsModalOpen(false)}>
//         <h3>Manage Permissions for {selectedUser?.Title}</h3>
      
//         <div style={{ marginTop: 20 }}>
//           <PrimaryButton text="Save Changes" onClick={saveChanges} />
//           <DefaultButton
//             text="Cancel"
//             onClick={() => setIsModalOpen(false)}
//             style={{ marginLeft: 10 }}
//           />
//         </div>
//       </Modal> */}
//       {/* <Modal isOpen={isModalOpen} onDismiss={() => setIsModalOpen(false)}>
//   <h3>Manage Permissions for {selectedUser?.Title}</h3>

//   {groups.map((g) => {
//     const isMember = userGroups.includes(g.Title);

//     return (
//       <div
//         key={g.Id}
//         style={{
//           display: "flex",
//           justifyContent: "space-between",
//           alignItems: "center",
//           marginBottom: 8,
//         }}
//       >
//         <span>{g.Title}</span>
//         {isMember ? (
//           <DefaultButton
//             text="Remove"
//             onClick={async () => {
//               try {
//                 await sp.web.siteGroups
//                   .getByName(g.Title)
//                   .users.removeByLoginName(selectedUser.LoginName);

//                 setUserGroups((prev) => prev.filter((grp) => grp !== g.Title));
//               } catch (err) {
//                 console.error("Error removing user:", err);
//               }
//             }}
//           />
//         ) : (
//           <PrimaryButton
//             text="Add"
//             onClick={async () => {
//               try {
//                 await sp.web.siteGroups
//                   .getByName(g.Title)
//                   .users.add(selectedUser.LoginName);

//                 setUserGroups((prev) => [...prev, g.Title]);
//               } catch (err) {
//                 console.error("Error adding user:", err);
//               }
//             }}
//           />
//         )}
//       </div>
//     );
//   })}

//   <div style={{ marginTop: 20 }}>
//     <DefaultButton text="Close" onClick={() => setIsModalOpen(false)} />
//   </div>
// </Modal> */}
// {/* <Modal isOpen={isModalOpen} onDismiss={() => setIsModalOpen(false)} >
//     <div style={{ width: "600px", minHeight: "400px", padding: "20px" }}>
//   <h3>Manage Permissions for {selectedUser?.Title}</h3>

  
  
//   <div style={{ marginBottom: "16px" }}>
//     <h4>Current Groups</h4>
//     {userGroups.length > 0 ? (
//       userGroups.map((g) => (
//         <div
//           key={g}
//           style={{
//             display: "flex",
//             justifyContent: "space-between",
//             alignItems: "center",
//             marginBottom: 6,
//             padding: "4px 8px",
//             background: "#f3f2f1",
//             borderRadius: 4,
//           }}
//         >
//           <span>{g}</span>
//           <DefaultButton
//             text="Remove"
//             onClick={async () => {
//               try {
//                 await sp.web.siteGroups
//                   .getByName(g)
//                   .users.removeByLoginName(selectedUser.LoginName);

//                 setUserGroups((prev) => prev.filter((grp) => grp !== g));
//               } catch (err) {
//                 console.error("Error removing user:", err);
//               }
//             }}
//           />
//         </div>
//       ))
//     ) : (
//       <p style={{ fontStyle: "italic", color: "gray" }}>
//         User does not belong to any groups.
//       </p>
//     )}
//   </div>


//   <div style={{ marginBottom: "12px" }}>
//     <input
//       type="text"
//       placeholder="Search groups to add..."
//       style={{
//         width: "100%",
//         padding: "8px",
//         borderRadius: 4,
//         border: "1px solid #ccc",
//       }}
//       onChange={(e) => setSearchQuery(e.target.value)}
//     />
//   </div>

//   <div>
//     <h4>Available Groups</h4>
//     {groups
//       .filter(
//         (g) =>
//           !userGroups.includes(g.Title) && // only show groups user is NOT in
//           g.Title.toLowerCase().includes(searchQuery.toLowerCase())
//       )
//       .map((g) => (
//         <div
//           key={g.Id}
//           style={{
//             display: "flex",
//             justifyContent: "space-between",
//             alignItems: "center",
//             marginBottom: 6,
//             padding: "4px 8px",
//             border: "1px solid #ddd",
//             borderRadius: 4,
//           }}
//         >
//           <span>{g.Title}</span>
//           <PrimaryButton
//             text="Add"
//             onClick={async () => {
//               try {
//                 await sp.web.siteGroups
//                   .getByName(g.Title)
//                   .users.add(selectedUser.LoginName);

//                 setUserGroups((prev) => [...prev, g.Title]);
//               } catch (err) {
//                 console.error("Error adding user:", err);
//               }
//             }}
//           />
//         </div>
//       ))}
//   </div>

//   <div style={{ marginTop: 20 }}>
//     <DefaultButton text="Close" onClick={() => setIsModalOpen(false)} />
//   </div>
//     </div>
// </Modal> */}
// {/* <Modal isOpen={isModalOpen} onDismiss={() => setIsModalOpen(false)} >
//   <div style={{ width: "600px", minHeight: "400px", padding: "20px" }}>
//     <h3>Manage Permissions for {selectedUser?.Title}</h3>


//     <div style={{ marginBottom: "16px" }}>
//       <h4>Current Groups</h4>
//       {userGroups.length > 0 ? (
//         userGroups.map((g) => {
//           const groupObj = groups.find((grp) => grp.Title === g);
//           const permission =
//             groupObj?.Roles?.[0]?.RoleDefinitionBindings?.[0]?.Name || "N/A";

//           return (
//             <div
//               key={g}
//               style={{
//                 display: "flex",
//                 justifyContent: "space-between",
//                 alignItems: "center",
//                 marginBottom: 6,
//                 padding: "4px 8px",
//                 background: "#f3f2f1",
//                 borderRadius: 4,
//               }}
//             >
//               <span>
//                 <strong>{g}</strong> – <em>{permission}</em>
//               </span>
//               <DefaultButton
//                 text="Remove"
//                 onClick={async () => {
//                   try {
//                     await sp.web.siteGroups
//                       .getByName(g)
//                       .users.removeByLoginName(selectedUser.LoginName);

//                     setUserGroups((prev) => prev.filter((grp) => grp !== g));
//                   } catch (err) {
//                     console.error("Error removing user:", err);
//                   }
//                 }}
//               />
//             </div>
//           );
//         })
//       ) : (
//         <p style={{ fontStyle: "italic", color: "gray" }}>
//           User does not belong to any groups.
//         </p>
//       )}
//     </div>

//     <div style={{ marginBottom: "12px" }}>
//       <input
//         type="text"
//         placeholder="Search groups to add..."
//         style={{
//           width: "100%",
//           padding: "8px",
//           borderRadius: 4,
//           border: "1px solid #ccc",
//         }}
//         onChange={(e) => setSearchQuery(e.target.value)}
//       />
//     </div>


//     <div>
//       <h4>Available Groups</h4>
//       {groups
//         .filter(
//           (g) =>
//             !userGroups.includes(g.Title) &&
//             g.Title.toLowerCase().includes(searchQuery.toLowerCase())
//         )
//         .map((g) => {
//           const permission =
//             g?.Roles?.[0]?.RoleDefinitionBindings?.[0]?.Name || "N/A";

//           return (
//             <div
//               key={g.Id}
//               style={{
//                 display: "flex",
//                 justifyContent: "space-between",
//                 alignItems: "center",
//                 marginBottom: 6,
//                 padding: "4px 8px",
//                 border: "1px solid #ddd",
//                 borderRadius: 4,
//               }}
//             >
//               <span>
//                 <strong>{g.Title}</strong> – <em>{permission}</em>
//               </span>
//               <PrimaryButton
//                 text="Add"
//                 onClick={async () => {
//                   try {
//                     await sp.web.siteGroups
//                       .getByName(g.Title)
//                       .users.add(selectedUser.LoginName);

//                     setUserGroups((prev) => [...prev, g.Title]);
//                   } catch (err) {
//                     console.error("Error adding user:", err);
//                   }
//                 }}
//               />
//             </div>
//           );
//         })}
//     </div>

//     <div style={{ marginTop: 20 }}>
//       <DefaultButton text="Close" onClick={() => setIsModalOpen(false)} />
//     </div>
//   </div>
// </Modal> */}

// <Modal isOpen={isModalOpen} onDismiss={() => setIsModalOpen(false)}>
//   <div style={{ width: "800px", minHeight: "500px", padding: "20px" }}>
//     <div style={{ marginTop: 20, textAlign: "right" }}>
//       <DefaultButton text="Close" onClick={() => setIsModalOpen(false)} />
//     </div>
//     <h3>Manage Permissions for {selectedUser?.Title}</h3>

//     {/* Two-column layout */}
//     <div style={{ display: "flex", gap: "20px" }}>
//       {/* Left side - Current Groups */}
//       <div style={{ flex: 1 }}>
//         <h4>Current Groups</h4>
//         {userGroups.length > 0 ? (
//           userGroups.map((g) => {
//             const groupObj = groups.find((grp) => grp.Title === g);
//             const permission =
//               groupObj?.Roles?.[0]?.RoleDefinitionBindings?.[0]?.Name || "N/A";

//             return (
//               <div
//                 key={g}
//                 style={{
//                   display: "flex",
//                   justifyContent: "space-between",
//                   alignItems: "center",
//                   marginBottom: 6,
//                   padding: "4px 8px",
//                   background: "#f3f2f1",
//                   borderRadius: 4,
//                 }}
//               >
//                 <span>
//                   <strong>{g}</strong> 
//                 </span>
//                 <DefaultButton
//                   text="Remove"
//                   onClick={async () => {
//                     try {
//                       await sp.web.siteGroups
//                         .getByName(g)
//                         .users.removeByLoginName(selectedUser.LoginName);

//                       setUserGroups((prev) => prev.filter((grp) => grp !== g));
//                     } catch (err) {
//                       console.error("Error removing user:", err);
//                     }
//                   }}
//                 />
//               </div>
//             );
//           })
//         ) : (
//           <p style={{ fontStyle: "italic", color: "gray" }}>
//             User does not belong to any groups.
//           </p>
//         )}
//       </div>

//       {/* Right side - Available Groups */}
//       <div style={{ flex: 1 }}>
//         <h4>Available Groups</h4>
//         <div style={{ marginBottom: "12px" }}>
//           <input
//             type="text"
//             placeholder="Search By Site Name"
//             style={{
//               width: "100%",
//               padding: "8px",
//               borderRadius: 4,
//               border: "1px solid #ccc",
//             }}
//             onChange={(e) => setSearchQuery(e.target.value)}
//           />
//         </div>

//         {groups
//           .filter(
//             (g) =>
//               !userGroups.includes(g.Title) &&
//               g.Title.toLowerCase().includes(searchQuery.toLowerCase())
//           )
//           .map((g) => {
//             const permission =
//               g?.Roles?.[0]?.RoleDefinitionBindings?.[0]?.Name || "N/A";

//             return (
//               <div
//                 key={g.Id}
//                 style={{
//                   display: "flex",
//                   justifyContent: "space-between",
//                   alignItems: "center",
//                   marginBottom: 6,
//                   padding: "4px 8px",
//                   border: "1px solid #ddd",
//                   borderRadius: 4,
//                 }}
//               >
//                 <span>
//                   <strong>{g.Title}</strong> 
//                 </span>
//                 <PrimaryButton
//                   text="Add"
//                   style={{backgroundColor:"#7fc4de" , border:"none"}  }
//                   onClick={async () => {
//                     try {
//                       await sp.web.siteGroups
//                         .getByName(g.Title)
//                         .users.add(selectedUser.LoginName);

//                       setUserGroups((prev) => [...prev, g.Title]);
//                     } catch (err) {
//                       console.error("Error adding user:", err);
//                     }
//                   }}
//                 />
//               </div>
//             );
//           })}
//       </div>
//     </div>

    
//   </div>
// </Modal>


//     </div>
//   );
// }



// new component for user permission management in one go
// import * as React from "react";
// import { DefaultButton } from "@fluentui/react";
// import { Modal } from "@fluentui/react/lib/Modal";
// import { TextField, PrimaryButton, DetailsList, IColumn, IDetailsHeaderProps } from "@fluentui/react";


// import { SPFI } from "@pnp/sp";
// import { getSP } from "../loc/pnpjsConfig";
// import './Manageuserpermissioninonego.css'
// import "@pnp/sp/webs";
// import "@pnp/sp/folders";
// import "@pnp/sp/files";
// import "@pnp/sp/sites"
// import "@pnp/sp/presets/all"
// import "@pnp/sp/site-groups";

// export default function UserPermissionManager() {
//   const [users, setUsers] = React.useState<any[]>([]);
//   const [groups, setGroups] = React.useState<any[]>([]);
//   const [selectedUser, setSelectedUser] = React.useState<any>(null);
//   const [userGroups, setUserGroups] = React.useState<string[]>([]);
//   const [isModalOpen, setIsModalOpen] = React.useState(false);
//   const [loading, setLoading] = React.useState<boolean>(true);
//   const [searchQuery, setSearchQuery] = React.useState("");
//   const [searchText, setSearchText] = React.useState("");
//   const sp: SPFI = getSP();

//   React.useEffect(() => {
//     loadData();
//   }, []);

//   const loadData = async () => {
//     try {
//       setLoading(true);

//       const siteGroups = await sp.web.siteGroups.expand("Roles/RoleDefinitionBindings")();
//       console.log("Fetched Groups:", siteGroups);
      
//       const groupsWithPermissions = await Promise.all(
//         siteGroups.map(async (group) => {
//           try {
//             const roleAssignment: any = await sp.web.roleAssignments.filter(`PrincipalId eq ${group.Id}`)
//               .expand("Member,RoleDefinitionBindings")();

//             return {
//               ...group,
//               Permissions: roleAssignment.RoleDefinitionBindings.map(
//                 (r: any) => r.Name
//               ),
//             };
//           } catch (e) {
//             console.error(`Error fetching permissions for group: ${group.Title}`, e);
//             return { ...group, Permissions: [] };
//           }
//         })
//       );

//       console.log("Groups with Permissions:", groupsWithPermissions);
//       setGroups(groupsWithPermissions);

//       let userMap: any = {};
//       await Promise.all(
//         siteGroups.map(async (g) => {
//           const grpUsers = await sp.web.siteGroups.getById(g.Id).users();
//           grpUsers.forEach((u) => {
//             if (!userMap[u.LoginName])
//               userMap[u.LoginName] = { ...u, Groups: [] };
//             userMap[u.LoginName].Groups.push(g.Title);
//           });
//         })
//       );

//       setUsers(Object.values(userMap));
//     } catch (err) {
//       console.error("Error loading user/group data", err);
//     } finally {
//       setLoading(false);
//     }
//   };

//   const openManageModal = (user: any) => {
//     setSelectedUser(user);
//     setUserGroups(user.Groups);
//     setIsModalOpen(true);
//   };

//   const toggleGroup = (group: string, checked?: boolean) => {
//     if (checked) {
//       setUserGroups((prev) => [...prev, group]);
//     } else {
//       setUserGroups((prev) => prev.filter((g) => g !== group));
//     }
//   };

//   const saveChanges = async () => {
//     if (!selectedUser) return;
//     try {
//       setLoading(true);

//       const oldGroups = selectedUser.Groups;
//       const added = userGroups.filter((g) => !oldGroups.includes(g));
//       const removed = oldGroups.filter((g: any) => !userGroups.includes(g));

//       await Promise.all(
//         added.map((g) =>
//           sp.web.siteGroups.getByName(g).users.add(selectedUser.LoginName)
//         )
//       );
//       await Promise.all(
//         removed.map((g: any) =>
//           sp.web.siteGroups.getByName(g).users.removeByLoginName(
//             selectedUser.LoginName
//           )
//         )
//       );

//       setIsModalOpen(false);
//       loadData();
//     } catch (err) {
//       console.error("Error updating permissions", err);
//     } finally {
//       setLoading(false);
//     }
//   };

//   // Keep the original users list (no filtering)
//   const filteredUsers = users; // Remove the filter logic

//   const columns: IColumn[] = [
//     {
//       key: "user",
//       name: "User",
//       fieldName: "Title",
//       minWidth: 120,
//       maxWidth: 150,
//       isMultiline: false,
//     },
//     {
//       key: "email",
//       name: "Email",
//       fieldName: "Email",
//       minWidth: 200,
//       maxWidth: 250,
//       isMultiline: false,
//     },
//     {
//       key: "groups",
//       name: "Groups",
//       fieldName: "Groups",
//       minWidth: 400,
//       isMultiline: true,
//       onRender: (item: any) => item.Groups.join(", "),
//     },
//     {
//       key: "action",
//       name: "Action",
//       minWidth: 100,
//       maxWidth: 120,
//       onRender: (item: any) => (
//         <PrimaryButton
//           text="Manage"
//           style={{ backgroundColor: "#7fc4de", border: "none" }}
//           onClick={() => openManageModal(item)}
//         />
//       ),
//     },
//   ];

//   // Custom header with just the input field display (no filtering)
// const onRenderDetailsHeader = (props?: IDetailsHeaderProps, defaultRender?: any) => {
//   if (!props) return <></>;

//   // Render custom header
//   return (
//     <div style={{ display: "flex" }}>
//       {props.columns.map((col) => {
//         if (col.key === "user") {
//           return (
//             <div
//               key={col.key}
//               style={{ minWidth: col.minWidth, maxWidth: col.maxWidth, padding: "4px" }}
//             >
//               <span>{col.name}</span>
//               <TextField
//                 placeholder="Seach users..."
//                 styles={{ root: { marginTop: 4 } }}
//                 onChange={() => {}}
//               />
//             </div>
//           );
//         }

//         // Default column header
//         return (
//           <div
//             key={col.key}
//             style={{ minWidth: col.minWidth, maxWidth: col.maxWidth, padding: "4px" }}
//           >
//             {col.name}
//           </div>
//         );
//       })}
//     </div>
//   );
// };


//   return (
//     <div>
//       <h2>User Permissions</h2>

//       {loading ? (
//         <div className="loader-center">
//     <img
//       // src={require("../assets/ESSAROLLER.gif")}
//       src={require("../assets/ESSAROLLER2.gif")}
//       className="alignrightl"
//       alt="Loading..."
//     />
//     <span id="loader">Loading user permissions...</span>
//   </div>
//         // <div>
//         //   <img
//         //     src={require("../assets/ESSAROLLER.gif")}
//         //     className="alignrightl"
//         //     alt="Loading..."
//         //   />
//         //   <span id="loader">Loading user permissions...</span>
//         // </div>
//       ) : (
//         // <DetailsList
//         //   items={users} // Use original users array (no filtering)
//         //   columns={columns}
//         //   onRenderDetailsHeader={onRenderDetailsHeader}
//         // />
//         <DetailsList
//   items={users}
//   columns={columns}
//   onRenderDetailsHeader={onRenderDetailsHeader}
//   selectionMode={0}
// />
//       )}

//       <Modal isOpen={isModalOpen} onDismiss={() => setIsModalOpen(false)}>
//         <div style={{ width: "800px", minHeight: "500px", padding: "20px" }}>
//           <div style={{ marginTop: 20, textAlign: "right" }}>
//             <DefaultButton text="Close" onClick={() => setIsModalOpen(false)} />
//           </div>
//           <h3>Manage Permissions for {selectedUser?.Title}</h3>

//           <div style={{ display: "flex", gap: "20px" }}>
//             <div style={{ flex: 1 }}>
//               <h4>Current Groups</h4>
//               {userGroups.length > 0 ? (
//                 userGroups.map((g) => {
//                   const groupObj = groups.find((grp) => grp.Title === g);
//                   const permission =
//                     groupObj?.Roles?.[0]?.RoleDefinitionBindings?.[0]?.Name || "N/A";

//                   return (
//                     <div
//                       key={g}
//                       style={{
//                         display: "flex",
//                         justifyContent: "space-between",
//                         alignItems: "center",
//                         marginBottom: 6,
//                         padding: "4px 8px",
//                         background: "#f3f2f1",
//                         borderRadius: 4,
//                       }}
//                     >
//                       <span>
//                         <strong>{g}</strong> 
//                       </span>
//                       <DefaultButton
//                         text="Remove"
//                         onClick={async () => {
//                           try {
//                             await sp.web.siteGroups
//                               .getByName(g)
//                               .users.removeByLoginName(selectedUser.LoginName);

//                             setUserGroups((prev) => prev.filter((grp) => grp !== g));
//                           } catch (err) {
//                             console.error("Error removing user:", err);
//                           }
//                         }}
//                       />
//                     </div>
//                   );
//                 })
//               ) : (
//                 <p style={{ fontStyle: "italic", color: "gray" }}>
//                   User does not belong to any groups.
//                 </p>
//               )}
//             </div>

//             <div style={{ flex: 1 }}>
//               <h4>Available Groups</h4>
//               <div style={{ marginBottom: "12px" }}>
//                 <input
//                   type="text"
//                   placeholder="Search By Site Name"
//                   style={{
//                     width: "100%",
//                     padding: "8px",
//                     borderRadius: 4,
//                     border: "1px solid #ccc",
//                   }}
//                   onChange={(e) => setSearchQuery(e.target.value)}
//                 />
//               </div>

//               {groups
//                 .filter(
//                   (g) =>
//                     !userGroups.includes(g.Title) &&
//                     g.Title.toLowerCase().includes(searchQuery.toLowerCase())
//                 )
//                 .map((g) => {
//                   const permission =
//                     g?.Roles?.[0]?.RoleDefinitionBindings?.[0]?.Name || "N/A";

//                   return (
//                     <div
//                       key={g.Id}
//                       style={{
//                         display: "flex",
//                         justifyContent: "space-between",
//                         alignItems: "center",
//                         marginBottom: 6,
//                         padding: "4px 8px",
//                         border: "1px solid #ddd",
//                         borderRadius: 4,
//                       }}
//                     >
//                       <span>
//                         <strong>{g.Title}</strong> 
//                       </span>
//                       <PrimaryButton
//                         text="Add"
//                         style={{ backgroundColor: "#7fc4de", border: "none" }}
//                         onClick={async () => {
//                           try {
//                             await sp.web.siteGroups
//                               .getByName(g.Title)
//                               .users.add(selectedUser.LoginName);

//                             setUserGroups((prev) => [...prev, g.Title]);
//                           } catch (err) {
//                             console.error("Error adding user:", err);
//                           }
//                         }}
//                       />
//                     </div>
//                   );
//                 })}
//             </div>
//           </div>
//         </div>
//       </Modal>
//     </div>
//   );
// }

import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import { Modal } from "@fluentui/react/lib/Modal";
//import { TextField, PrimaryButton, DetailsList, IColumn, IDetailsHeaderProps, SelectionMode, CheckboxVisibility } from "@fluentui/react";  // Aman 8/04/2026
import { SPFI } from "@pnp/sp";
import { getSP } from "../loc/pnpjsConfig";
import './Manageuserpermissioninonego.css'
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/sites"
import "@pnp/sp/presets/all"
import "@pnp/sp/site-groups";
import Select from "react-select";
// srs 17/3/26
interface IUserPermissionManagerProps {
  sp: SPFI;
  siteCollections: { value: string; label: string; siteUrl: string; id: any }[];
  context: any;
}

//Define this outside your component
interface IGroupWithUsers {
  Id: number;
  Title: string;
  LoginName: string;
  Users?: { 
    Title: string; 
    Email: string; 
    LoginName: string; 
  }[];
}

// export default function UserPermissionManager() {
// srs 17/3/26
export default function UserPermissionManager(props: IUserPermissionManagerProps) {
  const [users, setUsers] = React.useState([]);
  const [groups, setGroups] = React.useState([]);
  const [selectedUser, setSelectedUser] = React.useState(null);
  const [userGroups, setUserGroups] = React.useState([]);
  const [isModalOpen, setIsModalOpen] = React.useState(false);
  const [loading, setLoading] = React.useState(true);
  const [searchQuery, setSearchQuery] = React.useState("");
  const [tableSearchText, setTableSearchText] = React.useState("");
  // Aman 8/04/2026
  const [currentPage, setCurrentPage] = React.useState(1);     
  const itemsPerPage = 10;
  // Ritik 03/04/2026
  const [selectedSiteForUserPermission, setSelectedSiteForUserPermission] = React.useState<any>(null);
const [activeSp, setActiveSp] = React.useState<SPFI>(props.sp);
// end here
  // const sp: SPFI = getSP();
  //srs 17/3/26
  const { sp, siteCollections, context } = props; // Ritik 03/04/2026
// Re-run when sp instance changes ritik 03/04/2026
  React.useEffect(() => {
  if (selectedSiteForUserPermission) {
    loadData();
  }
}, [activeSp]); // end here

  
  // const loadData = async () => {
  //   try {
  //     setLoading(true);
  //     const siteGroups = await activeSp.web.siteGroups.expand("Roles/RoleDefinitionBindings")(); // Ritik 03/04/2026
  //     console.log("Fetched Groups:", siteGroups);

  //     const groupsWithPermissions = await Promise.all(
  //       siteGroups.map(async (group) => {
  //         try {
  //           const roleAssignment: any = await activeSp.web.roleAssignments.filter(`PrincipalId eq ${group.Id}`) 
  //             .expand("Member,RoleDefinitionBindings")(); // Ritik 03/04/2026
  //           return {
  //             ...group,
  //             Permissions: roleAssignment.RoleDefinitionBindings.map(
  //               (r: any) => r.Name
  //             ),
  //           };
  //         } catch (e) {
  //           console.error(`Error fetching permissions for group: ${group.Title}`, e);
  //           return { ...group, Permissions: [] };
  //         }
  //       })
  //     );

  //     console.log("Groups with Permissions:", groupsWithPermissions);
  //     setGroups(groupsWithPermissions);

  //     let userMap: any = {};
  //     await Promise.all(
  //       siteGroups.map(async (g) => {
  //         const grpUsers = await activeSp.web.siteGroups.getById(g.Id).users(); // Ritik 03/04/2026
  //         grpUsers.forEach((u) => {
  //           if (!userMap[u.LoginName]) userMap[u.LoginName] = { ...u, Groups: [] };
  //           userMap[u.LoginName].Groups.push(g.Title);
  //         });
  //       })
  //     );

  //     setUsers(Object.values(userMap));
  //   } catch (err) {
  //     console.error("Error loading user/group data", err);
  //   } finally {
  //     setLoading(false);
  //   }
  // };
// Ritik 03/04/2026
  
 const loadData = async () => {

  try {

    setLoading(true);
 
    // Use <IGroupWithUsers[]> to tell TypeScript the 'Users' property exists

    const siteGroups: IGroupWithUsers[] = await activeSp.web.siteGroups

      .select("Id", "Title", "LoginName", "Users/Title", "Users/LoginName", "Users/Email")

      .expand("Users")();
 
    const roleAssignments: any[] = await activeSp.web.roleAssignments

      .expand("RoleDefinitionBindings")();
 
    const permissionMap = new Map<number, string[]>();

    roleAssignments.forEach((ra) => {

      if (ra.RoleDefinitionBindings) {

        permissionMap.set(

          ra.PrincipalId,

          ra.RoleDefinitionBindings.map((r: any) => r.Name)

        );

      }

    });
 
    const userMap: { [key: string]: any } = {};

    const groupsWithPermissions = siteGroups.map((group) => {

      const permissions = permissionMap.get(group.Id) || [];
 
      // Now 'group.Users' will not show a TypeScript error

      // if (group.Users && Array.isArray(group.Users)) {

      //   group.Users.forEach((u) => {

      //     if (!userMap[u.LoginName]) {

      //       userMap[u.LoginName] = { 

      //         Title: u.Title, 

      //         Email: u.Email, 

      //         LoginName: u.LoginName, 

      //         Groups: [] 

      //       };

      //     }

      //     if (!userMap[u.LoginName].Groups.includes(group.Title)) {

      //       userMap[u.LoginName].Groups.push(group.Title);

      //     }

      //   });

      // }

      if (group.Users && Array.isArray(group.Users)) {
        group.Users.forEach((u) => {
          const isSystemGroup = group.Title.includes("Limited Access System Group For List") || 
                               group.Title.includes("Limited Access System Group For Web") || 
                               group.Title.includes("SharingLinks");

          if (!isSystemGroup) {
            if (!userMap[u.LoginName]) {
              userMap[u.LoginName] = { 
                Title: u.Title, 
                Email: u.Email, 
                LoginName: u.LoginName, 
                Groups: [] 
              };
            }
            if (!userMap[u.LoginName].Groups.includes(group.Title)) {
              userMap[u.LoginName].Groups.push(group.Title);
            }
          }
        });
      }
 
      return {

        ...group,

        Permissions: permissions,

      };

    });
    
     // --- Aman 8/04/26: Final Filter Logic (100% Safe) ---
   const filteredGroupsList = groupsWithPermissions.filter(group => 
      !group.Title.includes("Limited Access System Group For List") && 
      !group.Title.includes("Limited Access System Group For Web") && 
      !group.Title.includes("SharingLinks")
    );

    const filteredUsersList = Object.values(userMap).filter((user: any) => 
      !user.Title.includes("SLinkClaim") && 
      !user.Title.includes("Limited Access System Group For List") &&
      !user.Title.includes("Limited Access System Group For Web") &&
      !user.Title.includes("SharingLinks")
    );

    setGroups(filteredGroupsList as any);
    setUsers(filteredUsersList as any);
    // ----------------------------------------------------

    // setGroups(groupsWithPermissions as any);

    // setUsers(Object.values(userMap) as any);
 
  } catch (err) {

    console.error("Error loading user/group data:", err);

  } finally {

    setLoading(false);

  }

};
 
 
  const handleSiteChange = async (selected: any) => {
    setSelectedSiteForUserPermission(selected);
    setCurrentPage(1);                                 // Aman 8/04/2026
    if (selected && selected.siteUrl) {
      const { spfi, SPFx } = await import("@pnp/sp");
      const targetSp = spfi(selected.siteUrl).using(SPFx(context));
      setActiveSp(targetSp);
    } else {
      setActiveSp(sp);
    }
  };
// end here
  const openManageModal = (user: any) => {
    setSelectedUser(user);
    setUserGroups(user.Groups);
    setIsModalOpen(true);
  };

  const removeAllGroups = async () => {
    if (!selectedUser || userGroups.length === 0) return;
    
    try {
      await Promise.all(
        userGroups.map((g) =>
          activeSp.web.siteGroups.getByName(g).users.removeByLoginName(selectedUser.LoginName)
        )
      );
      setUserGroups([]); 
      loadData(); // Refresh data to reflect changes Ritik 03/04/2026
    } catch (err) {
      console.error("Error removing all groups:", err);
    }
  };

  const saveChanges = async () => {
    if (!selectedUser) return;

    try {
      setLoading(true);
      const oldGroups = selectedUser.Groups;
      const added = userGroups.filter((g) => !oldGroups.includes(g));
      const removed = oldGroups.filter((g: any) => !userGroups.includes(g));

      await Promise.all(
        added.map((g) => activeSp.web.siteGroups.getByName(g).users.add(selectedUser.LoginName))
      );

      await Promise.all(
        removed.map((g: any) =>
          activeSp.web.siteGroups.getByName(g).users.removeByLoginName(selectedUser.LoginName)
        )
      );

      setIsModalOpen(false);
      loadData();
    } catch (err) {
      console.error("Error updating permissions", err);
    } finally {
      setLoading(false);
    }
  };

  // Filter users based on search
  const filteredUsers = React.useMemo(() => {
    if (!tableSearchText) return users;
    
    const searchLower = tableSearchText.toLowerCase();
    return users.filter((user: any) => {
      const userName = (user.Title || "").toLowerCase();
      const userEmail = (user.Email || "").toLowerCase();
      const userGroups = (user.Groups || []).join(", ").toLowerCase();
      
      return userName.includes(searchLower) || 
             userEmail.includes(searchLower) || 
             userGroups.includes(searchLower);
    });
  }, [users, tableSearchText]);
// Aman 8/04/26 (commented the code here as we are using custom table instead of details list)
  // const columns: IColumn[] = [
  //   {
  //     key: "user",
  //     name: "User",
  //     fieldName: "Title",
  //     minWidth: 150,
  //     maxWidth: 200,
  //     isResizable: true,
  //     isMultiline: false,
  //   },
  //   {
  //     key: "email",
  //     name: "Email",
  //     fieldName: "Email",
  //     minWidth: 200,
  //     maxWidth: 300,
  //     isResizable: true,
  //     isMultiline: false,
  //   },
  //   {
  //     key: "groups",
  //     name: "Group",
  //     fieldName: "Groups",
  //     minWidth: 300,
  //     isResizable: true,
  //     isMultiline: true,
  //     onRender: (item: any) => (
  //       <div style={{ padding: '8px 0' }}>
  //         {item.Groups.join(", ")}
  //       </div>
  //     ),
  //   },
  //   {
  //     key: "action",
  //     name: "Action",
  //     minWidth: 100,
  //     maxWidth: 120,
  //     isResizable: false,
  //     onRender: (item: any) => (
  //       <DefaultButton
  //         className="manage-button"
  //         text="Manage"
  //         onClick={() => openManageModal(item)}
  //       />
  //     ),
  //   },
  // ];

// Aman 8/04/26 start
  const totalPages = Math.ceil(filteredUsers.length / itemsPerPage);
  
  const handlePageChange = (pageNumber: any) => {
    if (pageNumber > 0 && pageNumber <= totalPages) {
      setCurrentPage(pageNumber);
    }
  };

  const startIndex = (currentPage - 1) * itemsPerPage;
  const currentData = filteredUsers.slice(startIndex, startIndex + itemsPerPage);

  React.useEffect(() => {
    setCurrentPage(1);
  }, [tableSearchText]);

  const Pagination = ({ currentPage, totalPages, handlePageChange }: any) => {
    const pageLimit = 5; 
    const startPage = Math.max(1, currentPage - Math.floor(pageLimit / 2));
    const endPage = Math.min(totalPages, startPage + pageLimit - 1);
    const adjustedStartPage = Math.max(1, Math.min(startPage, totalPages - pageLimit + 1));
    const visiblePages = Array.from(
      { length: Math.min(pageLimit, totalPages) },
      (_, index) => adjustedStartPage + index
    );

    return (
      <nav className="pagination-container mt-3">
        <ul className="pagination justify-content-end">
          <li className={`page-item ${currentPage === 1 ? 'disabled' : ''}`}>
            <a className="page-link" onClick={() => handlePageChange(currentPage - 1)} style={{cursor:'pointer'}}>«</a>
          </li>
          {visiblePages.map((pageNumber) => (
            <li key={pageNumber} className={`page-item ${currentPage === pageNumber ? 'active' : ''}`}>
              <a className="page-link" onClick={() => handlePageChange(pageNumber)} style={{cursor:'pointer'}}>{pageNumber}</a>
            </li>
          ))}
          <li className={`page-item ${currentPage === totalPages ? 'disabled' : ''}`}>
            <a className="page-link" onClick={() => handlePageChange(currentPage + 1)} style={{cursor:'pointer'}}>»</a>
          </li>
        </ul>
      </nav>
    );
  };

  //Aman 8/04/26 end
  return (
    <div className="user-permission-container">
      <div className="page-header1">
        {/* <h1>User Permissions</h1> */}
        <div className="page-header1">
        {/* ritik 03/04/2026 */}
        <div className="row mb-3">
          
          <div className="d-flex align-items-center justify-content-end gap-2">
            <label className="mb-0">Select Location</label>
            <Select  
              options={siteCollections}
              onChange={handleSiteChange}
              value={selectedSiteForUserPermission}
              placeholder="Select Location..."
              noOptionsMessage={() => "No Locations Found..."}
              styles={{
    container: (base) => ({
      ...base,
      width: "300px"
    })
  }}
            />
          </div>
        </div>
      </div>
{/* end here  */}
      </div>
      {/* Ritik 03/04/2026 */}
      {!selectedSiteForUserPermission ? (
  <div style={{ padding: '20px', color: '#888', textAlign: 'center' }}>
    Please select a location to view user permissions.
  </div>
) : loading ? (
  <div className="loader-center">
    <div>Loading user permissions...</div>
  </div>
) : (
  <div style={{border:'1px solid #ccc', padding:'15px', borderRadius:'8px', display:'grid', width:'100%', overflow:'auto', background:'#fff'}} className="table-wrapper1 mb-3">
    {/*end here */}
          <div className="d-flex align-items-center justify-content-between">
            <h3 className="font-18 text-dark fw-bold mb-2 m-0">User Permissions</h3> 
 <div style={{width:'600px'}} className="search-container">
          <input
            type="text"
            className="table-search-input"
            placeholder="Search by user name, email or group..."
            value={tableSearchText}
            onChange={(e) => setTableSearchText(e.target.value)}
          />
        </div>

            </div>
          {/* --- Aman 8/04/26 --- */}
<table className="mtbalenew">
  <thead>
    <tr>
      <th style={{ minWidth: '55px', maxWidth: '55px' }}>S.No.</th>
      <th>User</th>
      <th>Email</th>
      <th>Group</th>
      <th style={{ minWidth: '100px', maxWidth: '120px', textAlign: 'center' }}>Action</th>
    </tr>
  </thead>
  <tbody>
    {currentData && currentData.map((item: any, index: number) => (
      <tr key={index}>
        <td style={{ minWidth: '55px', maxWidth: '55px' }}>
          <span className="indexdesign">
            {startIndex + index + 1}
          </span>
        </td>
        <td>{item.Title || ''}</td>
        <td>{item.Email || ''}</td>
        <td 
  style={{ padding: '8px 0', fontSize: '13px', cursor: 'help' }} 
  title={item.Groups && item.Groups.join(", ")}
>
  {item.Groups && item.Groups.join(", ")}
</td>
        <td style={{ textAlign: 'center' }}>
          {/* Action: Calls the same existing modal function */}
          <button 
            type="button" 
            className="manage-button" 
            onClick={() => openManageModal(item)}
            style={{
              cursor: 'pointer',
              padding: '4px 12px',
              border: '1px solid #ccc',
              borderRadius: '4px',
              backgroundColor: '#fff'
            }}
          >
            Manage
          </button>
        </td>
      </tr>
    ))}
    {/* No Data State */}
    {(!filteredUsers || filteredUsers.length === 0) && (
      <tr>
        <td colSpan={5} style={{ textAlign: 'center', padding: '20px' }}>
          No users found.
        </td>
      </tr>
    )}
  </tbody>
</table>

<Pagination
          currentPage={currentPage}
          totalPages={totalPages}
          handlePageChange={handlePageChange}
        />
        </div>
      )}


      <Modal
        isOpen={isModalOpen}
        onDismiss={() => setIsModalOpen(false)}
        isBlocking={false}
      >
        <div className="modal-header">
          <h2>Manage Permissions for {selectedUser?.Title}</h2>
          {/* <button 
            className="close-button"
            onClick={() => setIsModalOpen(false)}
          >
            ×
          </button> */}
           <a style={{right:'45px'}}  className="close-button"   onClick={() => setIsModalOpen(false)}>
                                   
                                   
     <img style={{width:'13px'}}
       src={require("../assets/crosssvg.svg")}
       alt="Clear"
       className="search-icon"
     />
   </a>
        </div>

        <div className="modal-content">
          {/* Current Groups Section */}
          <div className="current-groups-header">
            <h3>Current Groups</h3>
            {userGroups.length > 0 && (
              <button className="remove-all-button" onClick={removeAllGroups}>
                <span><a >
                                   
                                   
     <img style={{width:'13px'}}
       src={require("../assets/crosssvg.svg")}
       alt="Clear"
       className="search-icon"
     />
   </a></span> Remove from All
              </button>
            )}
          </div>

          <div className="groups-container">
            {userGroups.length > 0 ? (
              userGroups.map((g) => (
                <div key={g} className="group-pill">
                  <span>{g}</span>
                  <button
                    className="remove-group-button"
                    onClick={async () => {
                      try {
                        await activeSp.web.siteGroups
                          .getByName(g)
                          .users.removeByLoginName(selectedUser.LoginName);
                        setUserGroups((prev) => prev.filter((grp) => grp !== g));
                        loadData(); // Refresh data to reflect changes Ritik 03/04/2026
                      } catch (err) {
                        console.error("Error removing user:", err);
                      }
                    }}
                  >
                    Remove <span><a >
                                   
                                   
     <img style={{width:'13px'}}
       src={require("../assets/crosssvg.svg")}
       alt="Clear"
       className="search-icon"
     />
   </a> </span>
                  </button>
                </div>
              ))
            ) : (
              <div className="no-groups-message">
                User does not belong to any groups.
              </div>
            )}
          </div>

          {/* Available Groups Section */}
          <div className="available-groups-section">
            <h3>Available Groups</h3>
            
            <div className="search-box">
              <input
                type="text"
                placeholder="Search By Site Name..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                
              />
                 <a style={{ position: "absolute", right: "28px" , top: "3px" , width:'14px' }} >
   
                          <img
                            src={require("../assets/searchicon.png")}
                            alt="Search"
                            className="search-icon"
                            style={{width: "31px"}}
                          />
                        </a>
            </div>

            <div className="available-groups-list">
              {groups
                .filter(
                  (g) =>
                    !userGroups.includes(g.Title) &&
                    g.Title.toLowerCase().includes(searchQuery.toLowerCase())
                )
                .map((g) => (
                  <div key={g.Title} className="available-group-pill">
                    <span>{g.Title}</span>
                    <button
                      className="add-group-button"
                      onClick={async () => {
                        try {
                          await activeSp.web.siteGroups
                            .getByName(g.Title)
                            .users.add(selectedUser.LoginName);
                          setUserGroups((prev) => [...prev, g.Title]);
                          loadData(); // Refresh data to reflect changes Ritik 03/04/2026
                        } catch (err) {
                          console.error("Error adding user:", err);
                        }
                      }}
                    >
                      Add <span><a >
                                   
                                   
     <img style={{width:'13px'}}
       src={require("../assets/addcircle.svg")}
       alt="Clear"
      
     />
   </a></span>
                    </button>
                  </div>
                ))}
            </div>
          </div>
        </div>
      </Modal>
    </div>
  );
}
