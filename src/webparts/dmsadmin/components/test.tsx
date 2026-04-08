
  // @ts-ignore
  import * as React from "react";
  import { getSP } from "../loc/pnpjsConfig";
  import { SPFI } from "@pnp/sp";
  import "bootstrap/dist/css/bootstrap.min.css";
  import Swal from "sweetalert2";
  // import "bootstrap//dist/"
  import "../../../CustomCss/mainCustom.scss";
  // import "../../verticalSideBar/components/VerticalSidebar.scss";

  // import { useState , useEffect } from "react";
  import Provider from "../../../GlobalContext/provider";
  import { useMediaQuery } from "react-responsive";
  import "@pnp/sp/webs";
  import "@pnp/sp/folders";
  import "@pnp/sp/files";
  import { IFileInfo } from "@pnp/sp/presets/all";
  import { IFile } from '@pnp/sp/files';
  import "@pnp/sp/sites"
  import "@pnp/sp/presets/all"
  import "@pnp/sp/webs";
  import "@pnp/sp/sites";
  import "@pnp/sp/site-users/web";
  import { PermissionKind } from "@pnp/sp/security";
  import "bootstrap/dist/css/bootstrap.min.css";
  import "../../../CustomCss/mainCustom.scss";
   import "../../verticalSideBar/components/VerticalSidebar.scss";
  import "./dmscss.css";
  import { useState , useRef , useEffect} from "react";
  import { RenderListDataOptions } from "@pnp/sp/lists";
  
  import {IDmsMusaibProps} from './IDmsMusaibProps'

import "./testcss.css";
  
  const Test = ({ props }: any) => {
    const sp: SPFI = getSP();
 // Define proper interface that includes list item fields
 interface IFileBasicInfo {
  Name: string;
  Length: number;
  ServerRelativeUrl: string;
  UniqueId: string;
}
async function getFilesBypassThreshold(
  siteId: string,
  folderServerRelativeUrl: string
) {
  try {
    const site = await sp.site.openWebById(siteId);
    const web = site.web;

    // Get the list ID from folder metadata
    const folder = web.getFolderByServerRelativePath(folderServerRelativeUrl);
    const folderInfo: any = await folder
      .expand("ListItemAllFields", "ListItemAllFields/ParentList")
      .select("ListItemAllFields/ParentList/Id")();

    const listId = folderInfo?.ListItemAllFields?.ParentList?.Id;
    const list = web.lists.getById(listId); // FIXED: use list object, not just the ID

    const renderOptions = {
      ViewXml: `<View Scope='RecursiveAll'><Query></Query><RowLimit Paged='TRUE'>5000</RowLimit></View>`,
      FolderServerRelativeUrl: folderServerRelativeUrl,
    };

    let allItems: any[] = [];
    let position = null;

    do {
      const renderData: any = await list.renderListDataAsStream({
        ...renderOptions,
        Paging: position ? JSON.stringify({ Paged: "TRUE", p_ID: position }) : undefined,
      });

      allItems.push(...renderData.Row);
      position = renderData.NextHref
        ? new URLSearchParams(renderData.NextHref).get("p_ID")
        : null;
    } while (position);
    console.log("✅ Total Items Fetched:", allItems);
    console.log("✅ Total Items Fetched:", allItems.length);
    return allItems;
  } catch (err) {
    console.error("❌ Error fetching files:", err);
    return [];
  }
}
// Call with:
getFilesBypassThreshold(
  "fb84b27e-1841-4114-8bef-1bd6c19cde19",
  "/sites/Intranet/Group Information Technology Department/Archived Files/ARG_BO/EEE/Documents"
);
// async function getFilesBypassThreshold(
//   siteId: string,
//   folderServerRelativeUrl: string
// ): Promise<IFileWithListItem[]> {
//   try {
//       const web = (await sp.site.openWebById(siteId)).web;
//       const allFiles: IFileWithListItem[] = [];
      
//       // 1. First get the folder's item count
//       const folder = await web.getFolderByServerRelativePath(folderServerRelativeUrl);
//       const folderItem = await folder.listItemAllFields();
//       const itemCount = folderItem.ItemCount || 0;

//       // 2. If under threshold, get all at once
//       if (itemCount <= 5000) {
//           return folder.files.expand("listItemAllFields")
//               .select("*, listItemAllFields/Id, listItemAllFields/FileRef")();
//       }

//       // 3. For large folders, use date-based chunking
//       const dateRanges = await getDateRanges(web, folderServerRelativeUrl);
      
//       for (const range of dateRanges) {
//           const files = await web.getFolderByServerRelativePath(folderServerRelativeUrl)
//               .files
//               .expand("listItemAllFields")
//               .select("*, listItemAllFields/Id, listItemAllFields/FileRef")
//               .filter(`Created ge datetime'${range.start}' and Created le datetime'${range.end}'`)
//               .top(1000)
//               .orderBy("Created", true)() as IFileWithListItem[];
          
//           allFiles.push(...files);
//           await new Promise(resolve => setTimeout(resolve, 300)); // Throttle
//       }

//       return allFiles;
//   } catch (error) {
//       console.error("Error in getFilesBypassThreshold:", error);
//       throw error;
//   }
// }
// getFilesBypassThreshold("fb84b27e-1841-4114-8bef-1bd6c19cde19", "/sites/Intranet/Group Information Technology Department/Archived Files/ARG_BO/EEE/Documents")
// async function getDateRanges(web: any, folderPath: string): Promise<{start: string, end: string}[]> {
//   // Get min and max dates from the folder
//   const result = await web.getFolderByServerRelativePath(folderPath)
//       .files
//       .select("Created")
//       .top(1)
//       .orderBy("Created", true)();
  
//   const minDate = new Date(result[0].Created);
//   const maxDate = new Date();
  
//   // Create monthly chunks
//   const ranges = [];
//   let currentStart = minDate;
  
//   while (currentStart < maxDate) {
//       const currentEnd = new Date(currentStart);
//       currentEnd.setMonth(currentEnd.getMonth() + 1);
      
//       ranges.push({
//           start: currentStart.toISOString(),
//           end: currentEnd.toISOString()
//       });
      
//       currentStart = new Date(currentEnd);
//       currentStart.setDate(currentStart.getDate() + 1);
//   }
//   return ranges;
// }

// async function getAllFilesThresholdSafe(
//   siteId: string,
//   folderServerRelativeUrl: string
// ): Promise<IFileWithListItem[]> {
//   try {
//       const web = (await sp.site.openWebById(siteId)).web;
//       const allFiles: IFileWithListItem[] = [];
//       let lastId = 0;
//       let hasMore = true;
//       const batchSize = 1000; // Conservative batch size

//       while (hasMore) {
//           const filter = lastId > 0 ? `ListItemAllFields/Id gt ${lastId}` : ``;
          
//           const files = await web.getFolderByServerRelativePath(folderServerRelativeUrl)
//               .files
//               .expand("ListItemAllFields") // Note: CamelCase in query but lowercase in response
//               .select("*, ListItemAllFields/Id")
//               .filter(filter)
//               .top(batchSize)
//               .orderBy("ListItemAllFields/Id", true)() as IFileWithListItem[];
//           console.log("files all get", files);
//           alert("files all get"+ files);
//           if (files.length > 0) {
//               allFiles.push(...files);
//               lastId = files[files.length - 1].listItemAllFields.Id; // Note: lowercase in response
//           } else {
//               hasMore = false;
//           }

//           // Small delay between batches to avoid throttling
//           if (hasMore) {
//               await new Promise(resolve => setTimeout(resolve, 500));
//           }
//       }

//       return allFiles;
//   } catch (error) {
//       console.error("Error in getAllFilesThresholdSafe:", error);
//       throw error;
//   }
// }
// getAllFilesThresholdSafe("fb84b27e-1841-4114-8bef-1bd6c19cde19", "/sites/Intranet/Group Information Technology Department/Archived Files/ARG_BO/EEE/Documents")
    // async function getAllFilesFromFolderWithPaging(
    //   siteId: string,
    //   folderServerRelativeUrl: string
    // ): Promise<any[]> {
    //   try {
    //     // Initialize the web using the site ID
    //     const testidsub = sp.site.openWebById(siteId);
    
    //     // Get all files from the folder with paging to avoid threshold
    //     const allFiles: any[] = [];
    //     let files = await (await testidsub).web.getFolderByServerRelativePath(folderServerRelativeUrl)
    //     .files
    //     .expand("ListItemAllFields") // This is key to include the list item
    //     .select("*, ListItemAllFields/Id")
    //     .top(2000)
    //     .orderBy("ListItemAllFields/Id", true)();
    //      console.log("files all get", files);
    //     while (files.length > 0) {
    //       allFiles.push(...files);
          
    //       // Get the next batch using the ID of the last item in the current batch
    //       const lastId = files[files.length - 1].ListItemAllFields.Id;
    //       files = await (await testidsub).web.getFolderByServerRelativePath(folderServerRelativeUrl)
    //         .files.select("*, ListItemAllFields")
    //         .top(2000)
    //         .filter(`ListItemAllFields/Id gt ${lastId}`)
    //         .orderBy("ListItemAllFields/Id")();
    //     }
    
    //     return allFiles;
    //   } catch (error) {
    //     console.error("Error fetching files:", error);
    //     throw error;
    //   }
    // }
    
    // // Usage example
    // const siteId = "fb84b27e-1841-4114-8bef-1bd6c19cde19";
    // const folderPath = "/sites/Intranet/Group Information Technology Department/Archived Files/ARG_BO/EEE/Documents";
    
    // getAllFilesFromFolderWithPaging(siteId, folderPath)
    //   .then(files => {
    //     console.log(`Retrieved ${files.length} files`);
    //     // Process your files here
    //   })
    //   .catch(error => {
    //     console.error("Error:", error);
    //   });
  
    return (
      <div id="wrapper" >
          <p>Hello</p>

          <p>Test</p>
          <p>Test</p>
    
            </div>
          
     
          
    );
  };
  
  
  
  const TEST: React.FC<IDmsMusaibProps> = (props) =>{
    return (
      <Provider>
        <Test  props={props}/>
      </Provider>
    );
  };
  
  export default TEST;
  