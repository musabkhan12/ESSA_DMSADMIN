// @ts-ignore
import * as React from "react";
import { useEffect , useState , useRef} from "react";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";

import * as XLSX from 'xlsx';
import './uploadfilecss'
import Swal from 'sweetalert2';


interface UploadFileProps {
  currentfolderpath: { [key: string]: string };
  onReturnToMain: () => void;
  // myRequest: (event: React.MouseEvent<HTMLButtonElement>) => void;
}

let previewURLN;

const UploadFile: React.FC<UploadFileProps> = ({ currentfolderpath , onReturnToMain  }) => {
  const sp: SPFI = getSP();

console.log('currentfolderpath' , currentfolderpath)

const [data, setData] = useState({
  Entity: '',
  Entityurl: '',
  siteID: '',
  Devision: '',
  Department: '',
  DocumentLibrary: '',
  Folder: '',
  folderpath: '',
});

const [state, setState] = useState({});

const currentUserEmailRef = useRef('');

const getcurrentuseremail = async()=>{
  const userdata = await sp.web.currentUser();
  currentUserEmailRef.current = userdata.Email;
  // console.log(currentUserEmailRef.current, "currentuser")
 }

useEffect(() => {
  getcurrentuseremail()
  setData({...data , ...currentfolderpath});

}, []);


console.log(data, "data"  )
console.log(data.Entity, "entity"  )
const SubsiteID = data.siteID
const currentPath =data.folderpath; 
const documentLibraryName  = data.DocumentLibrary;
console.log("documentLibraryName" , documentLibraryName)


  // const [libraryName, setLibraryName] = useState("Shared Documents"); // Set your library name here

  // let selectedFile:any=null;

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files![0];
    if (file) {
      // selectedFile=file;
      uploadFile(file);
    }
  };

  const uploadFile = async (file: File) => {
    try {
      const folder = sp.web.getFolderByServerRelativePath('DMSOrphanDocs');
      const uploadResult = await folder.files.addChunked(file.name, file);
      console.log("File uploaded successfully", uploadResult);

      // Generate the preview URL dynamically
      const previewUrl = await generatePreviewUrl(uploadResult.data.ServerRelativeUrl);
      
      previewFile(previewUrl);
    } catch (error) {
      console.error("Error uploading file:", error);
    }
  };

  const generatePreviewUrl = async (serverRelativeUrl: string) => {
    // Encode the file name and construct the preview URL
    const encodedFilePath = encodeURIComponent(serverRelativeUrl);
    
    // Example: 
    // serverRelativeUrl = "/sites/AlRostmani/test/DocumentLibraryInsideTest/Book.xlsx"
    const parentFolder = serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/'));
    const siteUrl = window.location.origin;

    const previewUrl = `${siteUrl}/sites/AlRostmani/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
    console.log("Generated Preview URL:", previewUrl);
   if(previewUrl){
    console.log("enter herr")
    const deletebut = document.getElementById('closeCommand') as HTMLElement
    if(deletebut){
      console.log(" here " , deletebut)
    }
   }
    return previewUrl;
  };


const previewFile = async (previewUrl: string) => {
    try {
      console.log("Previewing file at URL:", previewUrl);
      const iframe = document.getElementById("filePreview") as HTMLIFrameElement;
      const spinner = document.getElementById("spinner") as HTMLElement;
  
      // Show the spinner and hide the iframe initially
      spinner.style.display = "block";
      iframe.style.display = "none";
      iframe.src = previewUrl;
  
      // Add an onload event listener to the iframe
      iframe.onload = () => {
        console.log("Iframe has loaded");
  
        const checkAndHideButton = () => {
          try {
            const iframeDocument = iframe.contentDocument || iframe.contentWindow?.document;
            if (iframeDocument) {
              const button = iframeDocument.getElementById("OneUpCommandBar") as HTMLElement;
              const excelToolbar = iframeDocument.getElementById("m_excelEmbedRenderer_m_ewaEmbedViewerBar") as HTMLElement;
              if(excelToolbar){
                excelToolbar.style.display= "none"
              }
              if (button) {
                console.log("Hiding the OneUpCommandBar element");
                button.style.display = "none";
  
                // Hide the spinner and show the iframe after the button is hidden
                spinner.style.display = "none";
                iframe.style.display = "block"; 

               // Exit the loop once the button is found and hidden
              } else {
                console.log("OneUpCommandBar not found, rechecking...");
              }
              
              const helpbutton = iframeDocument.getElementById("m_excelEmbedRenderer_m_ewaEmbedViewerBar") as HTMLElement; 
              if(helpbutton){
                helpbutton.style.display = "none"
              }
            }
          } catch (error) {
            console.error("Error accessing iframe content:", error);
          }
  
          // Re-check after a short delay if the button wasn't found
          setTimeout(checkAndHideButton, 100);
        };
  
        // Start checking for the button
        checkAndHideButton();
      };
    } catch (error) {
      console.error("Error previewing file:", error);
    }

  };

  // const entity=data.Entity;


  
React.useEffect(()=>{
  const  loadFormOptions = async ()=> {
    try {
      // const testidsub = await sp.site.openWebById(data.siteID);
      
      // console.log("current Entity URL",testidsub.data.Url
      //   )
      
      // const fields = await testidsub.web.lists.getByTitle(`${documentLibraryName}`).fields.filter("Hidden eq false and ReadOnlyField eq false")();
      // console.log("Fields of document Library",fields);


      // const forms = await sp.web.lists.getByTitle("DMSPreviewformfields").items.select('*','SiteTitle/Title', 'SiteTitle/SiteURL').expand('SiteTitle')
      // .filter(`SiteTitle/Title eq '${propsDeatils.currentEntity}' `)();
      // console.log(forms, "forms");

      // start
      
      const documentLibraryFields=await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("ColumnName","ColumnType","IsRequired")
      .filter(
            `SiteName eq '${currentfolderpath.Entity}' 
            and DocumentLibraryName eq '${currentfolderpath.DocumentLibrary}' 
            and AddorRemoveThisColumn eq  'Add To Library' `)();

      console.log("Document Library Fields",documentLibraryFields);
      // end

      const formSelector = document.getElementById("formSelector");
      const submitButton=document.createElement('button');
      const uploadFileDiv=document.createElement('div');

      const createElement=(fieldName:string,type:string,required:boolean)=>{
            const inputContainer = document.createElement("div"); 
            inputContainer.className = "input-container";
    
            // Create and set label
            const label = document.createElement("label");
            label.setAttribute("htmlFor", fieldName);
            label.textContent = fieldName;
            inputContainer.appendChild(label);
    
            let inputElement: HTMLInputElement | null = null;
    
            // Dynamically create the input field based on FieldType
            let modifiedType = type.replace(/\s+/g, '').toLowerCase();
            console.log("modifiedType",modifiedType);

            if (
                modifiedType === "singlelineoftext"
                || 
                modifiedType === "multiplelineoftext" 
                || 
                modifiedType === 'text'
            ){
              inputElement = document.createElement("input");
              inputElement.type = "text";
            } else if (
              modifiedType === "number"
            ) {
              inputElement = document.createElement("input");
              inputElement.type = "number";
            } else if (
              modifiedType === "date&time"
            ) {
              inputElement = document.createElement("input");
              inputElement.type = "date";
            } else if (
              modifiedType === "yesorno"
            ) {
              inputElement = document.createElement("input");
              inputElement.type = "checkbox";
            }

            if (inputElement) {
              inputElement.className="dynamic-input";
              inputElement.id = fieldName;
              // inputElement.required = required.toLowerCase() === "yes"; 
              inputElement.required=required;
              inputContainer.appendChild(inputElement); 
              formSelector.appendChild(inputContainer); 
            }

            return;
      }


      // start
      documentLibraryFields.forEach((field)=>{
        createElement(field.ColumnName,field.ColumnType,field.IsRequired);
        })
      // end


      // properties of upload file div
      // uploadFileDiv.className="uploadfile";
      uploadFileDiv.className="input-container";

      // input for upload file
      const uploadFileInput=document.createElement('input');
      uploadFileInput.className="dynamic-input";
      uploadFileInput.type="file";
      uploadFileInput.id="fileInput";
      uploadFileInput.addEventListener('change', (event:any) => handleFileChange(event))

      // Set Label For upload file
      const label = document.createElement("label");
      label.setAttribute("htmlFor", 'fileInput');
      label.textContent = 'Upload File';

      uploadFileDiv.appendChild(label);
      uploadFileDiv.appendChild(uploadFileInput);
      formSelector.appendChild(uploadFileDiv);

      // Submit Button property
      submitButton.type="submit";
      submitButton.addEventListener('click',handleSubmit)
      submitButton.textContent="Submit"
      formSelector.appendChild(submitButton);
      
    } catch (error) {
      console.error("Error loading form options:", error);
    }
      }  

      loadFormOptions();
},[])
  


const handleSubmit = async (event: any) => {
  event.preventDefault();
  console.log("Button clicked");

  const formSelector = document.getElementById("formSelector") as HTMLFormElement;
  if (!formSelector.checkValidity()) {
      formSelector.reportValidity(); // Show validation errors
      return;
  }

  // Prepare the payload for SharePoint dynamically
  const inputs = document.querySelectorAll('.dynamic-input');
  // console.log("inputs",inputs)
  const payload: any = {};

  inputs.forEach((input) => {
      const inputElement = input as HTMLInputElement;
      const fieldName = inputElement.id;
      if (!fieldName) return; // Skip if field name is invalid

      if (inputElement.type === "checkbox") {
          payload[fieldName] = inputElement.checked;
      } else if (inputElement.type !== "file") {
          if(inputElement.value === ""){
             console.log("skip");
          }else{
            payload[fieldName] = inputElement.value;
          }
          
      }
  });

  const fileInput = document.getElementById('fileInput') as HTMLInputElement;
  const selectedFile = fileInput?.files?.[0]; 

  if (!selectedFile) {
      console.error("No file selected.");
      alert("Please select the file...");
      return;
  }

  try {
      console.log("Payload:", payload);
      console.log("SiteID:", currentfolderpath.siteID);

      const testidsub = await sp.site.openWebById(currentfolderpath.siteID);
      if (!testidsub) throw new Error("Subsite not found.");

      const documentLibraryInWhichWeUploadTheFile = testidsub.web.getFolderByServerRelativePath(currentfolderpath.folderpath);
      console.log("Current Path:", documentLibraryInWhichWeUploadTheFile);

      const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(selectedFile.name, selectedFile);
      console.log("File uploaded successfully", uploadResult.data.Name);
        (payload as any).Status = "Pending"
      const listItem = await uploadResult.file.getItem();
      
      const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
      const siteUrl = window.location.origin;
      const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl);
      console.log(encodedFilePath , "encodedFilePath")
      const previewUrl = `${siteUrl}/sites/AlRostmani/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
      
      console.log("Generated Preview URL:", previewUrl);
      if (!listItem) throw new Error("List item not found for the uploaded file.");

      await listItem.update(payload);
      console.log("File metadata updated successfully with:", payload);

      const newItem = await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add({
          FileName: String(uploadResult.data.Name),
          FileSize: String(uploadResult.data.Length),
          FileVersion: String(uploadResult.data.MajorVersion),
          CurrentFolderPath: String(currentfolderpath.folderpath),
          FileUID: String(uploadResult.data.UniqueId),
          CurrentUser: String(currentUserEmailRef.current),
          SiteID: String(currentfolderpath.siteID),
          Status: String("Pending"),
          FilePreviewURL : String(previewUrl),
          DocumentLibraryName:String(currentfolderpath.DocumentLibrary),
          SiteName : String(currentfolderpath.Entity)
      });
      console.log(newItem, "New item added FileMaster");


      const AddIteminDMSFileApprovalList = await sp.web.lists.getByTitle('DMSFileApprovalList').items.add({
        SiteName : String(currentfolderpath.Entity),  
         DocumentLibraryName : String(currentfolderpath.DocumentLibrary)  ,
         RequestedBy  : String(currentUserEmailRef.current) ,
         FileName: String(uploadResult.data.Name)   ,
         FileUID: String(uploadResult.data.UniqueId)    ,
         FilePreviewUrl: String(previewUrl),
         Status: String('Pending')    ,
         FolderPath : String(currentfolderpath.folderpath),
         ApproveAction : String('Submitted'),
         ApprovedLevel : String(1)
    })

    if(newItem && AddIteminDMSFileApprovalList){
      Deletemedia()
      setTimeout(() => {
        location.reload()
        onReturnToMain(); // Call onReturnToMain after 3 seconds
    }, 3000); // 3000 milliseconds = 3 seconds
     }

    console.log(AddIteminDMSFileApprovalList, "New item added to DMSFileApprovalList");

  }catch (error) {
      console.error("Error during submission:", error);
  }


};

const Deletemedia = () => {
   
  Swal.fire({
    title: "File Upload SuccessFull",
    text: "File Upload SuccessFull",
    icon: "success"
  });


 setTimeout(() => {
    Swal.close(); // Close the pop-up
    onReturnToMain(); // Call onReturnToMain if needed
  }, 3000); // 3000 milliseconds = 3 seconds

}

 

// const handleSubmit=async (event:any)=>{
          
//           event.preventDefault();
//           console.log("Button clicked");


//           const inputs = document.querySelectorAll('.dynamic-input');
//           const formSelector = document.getElementById("formSelector") as HTMLFormElement;
//           if (!formSelector.checkValidity()) {
//               formSelector.reportValidity(); // Show validation errors
//               return;
//             }

//           // Prepare the payload for SharePoint dynamically
//           const payload: any = {};
//           inputs.forEach((input) => {
//               const inputElement = input as HTMLInputElement;
//               const fieldName = inputElement.id; 

//               // Based on input type, store the correct value
//               if (inputElement.type === "checkbox") {
//                   payload[fieldName] = inputElement.checked;
//               }else if(fieldName === 'fileInput'){
//                   console.log("skip");
//               }else{
//                   payload[fieldName] = inputElement.value;
//               }

//           });
  


  

//           try {
//                   console.log("payload",payload);
//                   console.log("SiteID",data.siteID);
//                   const testidsub = await sp.site.openWebById(data.siteID);
//                   // console.log("subsite context",testidsub)
          
//                   const documentLibraryInWhichWeUploadTheFile = testidsub.web.getFolderByServerRelativePath(data.folderpath)
//                   console.log("Current Path",documentLibraryInWhichWeUploadTheFile)
                  
//                   const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(selectedFile.name, selectedFile);
//                   console.log("File uploaded successfully", uploadResult.data.Name);
                  
//                   await new Promise((resolve) => {
//                     console.log(uploadResult , "uploadResult")
//                     setTimeout(resolve, 1000)
//                   });

//                   // Fetch the associated list item for the uploaded file
//                   const listItem = await uploadResult.file.getItem();
//                   console.log("List item before upJdate:", listItem);

//                   // Update the metadata on the file's list item
//                   await listItem.update(payload);
//                   console.log("File metadata updated successfully with:", payload);
//                   console.log(SubsiteID , "SubsiteID")
                  
//                   const newItem = await sp.web.lists.getByTitle(`DMS${data.Entity}FileMaster`).items.add({
              
//                     FileName: String(uploadResult.data.Name),          // Example: Set the Title field
//                     FileSize: String(uploadResult.data.Length),    // Example: Set the FileSize field
//                     FileVersion: String(uploadResult.data.MajorVersion),// Example: Set the FileVersion field
//                     CurrentFolderPath : String(currentPath),
//                     FileUID: String(uploadResult.data.UniqueId),// Example: Set the FileVersion field
//                     CurrentUser : String(currentUserEmailRef.current),
//                     SiteID : String(SubsiteID),
//                     Status:String("Pending")
//             });
//                     console.log(newItem, "Today")
//           } catch (error) {
//                   console.log("Error From Adding Field Name",error);
//           } 
// }

    // const ArgPoc=()=>{
      
    // }
    
    return (
      <>
          <button className='BackButton' 
          // onClick={(event) => {
          //       onReturnToMain();
          //       // myRequest(event);
          // }}
          onClick={onReturnToMain} 
          > Back 
          </button>
          <div className="container mt-4 UploadFileCont">
              <div className='main-containeruploadfile'>
                      <div className='column column1 p-3'>
                          <form id='formSelector'>
                              <h1>Upload file</h1>
                              {/* <div className="uploadfile">
                                  <input type="file" id="fileInput" onChange={handleFileChange} />
                              </div> */}
                          </form>
                      </div>
                      <div className='column column2 p-3'>
                          <h1>File Preview</h1>
                          <div id="spinner" style={{display: "none"}}>Loading...</div>
                          <iframe id="filePreview" width="100%" height="400"></iframe>
                      </div>
              </div>
          </div>
      </>
      // <div className="uploadfile">
      //   <input type="file" id="fileInput" onChange={handleFileChange} />
      //   <div id="spinner" style={{display: "none"}}>Loading...</div>
      //   <iframe id="filePreview" width="100%" height="500"></iframe>

      //   <div id="fileMetadata"></div>
      // </div>
    );
  }
export default UploadFile;