import * as React from 'react';
import Provider from '../../../GlobalContext/provider';
import  styles from './BasicForm.module.scss'
import classNames from "classnames";
import { getSP } from '../loc/pnpjsConfig';

import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/sites"
import "@pnp/sp/presets/all"
import "@pnp/sp/site-groups";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
// import { SPFI,spfi } from '@pnp/sp';
import { spfi, SPFx, SPFI } from "@pnp/sp";

import { useState} from "react";
import { EnvironmentType } from '@microsoft/sp-core-library';
// import { BasicFormProps }  from './IDmsMusaibProps'
import Swal from 'sweetalert2';
import { Context } from 'react-responsive';
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface BasicFormProps {
  sp: SPFI;
  context: WebPartContext;

  currentId: any;
  currentJobTitle: any;
  currentIsActive: any;
  onCancel: any;
  IsExternal: any;

  selectedSiteFilter?: string;
  siteCollections?: any[];

  setSelectedSiteFilter?: React.Dispatch<React.SetStateAction<string>>;
}
// interface BasicFormProps {
//       sp: SPFI;
//     currentId :any,
//     currentJobTitle:any,
//     currentIsActive:any,
//     onCancel:any,
//     IsExternal:any,
//     selectedSiteFilter?: string;
//     siteCollections?: any[];
// }

let currentusername="";

const Basic: React.FC<BasicFormProps> = ({
    sp,
    context, 
    selectedSiteFilter,
     setSelectedSiteFilter,
    siteCollections,
    currentId,
    currentJobTitle,
    currentIsActive,
    onCancel,
    IsExternal
})=>{

    //const sp: SPFI = getSP();
    // console.log(sp);
    console.log("currentId",currentId);
    const [jobTitle, setJobTitle] = useState(currentJobTitle || '');
    const [isActive, setIsActive] = useState(currentIsActive || '');
    const [isExternal, setIsExternal] = useState(IsExternal || '');
    const [update, setUpdate] = useState(currentId ? true : false);
    React.useEffect(() => {
    setIsActive(currentIsActive || '');
    setIsExternal(IsExternal || '');
}, [currentIsActive, IsExternal]);

    const [description,setDescription] = useState('');

    // Aman 21/4/26 start
const [formErrors, setFormErrors] = useState({
  location: false,
  title: false,
  isActive: false,
  isExternal: false,
  description: false
});
    // Aman 21/4/26 end
    // const [errorMessage, setErrorMessage] = useState<string | null>(null);
    const [disableInput, setDisableInput]=useState(false);

    React.useEffect(()=>{
        setDisableInput(true);
    },[]);

    if(currentId !== null && disableInput){
                const title=document.getElementById("jobTitle") as HTMLInputElement;;
                const description=document.getElementById("description") as HTMLInputElement;;
                console.log("title",title);
                console.log("description",description);
                title.disabled = true;
                description.disabled = true;
    }


    // Handle form submission 
    // Handle form submission 
//     const handleSubmit = async (event: any) => {
//         if (!selectedSiteFilter) {
//   Swal.fire("Please select a Site Collection");
//   return;
// }

//         // const url = window.location.origin + window.location.pathname;
//         const url = window.location.href;

//                 console.log("url addhyan",  url)
//       let targetSp = sp;

// if (selectedSiteFilter) {
//   targetSp = spfi(selectedSiteFilter).using(SPFx(context));
// }
//         //const siteUrlnew = sp.web.toUrl().replace('/_api/web', '');
//         const siteUrlnew = targetSp.web.toUrl().replace('/_api/web', '');
// console.log(siteUrlnew, "siteurlnew");

//        const sitecollection = siteUrlnew.split('/sites/')[1].split('/')[0];


const handleSubmit = async (event: any) => {
        if (!currentId && !selectedSiteFilter) {
            // Aman 21/4/26 start
            setFormErrors(prev => ({
                ...prev,
                location: true
            }));
            // Aman 21/4/26 end 
          Swal.fire("Please select a Site Collection");
          return;
        }

        const url = window.location.href;
        console.log("url addhyan", url)
// ritik 06/04
       let targetSp = sp;
if (selectedSiteFilter) {
  targetSp = spfi(selectedSiteFilter).using(SPFx(context));
}

        let siteUrlnew = '';
        let sitecollection = '';
        if (!currentId) {
          siteUrlnew = targetSp.web.toUrl().replace('/_api/web', '');
          console.log(siteUrlnew, "siteurlnew");
          sitecollection = siteUrlnew.includes('/sites/')
            ? siteUrlnew.split('/sites/')[1].split('/')[0]
            : '';
        }

//         event.preventDefault(); 
//         const form=document.getElementById('createMaster') as HTMLFormElement
//         if (!form.checkValidity()) {
//             // form.reportValidity(); // Show validation errors
//             // Aman 21/4/26 start

//             const errors = {
//     location: !currentId && !selectedSiteFilter,
//     title: !jobTitle,
//     isActive: !isActive,
//     isExternal: !isExternal,
//     description: !description
//   };

//   setFormErrors(errors);

//   Swal.fire(
//   "Please fill out the fields!",
//   "All fields are required"
// );

//   return;
//             // checkValidation();
//             // return;
//             // Aman 21/4/26 end
//         }
event.preventDefault(); 
        const form = document.getElementById('createMaster') as HTMLFormElement;

        // Validation Logic
        const errors = {
          // puja kumari 22/04/2026
          location: !currentId && (!selectedSiteFilter || selectedSiteFilter.trim() === "" || selectedSiteFilter === "ALL"),
            // location: !currentId && (!selectedSiteFilter || selectedSiteFilter === ""),
            //end
            title: !jobTitle.trim(),
            isActive: !isActive,
            isExternal: !isExternal,
            description: !description.trim()
        };

        setFormErrors(errors);

        
        if (!form.checkValidity() || errors.location) {
            Swal.fire("Please fill out the fields!", "All fields are required");
            return; 
        }

        // Remove alphanumeric characters and also check the limit start
        const nonAlphaNumericForEntity = jobTitle.replace(/[^a-zA-Z0-9 -]/g, '');
        // const nonAlphaNumericDescription = description.replace(/[^a-zA-Z0-9 ]/g, '');
        if (jobTitle !== nonAlphaNumericForEntity) {
            console.log('Special characters are not allowed in entity.');
            checkLimitAndAlphanumericCharacter('Special characters are not allowed in the Entity field.');
            return;
        } else if (nonAlphaNumericForEntity.length > 50) {
            console.log('Input cannot exceed 30 characters.');
            checkLimitAndAlphanumericCharacter('Input cannot exceed 50 characters in the Entity field.')
            return;
        }

        // if (description !== nonAlphaNumericDescription) {
        //     console.log('Alphanumeric characters are not allowed in entity.');
        //     checkLimitAndAlphanumericCharacter('Alphanumeric characters are not allowed in the Description field.')
        //     return;
        // } else 
        if (description && description.length > 250) {
            console.log("Description",description);
            console.log('Input cannot exceed 250 characters.');
            checkLimitAndAlphanumericCharacter('Input cannot exceed 250 characters in the  Description field.')
            return;
        }

         // Check for only white spaces start 
         const isOnlyWhitespace = (jobTitle:string) => jobTitle.trim() === "";
         console.log("isOnlyWhitespace",isOnlyWhitespace)
         if(isOnlyWhitespace(jobTitle)){
             console.log("White spaces only");
             checkWhiteSpaces()
             return;
         }
         // End
        // end

        const newItem = {
            Title: jobTitle, 
            Active: isActive,
            Description:description ,
            IsExternal:isExternal
        };
         const newItemforEssa = {
            Title: jobTitle, 
            Active: isActive,
            Description:description ,
            IsExternal:isExternal,
            SiteURL: `${siteUrlnew}/${jobTitle}`,
            FileMasterList: `DMS${jobTitle}FileMaster`
        };

        const newItemforEssaAllsite = {
            
            // Isactive: "Yes",
            // Isactive:"Yes",
            Siteurl: `${siteUrlnew}/${jobTitle}`,
            FileMaster: `DMS${jobTitle}FileMaster`,
            sitecollection: sitecollection,
        };





        console.log(newItem);

        const listTitle='MasterSiteURL';

        try {

            if(currentId){
                let alreadyExist=false;
                //const entity=await sp.web.lists.getByTitle(listTitle).items.getAll();
                const entity = await targetSp.web.lists.getByTitle(listTitle).items.getAll();
                console.log("entity",entity);
               entity.forEach((e)=>{
    if(e.Id !== currentId && e.Title !== null){
        if(e.Title.replace(/\s+/g, '').toLowerCase() === jobTitle.replace(/\s+/g, '').toLowerCase()){
            if(e.Active === isActive && e.IsExternal === isExternal){
                alreadyExistValue1(e.Active);
                alreadyExist=true;
                return;
            }
        }
    }
})

                if(!alreadyExist){
                    console.log("Edit Entity Id",currentId);
                    // Update item in actual site collection
await targetSp.web.lists.getByTitle(listTitle).items.getById(currentId).update({
  Active: isActive,
  IsExternal: isExternal
});

// ESSA sync
try {
//   const essaUrl = "https://officeindia.sharepoint.com/sites/ESSA/";
    const essaUrl = `${window.location.href.split('/sites/')[0]}/sites/ESSA/`;
  const essaSP = spfi(essaUrl).using(SPFx(context));  // context use karo, (sp as any)._context nahi
  const essaItems = await essaSP.web.lists
    .getByTitle("MasterSiteURL")
    .items
    .filter(`Title eq '${jobTitle}'`)();
  if (essaItems.length > 0) {
    await essaSP.web.lists
      .getByTitle("MasterSiteURL")
      .items.getById(essaItems[0].Id)
      .update({ Active: isActive, IsExternal: isExternal });
    console.log("ESSA MasterSiteURL updated ✅");
  }
} catch(essaErr) {
  console.error("ESSA update failed:", essaErr);
}
    // end here
                    updateValue(jobTitle);
                    // alert('Division updated successfully');
                    // Back to entity table start
                    setTimeout(()=>{
                        onCancel();
                    },1000)
                    // end
                }
            }else{

            // Check if already exist
            let alreadyExist=false;
            //const entity=await sp.web.lists.getByTitle(listTitle).items.getAll();
            const entity = await targetSp.web.lists.getByTitle(listTitle).items.getAll();
            // console.log("entity",entity);

            entity.forEach((e)=>{
                if(e.Title !== null){
                    if(e.Title.replace(/\s+/g, '').toLowerCase() === jobTitle.replace(/\s+/g, '').toLowerCase()){
                        // alert(`${jobTitle} Already exist,`);
                        // setErrorMessage(`${jobTitle} already exists.`);
                        alreadyExistValue(jobTitle);
                        alreadyExist=true;
                        return;
                    }
            }
            })

            //Create New Entity
            if(!alreadyExist){
                // setErrorMessage(null);
                //const data=await sp.web.lists.getByTitle(listTitle).items.add(newItem);
                const data = await targetSp.web.lists.getByTitle(listTitle).items.add(newItem);
                console.log("Testing");
                
                //Addhyan - 12/03/2026 jab sitecollect me entity create hoga tab essa ke mastersite url me v data jaega 

                const fullUrl = window.location.href;
                const url = fullUrl.split("/_layouts")[0];
                console.log("url addhyan",  url)
                const essa ="https://officeindia.sharepoint.com/sites/ESSA/"
                const otherSiteSP = spfi(essa).using(SPFx((sp as any)._context));
                const otherSiteSPs = spfi(essa).using(SPFx((sp as any)._context));
                const listTitles = "MasterSiteURL";
                const listTitle2 = "Allsitesfilemaster"

                    const entity = await otherSiteSP.web.lists
                        .getByTitle(listTitles).items.add(newItemforEssa);

                    const entity2 = otherSiteSPs.web.lists
                        .getByTitle(listTitle2).items.add(newItemforEssaAllsite);

                        console.log("Allsitefilemaster",entity2 )

                    console.log("mastersiteurl", entity);


    //addhyan -12/03/2026 -end


















                // alert(`${jobTitle} added successfully`);
                onSuccess(jobTitle);
                clearForm();
                setTimeout(()=>{
                            onCancel();
                    },1000)
            }
        }

        










        } catch (error) {
            console.error('Error adding item:', error);
            alert('Error adding item');
        }

    };

    const clearForm=()=>{
        setJobTitle("");
        setIsActive("");
        setIsExternal(""); // Aman 21/4/26
        setDescription("");
        // Aman 21/4/26 start
        setFormErrors({
      location: false,
      title: false,
      isActive: false,
      isExternal: false,
      description: false
    });
    // Aman 21/4/26 end
    }

    const checkValidation=()=>{
        Swal.fire("Please fill out the fields!", "All fields are required");
    }

    const alreadyExistValue=(jobTitle:any)=>{
        Swal.fire(`${jobTitle} Exist`, "Please change the entity name", "warning");
    }

    const alreadyExistValue1=(status:any)=>{
        // if()
        // console.log("status",status);
        let st=""
        if(status === "No"){
            st="Inactive"
        }else if(status === "Yes"){
            st="Active"
        }
        Swal.fire(`Please update the status`, `The entity is already ${st}.`, "warning");
    }

    const onSuccess=(jobTitle:any)=>{
        // ritik 21/4/26 start
        //Swal.fire(`${jobTitle} is creating, This will reflect shortly in DMS`,"", "success");
         Swal.fire({
                    title: "Saved successfully.",
                    
                    icon: "success"
                  });
        // ritik 21/4/26 end
    }

    const updateValue=(jobTitle:any)=>{
        // ritik 21/4/26 start
        //Swal.fire(`${jobTitle} Updated`,"", "success");
         Swal.fire({
                    title: "Submitted successfully.",
                    
                    icon: "success"
                  });
       // ritik 21/4/26 end
    }

    // new function added for check limit and alphanumeric character start
    const checkLimitAndAlphanumericCharacter=(message:any)=>{
        Swal.fire(`Invalid Input`,`${message}`, "warning");
    }

    const checkWhiteSpaces=()=>{
        Swal.fire(`Invalid Input`,`Only spaces are not allowed`, "warning");
    }
    // end

  return (
        
    <>  
      <div style={{marginTop:'10px'}} className={styles.DmsAdminForm}>
        <div className={styles.formcontainer}>            
            {/* <div className={styles.apphier}>
                <h1 className={styles.apptitle}>Create Entity</h1>
            </div> */}
       
            <form id="createMaster" onSubmit={handleSubmit}>
                <div className="p-3">
                    <div className='d-flex align-items-center justify-content-between'>
                    <div className='page-title fw-bold mb-3 mt-0 font-20 '>{update ? 'Update' : 'Create'} Department</div>
                    {/* Dropdown */}

                    {/* <div className={classNames(styles.halfleftform, styles.form1)}>
                        <label className={styles.label} htmlFor="company">
                            Name
                        </label>
                        <input
                            disabled
                            value={currentusername} 
                            className={styles.inputform1}
                            type="text"
                            id="company"
                            name="company"
                            required
                        />
                    </div> */}</div>
                    <div className='row'>
                        <div className="col-sm-4 mb-3">
                            {/* ritik 21/4/26 start  */}
                                  {currentId && (
                                      <div>
                                          <div>
                                              <label className={styles.label}>Select Location <span style={{ color: 'red', fontWeight: "Bold" }}> *</span></label>
                                              <select style={{ padding: '5px 10px', color: 'black', backgroundColor: 'white' }}
                                                  className="form-select"
                                                  value={selectedSiteFilter}
                                                  disabled
                                              >
                                                  <option value={selectedSiteFilter}>
                                                      {siteCollections?.find((site: any) => site.siteUrl === selectedSiteFilter)?.label || selectedSiteFilter}
                                                  </option>
                                              </select>
                                          </div>
                                      </div>
                                  )}
     {/* ritik 21/4/26 end */}
                            {!currentId && (
        
//   <div>
//    <div> <label className={styles.label}>Select Location <span style={{
//                           color:'red',
//                           fontWeight:"Bold"
//                         }}> *</span></label>
//                         {/* Aman 21/4/26 start*/}
//     {/* <select style={{padding:'5px 10px'}}
//       className="form-select"
//       value={selectedSiteFilter}
//       onChange={(e) => setSelectedSiteFilter?.(e.target.value)}
//     > */}
//     <select
//   style={{
//     padding:'5px 10px',
//     border: formErrors.location ? "2px solid red" : undefined,
//     backgroundColor: formErrors.location ? "#fee6e6" : undefined
//   }}
//   className="form-select"
//   value={selectedSiteFilter}
//   onChange={(e) => {
//     setSelectedSiteFilter?.(e.target.value);
//     if (e.target.value) {
//       setFormErrors(prev => ({ ...prev, location: false }));
//     }
//   }}
//   required
// >

//     {/* Aman 21/4/26 end*/}
//       <option value="">Select Location...</option>
//       {siteCollections?.map((site: any) => (
//         <option key={site.siteUrl} value={site.siteUrl}>
//           {site.label}
//         </option>
//       ))}
//     </select>
//   </div> </div>
//Puja kumari 22/04/2026
<div>
   <div> <label className={styles.label}>Select Location<span style={{
                          color:'red',
                          fontWeight:"Bold"
                        }}> *</span></label>
                        {/* Puja - Added wrapper div for consistent border styling */}
                        <div
                          style={{
                            border: formErrors.location ? "2px solid #fe0100" : "1px solid #dee2e6",
                            borderRadius: "5px",
                            backgroundColor: formErrors.location ? "#fee6e6" : undefined,
                            padding: "2px"
                          }}
                        >
    <select
  style={{
    padding:'5px 10px',
    border: "none",
    backgroundColor: "transparent",
    width: "100%"
  }}
  className="form-select"
  value={selectedSiteFilter}
  onChange={(e) => {
    const newValue = e.target.value;
    setSelectedSiteFilter?.(newValue);
    if (newValue && newValue.trim() !== "" && newValue.toUpperCase() !== "ALL") {
      setFormErrors(prev => ({ ...prev, location: false }));
    }
  }}
  required
>
 
    {/* Aman 21/4/26 end*/}
      <option value="">Select Location...</option>
      {siteCollections?.map((site: any) => (
        <option key={site.siteUrl} value={site.siteUrl}>
          {site.label}
        </option>
      ))}
    </select>
                        </div>
  </div> </div>
)}
{/* end here  */}

                        </div>
                 <div className="col-sm-4 mb-3">
                    <label className={styles.label} htmlFor="jobTitle">
                            Title<span style={{
                          color:'red',
                          fontWeight:"Bold"
                        }}> *</span>
                        </label>
                        {/* Aman 21/4/26 start*/}
                        {/* <input
                            className={styles.inputform1}
                            id="jobTitle"
                            name="jobTitle"
                            value={jobTitle}
                            onChange={(e) => setJobTitle(e.target.value)}
                            required
                        /> */}
                        <input
  className={styles.inputform1}
  style={{
    border: formErrors.title ? "2px solid #fe0100" : undefined,
    backgroundColor: formErrors.title ? "#fee6e6" : undefined
  }}
  id="jobTitle"
  name="jobTitle"
  value={jobTitle}
  onChange={(e) => {
    setJobTitle(e.target.value);
    if (e.target.value) {
      setFormErrors(prev => ({ ...prev, title: false }));
    }
  }}
  required
/>
     {/* Aman 21/4/26 end*/}
                    </div>
                    
                    <div className="col-sm-4 mb-3">
                    <label className={styles.label} htmlFor="isActive">
                            Active<span style={{
                          color:'red',
                          fontWeight:"Bold"
                        }}> *</span>
                        </label>
                        {/* Aman 21/4/26 start */}
                        {/* <div className={styles.radioContainer}> */}
                        <div
  style={{
    border: formErrors.isActive ? "2px solid #fe0100" : undefined,
    padding: "5px",
    borderRadius: "5px",
    backgroundColor: formErrors.isActive ? "#fee6e6" : undefined
  }}
>
                      {/* Aman 21/4/26 end */}
                        <div className={styles.radioContainer}>
                            <div className={styles.radioItem}>
                            <input
                                type="radio"
                                id="yesOption"
                                name="isActive"
                                value="Yes"
                                checked={isActive === 'Yes'}
                                // Aman 21/4/26 start
                                //onChange={(e) => setIsActive(e.target.value)}
                                                  onChange={(e) => {
                                                      setIsActive(e.target.value);
                                                      setFormErrors(prev => ({ ...prev, isActive: false }));
                                                  }}
                                                  // Aman 21/4/26 end
                                required
                            />
                            <label htmlFor="yesOption" className='newf'>Yes</label>
                            </div>
                            <div className={styles.radioItem}>
                            <input
                                type="radio"
                                id="noOption"
                                name="isActive"
                                value="No"
                                checked={isActive === 'No'}
                                // Aman 21/4/26 start
                               // onChange={(e) => setIsActive(e.target.value)}
                                                  onChange={(e) => {
                                                      setIsActive(e.target.value);
                                                      setFormErrors(prev => ({ ...prev, isActive: false }));
                                                  }}
                                // Aman 21/4/26 end
                                required
                            />
                            <label htmlFor="noOption" className='newf'>No</label>
                            </div>
                        </div>
                    </div>
                    </div>
                    <div className="col-sm-4 mb-3">
                    <label className={styles.label} htmlFor="isExternal">
                            Public<span style={{
                          color:'red',
                          fontWeight:"Bold"
                        }}> *</span>
                        </label>
                        {/* Aman 21/4/26 start */}
                        {/* <div className={styles.radioContainer}> */}
                        <div
  style={{
    border: formErrors.isExternal ? "2px solid #fe0100" : undefined,
    padding: "5px",
    borderRadius: "5px",
    backgroundColor: formErrors.isExternal ? "#fee6e6" : undefined
  }}
>
                      {/* Aman 21/4/26 end */}
                        <div className={styles.radioContainer}>
                            <div className={styles.radioItem}>
                            <input
                                type="radio"
                                id="yesOption"
                                name="isExternal"
                                value="Yes"
                                checked={isExternal === 'Yes'}
                                // Aman 21/4/26 start
                                //onChange={(e) => setIsExternal(e.target.value)}
                                                  onChange={(e) => {
                                                      setIsExternal(e.target.value);
                                                      setFormErrors(prev => ({ ...prev, isExternal: false }));
                                                  }}
                                // Aman 21/4/26 end
                                required
                            />
                            <label htmlFor="yesOption" className='newf'>Yes</label>
                            </div>
                            <div className={styles.radioItem}>
                            <input
                                type="radio"
                                id="noOption"
                                name="isExternal"
                                value="No"
                                checked={isExternal === 'No'}
                                // Aman 21/4/26 start
                                // onChange={(e) => setIsExternal(e.target.value)}
                                                  onChange={(e) => {
                                                      setIsExternal(e.target.value);
                                                      setFormErrors(prev => ({ ...prev, isExternal: false }));
                                                  }}
                                // Aman 21/4/26 end
                                required
                            />
                            <label htmlFor="noOption" className='newf'>No</label>
                            </div>
                        </div>
                    </div>
                    </div>

                    <div className="col-sm-8 mb-1">
                    <label className={styles.label} htmlFor="isActive">
                            Description<span style={{
                          color:'red',
                          fontWeight:"Bold"
                        }}> *</span>
                        </label>
                        {/* Aman 21/4/26 start  */}
                        {/* <input style={{height:'80px'}}
                            className={styles.inputform1}
                            id="description"
                            name="description"
                            value={description}
                            onChange={(e) => setDescription(e.target.value)}
                            required
                        /> */}
                                  <input
                                      style={{
                                          height: '80px',
                                          border: formErrors.description ? "2px solid #fe0100" : undefined,
                                          backgroundColor: formErrors.description ? "#fee6e6" : undefined
                                      }}
                                      className={styles.inputform1}
                                      id="description"
                                      name="description"
                                      value={description}
                                      onChange={(e) => {
                                          setDescription(e.target.value);
                                          if (e.target.value) {
                                              setFormErrors(prev => ({ ...prev, description: false }));
                                          }
                                      }}
                                      required
                                  />
                                  {/* Aman 21/4/26 end  */}
                    </div>
                </div>
                </div>
                </form>
        </div>
        
        <div className="mt-2 text-center  mb-2">
                    <button type="submit" className={styles.backbuttonform1} onClick={handleSubmit}>
                        <p className={styles.Addtext}>Submit</p>
                    </button>
                    <button type="button" style={{marginBottom:'15px'}} className={styles.addbuttonargform1}
                    // Aman 21/4/26 start
                        //onClick={onCancel}
                      onClick={() => {
                          clearForm();
                          onCancel();
                      }}
                      // Aman 21/4/26 end
                    >
                        <p 
                            className={styles.Addtext}
                            
                        >
                            Cancel
                        </p>
                    </button>
        </div>
    </div>
    </>
  )
}

// const BasicForm = () => {
//     return (
//         <Provider>
//             <Basic/>
//         </Provider>
//     );
// };

// export default BasicForm;

const BasicForm: React.FC<BasicFormProps> = ({
      sp,
        context,
      selectedSiteFilter,
        setSelectedSiteFilter,
    siteCollections,
    currentId,
    currentJobTitle,
    currentIsActive,
    onCancel,
    IsExternal
})=>{ return (
        <Provider>
            <Basic
            sp={sp}
            context={context}
            selectedSiteFilter={selectedSiteFilter}
            setSelectedSiteFilter={setSelectedSiteFilter}
            siteCollections={siteCollections}
            currentId={currentId}
            currentJobTitle={currentJobTitle}
            currentIsActive={currentIsActive}
            onCancel={onCancel} 
            IsExternal={IsExternal}
            />
        </Provider>
    );
};

export default BasicForm;
