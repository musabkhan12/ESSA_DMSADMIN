import * as React from 'react';
import Provider from '../../../GlobalContext/provider';
import  styles from './BasicForm.module.scss'
// import  styles from './prac.module.scss'
import classNames from "classnames";
import { SPFI } from '@pnp/sp';
import { useState} from "react";
import { BasicFormProps }  from '../components/BasicFormProps'
import Swal from 'sweetalert2';
// import "bootstrap/dist/css/bootstrap.min.css";
// import "../../../CustomCss/mainCustom.scss";
// // import "../../verticalSideBar/components/VerticalSidebar.scss"
// import VerticalSideBar from '../../verticalSideBar/components/VerticalSideBar';
// import UserContext from '../../../GlobalContext/context';

// import { useMediaQuery } from 'react-responsive';
// import context from '../../../GlobalContext/context';
// import styles from './ArgAutomation.module.scss'
// import { useState, useEffect, useRef , useMemo } from "react";
// import JoditEditor from "jodit-react";
// import Jodit from 'jodit-react';
// import "bootstrap/dist/css/bootstrap.min.css";
// import "../../../CustomCss/mainCustom.scss";
// // import "../../verticalSideBar/components/VerticalSidebar.scss"

//  function Basic({ currentDivisionId, currentJobTitle, currentIsActive, onSubmit, onCancel }) {


interface CreateDivisionProps extends BasicFormProps {
  sp: SPFI;
  onSubmit?: () => void; // Ritik 25-03-2026
}

const CreateDevision2: React.FC<CreateDivisionProps> = ({
  sp,
  currentId,
  currentJobTitle,
  currentIsActive,
  onCancel,
  onSubmit,   // Ritik 25-03-2026
}) => {
    
    // console.log(sp);
    
    const dynamicHeading = currentId ? "Edit Division" : "Create Division";
    // console.log("Heading",dynamicHeading)
    const [jobTitle, setJobTitle] = useState(currentJobTitle || '');
    const [isActive, setIsActive] = useState(currentIsActive || '');

    // State for error message
    // const [errorMessage, setErrorMessage] = useState<string | null>(null);
    // console.log("idfromparent",currentId,"title",currentJobTitle,"state",currentIsActive,"oncancle",onCancel)
    // Handle form submission  
    // const handleSubmit = async (event:any) => {
    //     event.preventDefault(); 
    //     const form=document.getElementById('createDevision') as HTMLFormElement;

    //     const newItem = {
    //         Title: jobTitle, 
    //         Active: isActive, 
    //     };
    //     console.log(newItem);

    //     const listTitle='DivisionMasterList';

    //     try {
    //         // Insert data into SharePoint list
    //         const data=await sp.web.lists.getByTitle(listTitle).items.add(newItem);
    //         console.log(data)
    //         alert('Item added successfully');

    //     } catch (error) {
    //         console.error('Error adding item:', error);
    //         alert('Error adding item');
    //     }

    //     if(form){
    //         form.submit();
    //     }
    // };
    const handleSubmit = async (event:any) => {
        event.preventDefault();
        const newItem = { Title: jobTitle, Active: isActive };
        const listTitle = 'DivisionMasterList';
        const form=document.getElementById('createDivision') as HTMLFormElement
        if (!form.checkValidity()) {
            // form.reportValidity(); // Show validation errors
            checkValidation();
            return;
        }

        // Remove alphanumeric characters and also check the limit start
        const nonAlphaNumericForEntity = jobTitle.replace(/[^a-zA-Z0-9 _-]/g, '');
        // const nonAlphaNumericDescription = description.replace(/[^a-zA-Z0-9 ]/g, '');
        if (jobTitle !== nonAlphaNumericForEntity) {
            console.log('Special characters are not allowed in entity.');
            checkLimitAndAlphanumericCharacter('Special characters are not allowed in the Division field.');
            return;
        } else if (nonAlphaNumericForEntity.length > 100) {
            console.log('Input cannot exceed 30 characters.');
            checkLimitAndAlphanumericCharacter('Input cannot exceed 100 characters in the Division field.')
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
        // End

        try {
            if (currentId) {
                // fetch the data from DMSfolderMaster
                const folderMaster=await sp.web.lists.getByTitle("DMSFolderMaster").items.select("Devision","Id").filter(`Devision eq '${currentJobTitle}'`)();
                console.log("folderMaster",folderMaster);
                  // check if devision exist
                  let alreadyExist=false;
                  const devisions=await sp.web.lists.getByTitle(listTitle).items.getAll();
                //   console.log("devision",devisions);
                  devisions.forEach((devision)=>{
                    //   console.log("deparment ",devision.Title);
                      if(devision.Title !== null){
                              if(devision.Title.replace(/\s+/g, '').toLowerCase() === jobTitle.replace(/\s+/g, '').toLowerCase()){
                                  // alert(`${jobTitle} Already exist,`);
                                //   setErrorMessage(`${jobTitle} already exists.`);
                                  if(devision.Active === isActive){
                                    alreadyExistValue1();
                                    alreadyExist=true;
                                    return;
                                  }
                                  
                              }
                      }
                  })


                  if(!alreadyExist){
                    folderMaster.forEach(async(devisionToUpdate)=>{
                       const updatedData= await sp.web.lists.getByTitle('DMSFolderMaster').items.getById(devisionToUpdate.ID).update({
                            Devision:jobTitle
                        });
                        console.log("Updated department in DMSFolderMaster",updatedData);
                    })
                }
                
                if(!alreadyExist){

                    // Update existing division
                    await sp.web.lists.getByTitle(listTitle).items.getById(currentId).update(newItem);
                    updateValue(jobTitle);
                    updateValue(jobTitle);
                    clearForm();
                    if (onSubmit) onSubmit(); // Ritik 25-03-2026
                    setTimeout(()=>{
                            onCancel();
                    },1000)
                    // alert('Division updated successfully');
                }
                

            } else {
                  // check if devision exist
                  let alreadyExist=false;
                  const devisions=await sp.web.lists.getByTitle(listTitle).items.getAll();
                //   console.log("devision",devisions);
                  devisions.forEach((devision)=>{
                    //   console.log("deparment ",devision.Title);
                      if(devision.Title !== null){
                              if(devision.Title.replace(/\s+/g, '').toLowerCase() === jobTitle.replace(/\s+/g, '').toLowerCase()){
                                  // alert(`${jobTitle} Already exist,`);
                                //   setErrorMessage(`${jobTitle} already exists.`);
                                  alreadyExistValue(jobTitle);
                                  alreadyExist=true;
                                  return;
                              }
                      }
                  })
                // Create new division
                if(!alreadyExist){
                    // setErrorMessage(null);
                    await sp.web.lists.getByTitle(listTitle).items.add(newItem);
                    // alert(`${jobTitle} added successfully`);
                    onSuccess(jobTitle);
                    clearForm();
                    if (onSubmit) onSubmit(); // Ritik 25-03-2026
                    setTimeout(()=>{
                            onCancel();
                    },1000)
                }
            }
        } catch (error) {
            console.error('Error saving division:', error);
            alert('Error saving division');
        }
        
    };

    const clearForm=()=>{
        setJobTitle("");
        setIsActive("");
        // setDescription("");
    }

    const checkValidation=()=>{
        Swal.fire("Please fill out the fields!", "All fields are required");
    }

    const alreadyExistValue=(jobTitle:any)=>{
        Swal.fire(`${jobTitle} Exist`, "Please change the devision name", "warning");
    }
    const onSuccess=(jobTitle:any)=>{
        Swal.fire(`${jobTitle} Created`,"", "success");
    }
    const updateValue=(jobTitle:any)=>{
        Swal.fire(`${jobTitle} Updated`,"", "success");
    }

    const alreadyExistValue1=()=>{
        Swal.fire(`Already Exist`, "Please change the devision name or type", "warning");
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
        <div className={styles.DmsAdminForm}>
        <div className="p-4">           
            {/* <div className={styles.apphier}>
                <h1 className={styles.apptitle}>Create Division</h1>
            </div>
            <hr className={styles.hrtag}></hr> */}
            <form id="createDivision" onSubmit={handleSubmit}>
            <div className='row'>
            <div className='col-sm-4'>
                    <label className={styles.label} htmlFor="jobTitle">
                            {dynamicHeading}<span style={{
                          color:'red',
                          fontWeight:"Bold"
                        }}> *</span>
                        </label>
                        <input
                            className={styles.inputform1}
                            id="jobTitle"
                            name="jobTitle"
                            value={jobTitle}
                            onChange={(e) => setJobTitle(e.target.value)}
                            required
                        />
                    </div>
                    <div className='col-sm-4'>
                    <label className={styles.label} htmlFor="isActive">
                            Active<span style={{
                          color:'red',
                          fontWeight:"Bold"
                        }}> *</span>
                        </label>
                        <div className={styles.radioContainer}>
                        <div className={styles.radioContainer}>
                            <div className={styles.radioItem}>
                            <input
                                type="radio"
                                id="yesOption"
                                name="isActive"
                                value="Yes"
                                checked={isActive === 'Yes'}
                                onChange={(e) => setIsActive(e.target.value)}
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
                                onChange={(e) => setIsActive(e.target.value)}
                                required
                            />
                            <label htmlFor="noOption" className='newf'>No</label>
                            </div>
                        </div>
                    </div>
                    </div>
                </div>
                </form>
        </div>
        
        <div className={styles.approvecancel}>
                    <button type="submit" className={styles.backbuttonform1} onClick={handleSubmit}>
                        <p className={styles.Addtext}>Submit</p>
                    </button>
                    <button type="button" className={styles.addbuttonargform1}   onClick={onCancel}>
                        <p className={styles.Addtext}>Cancel</p>
                    </button>
        </div>
        </div>
    </>
  )
}

const CreateDevision: React.FC<CreateDivisionProps> = (props) => {
  return (
    <Provider>
      <CreateDevision2 {...props} />
    </Provider>
  );
};

export default CreateDevision;
