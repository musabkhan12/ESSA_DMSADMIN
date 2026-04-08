import * as React from 'react';
import Provider from '../../../GlobalContext/provider';
import  styles from './BasicForm.module.scss'
import classNames from "classnames";
import { getSP } from '../loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import { useState} from "react";
import {BasicFormProps} from '../components/BasicFormProps'
import Swal from 'sweetalert2';
// import swal from 'sweetalert2';

interface CreateDepartmentProps extends BasicFormProps {
  sp: SPFI;
  onSubmit?: () => void; // Ritik 25-03-2026 added onSubmit function
}
const CreateDepartment: React.FC<CreateDepartmentProps> = ({
    sp,
    currentId,
    currentJobTitle,
    currentIsActive,
    onCancel,
    onSubmit // Ritik 25-03-2026
})=>{

    // const sp: SPFI = getSP();
    // console.log(sp);

    const dynamicHeading =currentId ? "Department" : "Department";

    const [jobTitle, setJobTitle] = useState(currentJobTitle || '');
    const [isActive, setIsActive] = useState(currentIsActive || '');
    // console.log("Oncancle",onCancel);

    // State for error message
    // const [errorMessage, setErrorMessage] = useState<string | null>(null);



    const handleSubmit = async (event: any) => {
        event.preventDefault(); 

        // console.log(jobTitle,isActive)
        const form=document.getElementById('createDepartment') as HTMLFormElement
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
            checkLimitAndAlphanumericCharacter('Special characters are not allowed in the Department field.');
            return;
        } else if (nonAlphaNumericForEntity.length > 100) {
            console.log('Input cannot exceed 30 characters.');
            checkLimitAndAlphanumericCharacter('Input cannot exceed 50 characters in the Department field.')
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

        const newItem = {
            Title: jobTitle, 
            Active: isActive, 
        };
        console.log(newItem);

        const listTitle='DepartmentMasterList';

        try {

            if (currentId) {

                // fetch the data from DMSfolderMaster
                const folderMaster=await sp.web.lists.getByTitle("DMSFolderMaster").items.select("Department","Id").filter(`Department eq '${currentJobTitle}'`)();
                console.log("folderMaster",folderMaster);
                 // check if department exist
                 let alreadyExist=false;
                 const departments=await sp.web.lists.getByTitle(listTitle).items.getAll();
                 // console.log("deaprtments",departments);
                 departments.forEach((department)=>{
                     // console.log("deparment ",department.Title);
                     if(department.Title !== null){
                             if(department.Title.replace(/\s+/g, '').toLowerCase() === jobTitle.replace(/\s+/g, '').toLowerCase()){
                                 // alert(`${jobTitle} Already exist,`);
                                 // setErrorMessage(`${jobTitle} already exists.`);
                                 if(department.Active === isActive){
                                    alreadyExistValue1();
                                    alreadyExist=true;
                                    return;
                                 }
                                
                             }
                     }
                 })

                if(!alreadyExist){
                    folderMaster.forEach(async(departmentToUpdate)=>{
                       const updatedData= await sp.web.lists.getByTitle('DMSFolderMaster').items.getById(departmentToUpdate.ID).update({
                            Department:jobTitle
                        });
                        console.log("Updated department in DMSFolderMaster",updatedData);
                    })
                }
                // Update existing Department
                if(!alreadyExist){
                    await sp.web.lists.getByTitle(listTitle).items.getById(currentId).update(newItem);
                    updateValue(jobTitle);
                    clearForm();
                    if (onSubmit) onSubmit(); // Ritik 25-03-2026
                    setTimeout(()=>{
                            onCancel();
                    },1000)
                    // alert('Department updated successfully');
                }
                
                

            } else {  
                
                // check if department exist
                let alreadyExist=false;
                const departments=await sp.web.lists.getByTitle(listTitle).items.getAll();
                // console.log("deaprtments",departments);
                departments.forEach((department)=>{
                    // console.log("deparment ",department.Title);
                    if(department.Title !== null){
                            if(department.Title.replace(/\s+/g, '').toLowerCase() === jobTitle.replace(/\s+/g, '').toLowerCase()){
                                // alert(`${jobTitle} Already exist,`);
                                // setErrorMessage(`${jobTitle} already exists.`);
                                alreadyExistValue(jobTitle);
                                alreadyExist=true;
                                return;
                            }
                    }
                })

                // Create new Department
                if(!alreadyExist){
                    // setErrorMessage(null); 
                    await sp.web.lists.getByTitle(listTitle).items.add(newItem);
                    onSuccess(jobTitle);
                    // alert(`${jobTitle} added successfully`);
                    clearForm();
                    if (onSubmit) onSubmit(); // Ritik 25-03-2026
                    setTimeout(()=>{
                            onCancel();
                    },1000)
                }
            }

        } catch (error) {
            console.error('Error adding item:', error);
            alert('Error adding item');
        }
        // clearForm();
        // setTimeout(()=>{
        //     onCancel();
        // },1000)
        
        // clearForm();
       
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
        Swal.fire(`${jobTitle} Exist`, "Please change the department name", "warning");
    }

    const alreadyExistValue1=()=>{
        Swal.fire(`Already Exist`, "Please change the department name or status", "warning");
    }

    const onSuccess=(jobTitle:any)=>{
        Swal.fire(`${jobTitle} Created`,"", "success");
    }

    const updateValue=(jobTitle:any)=>{
        Swal.fire(`${jobTitle} Updated`,"", "success");
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
                <h1 className={styles.apptitle}>Create Department</h1>
            </div>
            <hr className={styles.hrtag}></hr> */}
            <form id="createDepartment" onSubmit={handleSubmit}>
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
                    <button type="button" className={styles.addbuttonargform1}     onClick={onCancel}>
                        <p className={styles.Addtext}>Cancel</p>
                    </button>
        </div>
    </div>
    </>
  )
}

const CreateDepartment2: React.FC<CreateDepartmentProps> = ({
    sp,
    currentId,
    currentJobTitle,
    currentIsActive,
    onCancel,
    onSubmit // Ritik 25-03-2026
})=> {
    return (
        <Provider>
            <CreateDepartment
                    currentId={currentId}
                    currentJobTitle={currentJobTitle}
                    currentIsActive={currentIsActive}
                    onCancel={onCancel} 
                    sp={sp}
                    onSubmit={onSubmit} // Ritik 25-03-2026
            />
        </Provider>
    );
};

export default CreateDepartment2;
