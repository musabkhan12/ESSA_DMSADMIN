import * as React from 'react';
import Provider from '../../../GlobalContext/provider';
import  styles from './BasicForm.module.scss'
import classNames from "classnames";
import { getSP } from '../loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import { useState} from "react";
import { EnvironmentType } from '@microsoft/sp-core-library';
// import { BasicFormProps }  from './IDmsMusaibProps'
import Swal from 'sweetalert2';
interface BasicFormProps {
    currentId :any,
    currentJobTitle:any,
    currentIsActive:any,
    onCancel:any,
    IsExternal:any
}

let currentusername="";

const Basic: React.FC<BasicFormProps> = ({
    currentId,
    currentJobTitle,
    currentIsActive,
    onCancel,
    IsExternal
})=>{

    const sp: SPFI = getSP();
    // console.log(sp);
    console.log("currentId",currentId);
    const [jobTitle, setJobTitle] = useState(currentJobTitle || '');
    const [isActive, setIsActive] = useState(currentIsActive || '');
    const [isExternal, setIsExternal] = useState(IsExternal || '');
    const [description,setDescription] = useState('');
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
    const handleSubmit = async (event: any) => {
        event.preventDefault(); 
        const form=document.getElementById('createMaster') as HTMLFormElement
        if (!form.checkValidity()) {
            // form.reportValidity(); // Show validation errors

            checkValidation();
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
        console.log(newItem);

        const listTitle='MasterSiteURL';

        try {

            if(currentId){
                let alreadyExist=false;
                const entity=await sp.web.lists.getByTitle(listTitle).items.getAll();
                console.log("entity",entity);
                entity.forEach((e)=>{
                    if(e.Title !== null){
                        if(e.Title.replace(/\s+/g, '').toLowerCase() === jobTitle.replace(/\s+/g, '').toLowerCase()){
                            // alert(`${jobTitle} Already exist,`);
                            // setErrorMessage(`${jobTitle} already exists.`);
                            if(e.Active === isActive){
                                // console.log("e.Active",e.Active,"isActive",isActive);
                                alreadyExistValue1(e.Active);
                                alreadyExist=true;
                                return;
                            }
                            
                        }
                }
                })

                if(!alreadyExist){
                    console.log("Edit Entity Id",currentId);
                    await sp.web.lists.getByTitle(listTitle).items.getById(currentId).update({
                    Active:isActive
                    });
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
            const entity=await sp.web.lists.getByTitle(listTitle).items.getAll();
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
                const data=await sp.web.lists.getByTitle(listTitle).items.add(newItem);
                console.log("Testing");
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
        setDescription("");
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
        Swal.fire(`${jobTitle} is creating, This will reflect shortly in DMS`,"", "success");
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
        <div className={styles.formcontainer}>            
            {/* <div className={styles.apphier}>
                <h1 className={styles.apptitle}>Create Entity</h1>
            </div> */}
       
            <form id="createMaster" onSubmit={handleSubmit}>
                <div className="p-4">
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
                    </div> */}
                    <div className='row'>
                 <div className="col-sm-4 mb-3">
                    <label className={styles.label} htmlFor="jobTitle">
                            Title<span style={{
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
                    
                    <div className="col-sm-3 mb-3">
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
                    <div className="col-sm-3 mb-3">
                    <label className={styles.label} htmlFor="isExternal">
                            Public<span style={{
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
                                name="isExternal"
                                value="Yes"
                                checked={isExternal === 'Yes'}
                                onChange={(e) => setIsExternal(e.target.value)}
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
                                onChange={(e) => setIsExternal(e.target.value)}
                                required
                            />
                            <label htmlFor="noOption" className='newf'>No</label>
                            </div>
                        </div>
                    </div>
                    </div>

                    <div className="col-sm-12 mb-1">
                    <label className={styles.label} htmlFor="isActive">
                            Description<span style={{
                          color:'red',
                          fontWeight:"Bold"
                        }}> *</span>
                        </label>
                        <input style={{height:'80px'}}
                            className={styles.inputform1}
                            id="description"
                            name="description"
                            value={description}
                            onChange={(e) => setDescription(e.target.value)}
                            required
                        />
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
                        onClick={onCancel}
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
    currentId,
    currentJobTitle,
    currentIsActive,
    onCancel,
    IsExternal
})=>{ return (
        <Provider>
            <Basic
            
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
