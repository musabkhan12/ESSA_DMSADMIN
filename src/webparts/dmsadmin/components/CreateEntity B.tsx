import * as React from 'react';
import Provider from '../../../GlobalContext/provider';
import  styles from './BasicForm.module.scss'
import classNames from "classnames";
import { getSP } from '../loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import { useState} from "react";

let currentusername="";

 function CreateEntity() {

    const sp: SPFI = getSP();
    // console.log(sp);

    const [jobTitle, setJobTitle] = useState('');
    const [isActive, setIsActive] = useState('');
    const [description,setDescription] = useState('');
    

    // Handle form submission 
    const handleSubmit = async (event: any) => {
        event.preventDefault(); 

        // console.log(jobTitle,isActive,description)
        const form=document.getElementById('createMaster') as HTMLFormElement

        const newItem = {
            Title: jobTitle, 
            Active: isActive,
            Description:description 
        };
        console.log(newItem);

        const listTitle='MasterSiteURL';

        try {
            // Insert data into SharePoint list
            const data=await sp.web.lists.getByTitle(listTitle).items.add(newItem);
            console.log(data)
            // alert('Item added successfully');

        } catch (error) {
            console.error('Error adding item:', error);
            // alert('Error adding item');
        }

        if(form){
            form.submit();
        }
    };

  return (
        
    <>  
        <div className={styles.DmsAdminForm}>
        <div className={styles.formcontainer}>            
            {/* <div className={styles.apphier}>
                <h1 className={styles.apptitle}>Create Entity</h1>
            </div> */}
            <hr className={styles.hrtag}></hr>
            <form id="createMaster" onSubmit={handleSubmit}>
                <div className={classNames(styles.formgroup, styles.topformgroup)}>
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
                        />
                    </div> */}
                    <div className='row'>
                    <div className="col-sm-4">
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
                    
                    <div className="col-sm-4">
                        <label className={styles.label} htmlFor="isActive">
                            Active
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
                            />
                            <label htmlFor="yesOption">Yes</label>
                            </div>
                            <div className={styles.radioItem}>
                            <input
                                type="radio"
                                id="noOption"
                                name="isActive"
                                value="No"
                                checked={isActive === 'No'}
                                onChange={(e) => setIsActive(e.target.value)}
                            />
                            <label htmlFor="noOption">No</label>
                            </div>
                        </div>
                    </div>
                    </div>

                    <div className="col-sm-12">
                        <label className={styles.label} htmlFor="description">
                            Description
                        </label>
                        <input
                            className={styles.inputform1}
                            id="description"
                            name="description"
                            value={description}
                            onChange={(e) => setDescription(e.target.value)}
                        />
                    </div>
                    </div>
                </div>
                </form>
                <div className={styles.approvecancel}>
                    <button type="submit" className={styles.backbuttonform1} onClick={handleSubmit}>
                        <p className={styles.Addtext}>Submit</p>
                    </button>
                    <button type="button" className={styles.addbuttonargform1}>
                        <p className={styles.Addtext}>Cancel</p>
                    </button>
        </div>
        </div>
        
        
    </div>
    </>
  )
}


export default CreateEntity;
