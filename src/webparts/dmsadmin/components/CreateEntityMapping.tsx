import * as React from 'react';
import Provider from '../../../GlobalContext/provider';
import styles from './BasicForm.module.scss';
import classNames from 'classnames';
//import { getSP } from '../loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import { useState, useEffect,useRef } from 'react';
import Swal from 'sweetalert2';

export interface forMapping{
    sp: SPFI;
    currentId: number | null;
    currentEntity: string;
    currentDevision: string;
    currentDepartment: string;
    onCancel: () => void;
}

const CreateEntityMapping: React.FC<forMapping> = ({
    sp,
    currentId,
    currentEntity,
    currentDevision,
    currentDepartment,
    onCancel,
})=>{
    //const sp: SPFI = getSP();

    
    
    // State to manage fetched data
    const [departmentList, setDepartmentList] = useState<string[]>([]);
    const [divisionList, setDivisionList] = useState<string[]>([]);
    const [masterList, setMasterList] = useState<string[]>([]);

    // State to manage dropdown visibility
    const [showDepartmentDropdown, setShowDepartmentDropdown] = useState(false);
    const [showDivisionDropdown, setShowDivisionDropdown] = useState(false);
    const [showTitleDropdown, setShowTitleDropdown] = useState(false);

    // State to manage selected values
    const [selectedDepartment, setSelectedDepartment] = useState(currentDepartment || '');
    const [selectedDivision, setSelectedDivision] = useState(currentDevision || "");
    const [selectedTitle, setSelectedTitle] = useState(currentEntity || '');

    const departmentRef = useRef<HTMLDivElement | null>(null);
    const divisionRef = useRef<HTMLDivElement | null>(null);
    const titleRef = useRef<HTMLDivElement | null>(null);

    // disable the enity when edit
    React.useEffect(()=>{
        setDisableInput(true);
    },[]);
    const [disableInput, setDisableInput]=useState(false);
    if(currentId !== null && disableInput){
        const title=document.getElementById("Title") as HTMLInputElement;
        title.disabled = true;
    }

    // const [entityMapping,setEntityMappingDetails]=useState<any[]>([]);

    // Handle form submission
    const handleSubmit = async (event: any) => {
        event.preventDefault();
        const form = document.getElementById('createMaster') as HTMLFormElement;
        if (!form.checkValidity()) {
            // form.reportValidity(); // Show validation errors
            checkValidation();
            return;
        }

        try {

             // fetched data from entityDevisonDepartmentMapping to verify the already exist mapping
             const entityMap = await sp.web.lists
             .getByTitle("EntityDivisionDepartmentMappingMasterList")
             .items.select(
               "Entitylookup/Title",
               "Entitylookup/Active",
               "Devisionlookup/Title",
               "Departmentlookup/Title",
               "Devisionlookup/Active",
               "Departmentlookup/Active",
               "Id",
               "UniqueId",
               "Created",
               "Author/Title"
             )
             .expand("Entitylookup", "Devisionlookup", "Departmentlookup","Author")();
            //  setEntityMappingDetails(entity);

            // Entity Lookup
            const entity = await sp.web.lists
                .getByTitle('MasterSiteURL')
                .items.filter(`Title eq '${selectedTitle}'`)
                .select('ID')()
                .then((items) => items[0]);

            if (!entity) throw new Error(`Entity with title "${selectedTitle}" not found.`);

            // Division Lookup
            const division = await sp.web.lists
                .getByTitle('DivisionMasterList')
                .items.filter(`Title eq '${selectedDivision}'`)
                .select('ID')()
                .then((items) => items[0]);

            if (!division) throw new Error(`Division with title "${selectedDivision}" not found.`);

            // Department Lookup
            console.log("selectedDepartment",selectedDepartment);
            let department;
            if(selectedDepartment !== ""){
                department = await sp.web.lists
                .getByTitle('DepartmentMasterList')
                .items.filter(`Title eq '${selectedDepartment}'`)
                .select('ID')()
                .then((items) => items[0]);

                if (!department) throw new Error(`Department with title "${selectedDepartment}" not found.`);
            }
            let payload:any={}
            if(selectedDepartment === ""){
                payload.EntitylookupId=entity.ID;
                payload.DevisionlookupId=division.ID;
            }else{
                payload.EntitylookupId=entity.ID;
                payload.DevisionlookupId=division.ID;
                payload.DepartmentlookupId=department.ID;
            }

            console.log("payload",payload);

            console.log("currentId",currentId);
            if(currentId){
                console.log("Inside the update",currentId);
                let payloadForUpdate:any={}
                if(selectedDepartment === ""){
                    payloadForUpdate.DepartmentlookupId=null;
                    payloadForUpdate.DevisionlookupId=division.ID;
                }else{
                    payloadForUpdate.DevisionlookupId=division.ID;
                    payloadForUpdate.DepartmentlookupId=department.ID;
                }
                console.log("payloadForUpdate",payloadForUpdate);
                console.log("selectedDepartment",selectedDepartment);
                              // check that of mapping already exist
                              let alreadyExist=false;
                              entityMap.forEach((mappedData)=>{
                                      // Check if the entity and division match
                                      const entityMatches = mappedData.Entitylookup?.Title === selectedTitle;
                                      const divisionMatches = mappedData.Devisionlookup?.Title === selectedDivision;
                                      // Check if the department is either not selected or matches when selected
                                    //   const departmentMatches = !selectedDepartment || mappedData.Departmentlookup?.Title === selectedDepartment;
                                    let selectd;
                                    if(selectedDepartment === ""){
                                        selectd=undefined;
                                    }else{
                                        selectd=selectedDepartment;
                                    }
                                    console.log("mappedData.Departmentlookup?.Title",mappedData.Departmentlookup?.Title);
                                    const departmentMatches = mappedData.Departmentlookup?.Title === selectd;
                                    console.log("departmentMatches",departmentMatches);
                                      // If all conditions are met, the mapping already exists
                                      if (entityMatches && divisionMatches && departmentMatches) {
                                              alreadyExist = true;
                                              console.log("Already Mapped");
                                              alreadyExistValue();
                                          }
                              })

                              if(!alreadyExist){
                                await sp.web.lists.getByTitle('EntityDivisionDepartmentMappingMasterList').items.getById(currentId).update(payloadForUpdate);
                                console.log("Data Updated Successfuuly")
                                clearForm();
                                updateValue();
                                setTimeout(()=>{
                                    onCancel();
                                    },1000)
                              }
            }else{
                            // check that of mapping already exist
                    let alreadyExist=false;
                    entityMap.forEach((mappedData)=>{
                            // Check if the entity and division match
                            const entityMatches = mappedData.Entitylookup?.Title === selectedTitle;
                            const divisionMatches = mappedData.Devisionlookup?.Title === selectedDivision;

                            // Check if the department is either not selected or matches when selected
                            const departmentMatches = !selectedDepartment || mappedData.Departmentlookup?.Title === selectedDepartment;

                            // If all conditions are met, the mapping already exists
                            if (entityMatches && divisionMatches && departmentMatches) {
                                    alreadyExist = true;
                                    console.log("Already Mapped");
                                    alreadyExistValue();
                                }
                    })

            

                if(!alreadyExist  ){
                        // Adding data to the EntityDivisionDepartmentMappingMasterList
                    // const newItem = await sp.web.lists.getByTitle('EntityDivisionDepartmentMappingMasterList').items.add({
                        
                    //     EntitylookupId: entity.ID, 
                    //     // Entitylookup_x003a_Title: `${selectedTitle}`,
                        
                    //     DevisionlookupId: division.ID, 
                    //     // Devisionlookup_x003a_Devision:`${selectedDepartment}`,

                    //     DepartmentlookupId: department.ID,
                    //     // Departmentlookup_x003a_Departmen:`${selectedDepartment}`

                    // });
                    const newItem = await sp.web.lists.getByTitle('EntityDivisionDepartmentMappingMasterList').items.add(payload);
                    clearForm();

                    onSuccess();
                    setTimeout(()=>{
                        onCancel();
                        },1000)
                    }
                }
            
           
        } catch (error) {
            console.log('Error', error);
        }

        // if (form) {
        //     form.submit();
        // }
        // clearForm();
    };

    const clearForm=()=>{
        setSelectedDepartment('');  
        setSelectedDivision("");
        setSelectedTitle("");
    }

    const checkValidation=()=>{
        Swal.fire("Please fill out the fields!", "Entity and Devision are required");
    }

    const alreadyExistValue=()=>{
        Swal.fire(`Already Mapped`, "Please change the mapping", "warning");
    }

    const alreadyExistValue1=()=>{
        Swal.fire(`Already Exist`, "Please change the type", "warning");
    }

    const onSuccess=()=>{
        Swal.fire(`Mapping Done`,"", "success");
    }

    const updateValue=()=>{
        Swal.fire(`Mapping Updated`,"", "success");
    }

    // Fetch data for dropdowns
    useEffect(() => {
        const getData = async () => {
            try {
                const departmentData = await sp.web.lists.getByTitle('DepartmentMasterList').items.select('Title')();
                const divisionData = await sp.web.lists.getByTitle('DivisionMasterList').items.select('Title')();
                const masterData = await getAllMasterData();
                console.log(masterData);
                
                console.log('here is my master data ',masterData);
                setDepartmentList(departmentData.filter(item => item.Title !== null).map(item => item.Title));
                setDivisionList(divisionData.filter(item => item.Title !== null).map(item => item.Title));
                setMasterList(masterData.filter((item:any) => item.Title !== null).map((item:any) => item.Title));
            } catch (error) {
                console.error('Error fetching data:', error);
            }
        };

        const getAllMasterData = async () => {
            let items:any = [];
            let pagedItems = await sp.web.lists.getByTitle('MasterSiteURL').items.select('Title').getPaged();
        
            items = items.concat(pagedItems.results);
        
            while (pagedItems.hasNext) {
                pagedItems = await pagedItems.getNext();
                items = items.concat(pagedItems.results);
            }
        
            return items;
        };

        getData();
    }, []);

    // Function to handle selecting a dropdown option
    const handleSelect = (
        setSelectedValue: React.Dispatch<React.SetStateAction<string>>,
        value: string,
        setShowDropdown: React.Dispatch<React.SetStateAction<boolean>>,
        setShowOtherDropdowns: () => void
    ) => {
        setSelectedValue(value);
        setShowDropdown(false);
        setShowOtherDropdowns();
    };

    // used to hide the dropdown when click anywhere in the body
    const handleClickOutside = (event: MouseEvent) => {
        if (titleRef.current && !titleRef.current.contains(event.target as Node)) {
            setShowTitleDropdown(false);
        }
        if (divisionRef.current && !divisionRef.current.contains(event.target as Node)) {
            setShowDivisionDropdown(false);
        }
        if (departmentRef.current && !departmentRef.current.contains(event.target as Node)) {
            setShowDepartmentDropdown(false);
        }
    };

    useEffect(() => {
        document.addEventListener('mousedown', handleClickOutside);
        return () => {
            document.removeEventListener('mousedown', handleClickOutside);
        };
    }, []);


    return (
        <>
            <div className={styles.DmsAdminForm}>
                <div className={styles.formcontainer}>
                    {/* <div className={styles.apphier}>
                        <h1 className={styles.apptitle}>Entity Map Division & Department</h1>
                    </div>
                    <hr className={styles.hrtag}></hr> */}
                    <form id="createMaster" onSubmit={handleSubmit}>
                    <div className="p-4">
                            {/* Title Input */}
                            <div className='row'>
                            <div className="col-sm-4 mb-3">
                                <label className={styles.label} htmlFor="Title">
                                    Title
                                </label>
                                <input
                                    className={styles.inputform1}
                                    id="Title"
                                    name="Title"
                                    value={selectedTitle}
                                    onChange={(e) => setSelectedTitle(e.target.value)}
                                    onFocus={() => {
                                        setShowTitleDropdown(true);
                                        setShowDivisionDropdown(false);
                                        setShowDepartmentDropdown(false);
                                    }}
                                    required
                                />
                                {showTitleDropdown && (
                                    <div className={styles.dropdown}>
                                        {masterList
                                            .filter((item) =>
                                                item && selectedTitle ? 
        item.toLowerCase().includes(selectedTitle.toLowerCase()) : false
                                            )
                                            .map((item, index) => (
                                                <div
                                                    key={index}
                                                    className={styles.dropdownItem}
                                                    onClick={() => handleSelect(setSelectedTitle, item, setShowTitleDropdown, () => {
                                                        setShowTitleDropdown(false);
                                                        setShowDivisionDropdown(false);
                                                        setShowDepartmentDropdown(false);
                                                    })}
                                                >
                                                    {item}
                                                </div>
                                            ))}
                                    </div>
                                )}
                            </div>
                            <div className="col-sm-4 mb-3">
                            <label className={styles.label} htmlFor="Division">
                                    Division
                                </label>
                                <input
                                    className={styles.inputform1}
                                    id="Division"
                                    name="Division"
                                    value={selectedDivision}
                                    onChange={(e) => setSelectedDivision(e.target.value)}
                                    required
                                    onFocus={() => {
                                        setShowTitleDropdown(false);
                                        setShowDivisionDropdown(true);
                                        setShowDepartmentDropdown(false);
                                    }}
                                    disabled={!selectedTitle} // Disable if no Title is selected
                                />
                                {showDivisionDropdown && selectedTitle && (
                                    <div className={styles.dropdown}>
                                        {divisionList
                                            .filter((item) =>
                                                item && selectedDivision ? 
                                            item.toLowerCase().includes(selectedDivision.toLowerCase()) : false
                                            )
                                            .map((item, index) => (
                                                <div
                                                    key={index}
                                                    className={styles.dropdownItem}
                                                    onClick={() => handleSelect(setSelectedDivision, item, setShowDivisionDropdown, () => {
                                                        setShowTitleDropdown(false);
                                                        setShowDivisionDropdown(false);
                                                        setShowDepartmentDropdown(false);
                                                    })}
                                                >
                                                    {item}
                                                </div>
                                            ))}
                                    </div>
                                )}


                                </div>

                                <div className='col-sm-4 mb-3'>
                                <label className={styles.label} htmlFor="Department">
                                    Department
                                </label>
                                <input
                                    className={styles.inputform1}
                                    id="Department"
                                    name="Department"
                                    value={selectedDepartment}
                                    onChange={(e) => setSelectedDepartment(e.target.value)}
                                    onFocus={() => {
                                        setShowTitleDropdown(false);
                                        setShowDivisionDropdown(false);
                                        setShowDepartmentDropdown(true);
                                    }}
                                    disabled={!selectedDivision || !selectedTitle} // Disable if no Division and Title is selected is selected
                                />
                                {showDepartmentDropdown && selectedDivision && (
                                    <div className={styles.dropdown}>
                                        {departmentList
                                            .filter((item) =>
                                                item && selectedDepartment ? 
                                            item.toLowerCase().includes(selectedDepartment.toLowerCase()) : false
                                            )
                                            .map((item, index) => (
                                                <div
                                                    key={index}
                                                    className={styles.dropdownItem}
                                                    onClick={() => handleSelect(setSelectedDepartment, item, setShowDepartmentDropdown, () => {
                                                        setShowTitleDropdown(false);
                                                        setShowDivisionDropdown(false);
                                                        setShowDepartmentDropdown(false);
                                                     })}
                                                >
                                                    {item}
                                                </div>
                                            ))}
                                    </div>
                                )}

                                </div>



                            </div>

                            {/* Division Input (Disabled until Title is selected) */}
                            

                            {/* Department Input (Disabled until Division is selected) */}
                         
                        </div>
                    </form>
                </div>
                <div className={styles.approvecancel}>
                    <button
                        type="submit"
                        className={styles.backbuttonform1}
                        onClick={handleSubmit}
                        // disabled={!selectedTitle || !selectedDivision || !selectedDepartment} // Disable submit if any field is missing
                    >
                        <p className={styles.Addtext}>Create</p>
                    </button>
                    <button type="button" className={styles.addbuttonargform1} onClick={onCancel}>
                        <p className={styles.Addtext}>Cancel</p>
                    </button>
                </div>
            </div>
        </>
    );
}


export default CreateEntityMapping;
