import * as React from 'react';
import { useState, useEffect, useRef } from "react";

import styles from './TimeSheet.module.scss';

import { Stack, IStackStyles, IStackTokens, IStackItemStyles } from "@fluentui/react/lib/Stack";

import { Label } from '@fluentui/react/lib/Label';

import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Dropdown, IDropdown, IDropdownOption, IDropdownProps} from '@fluentui/react/lib/Dropdown';
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/DateTimePicker';
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";

import { Separator } from '@fluentui/react/lib/Separator';

import ITimeSheet from '../../../models/ITimeSheet';
import IProject from '../../../models/IProject';
import ITask from '../../../models/ITask';

import { TextField, PrimaryButton,DefaultButton } from '@fluentui/react';

import { useBoolean } from '@fluentui/react-hooks';
import { INewFormProps } from './INewFormProps';

import getSP from '../../../common/data';
import { fromPairs } from '@microsoft/sp-lodash-subset';
import { DetailsColumnBase } from 'office-ui-fabric-react';
import { IItemAddResult } from '@pnp/sp/items';


export default function NewForm(props: INewFormProps) : JSX.Element {

    const [projects, setProjects] = useState([] as IProject[]);
    const [tasks,setTasks] = useState([] as ITask[]);
    const [selProject,setSelProject] = useState(null);
    const [selTask,setSelTask] = useState(null);
    const [projectOptions, setProjectOptions] = useState([] as IDropdownOption[]);
    const [taskOptions, setTaskOptions] = useState([] as IDropdownOption[]);
    
    const [fromDate, setFromDate] = useState(new Date());
    const [toDate, setToDate] = useState(new Date());

    // Add Panel - Component Refs
    const refs = {
      titleRef: useRef(null),
      projRef: useRef(null),
      taskRef: useRef<IDropdown>(null),
      fromRef: useRef<DateTimePicker>(null),
      toRef: useRef<DateTimePicker>(null),
      notesRef: useRef(null)
    };

    const [addOpen, { setTrue: openAddPanel, setFalse: closeAddPanel, toggle: toggleAddPanel }] = useBoolean(false);


    useEffect(()=> {
        // Opt-out of pnp telemetry
        populateProjects();
    },[true]);

    // Populate Project Dropdown
    const populateProjects = () => {
        (async () => {
            if(projects.length==0) {
                const sp = getSP(props.wpContext);

                let data : IProject[] = await sp.web.lists.getByTitle("Projects").items();
                setProjects(data);

                let options : IDropdownOption[] = [];

                for(let i of data) {
                    options.push({
                    key: i["ID"].toString(),
                    text: `${i.Title} :: ${ i.Status }`
                    } as IDropdownOption);
                }
                setProjectOptions(options);
            }
        })();
    }

    const populateTasks = (projID?: number) => {
        (async () => {
            const sp = getSP(props.wpContext);

            let data  = await sp.web.lists.getByTitle("Tasks").items
                                .expand("Project")
                                .select("ID","Title","Project/Title","Project/ID")
                                .filter(`Project/ID eq ${ projID }`)
                                .getAll();

            console.log("Task Items: " + JSON.stringify(data));

            setTasks(data);

            let options : IDropdownOption[] = [];

            for(let i of data) {
                options.push({
                    key: i["ID"].toString(),
                    text: `${i.Title}`
                } as IDropdownOption);
            }

            setTaskOptions(options);
        })();    
    }

    // Create TimeSheet item from entry Form
    const getItemFromForm = () : ITimeSheet => {
      return {
        Title: refs.titleRef.current.value,
        ProjectTaskId: parseInt(refs.taskRef.current.selectedOptions[0].key as string),
        From: fromDate,
        To: toDate,
        Hours: DateDiffHrs(toDate,fromDate),
        PersonId: props.currentUser.Id,
        Notes: refs.notesRef.current.value
      } as ITimeSheet
    }

    // Adds New TimeSheet item
    const addNewItem = async (newItem: ITimeSheet) : Promise<string> => {
      try {
        let sp = getSP(props.wpContext);

        console.log("126:Item : " + JSON.stringify(newItem));

        let result: IItemAddResult = await sp.web.lists.getByTitle("TimeSheet").items.add(newItem);
        
        console.log(result);

        return "";
        
      } catch(ex) {

        console.log(JSON.stringify(ex));

        return ex.toString();
      }
    }

    // Common Panel Footer Buttons
    const onRenderPanelFooter = React.useCallback(
        () => (
        <div>
            <PrimaryButton onClick={ () => {
                // TODO: Validations

                console.log("113:SELECTED TASK : " + JSON.stringify(selTask));

                // Save
                let newItem : ITimeSheet = getItemFromForm();

                addNewItem(newItem).then(m => {
                  if(m=="") {
                    alert("Added Item Successfully!");
                    props.onItemAdded(true);

                    props.onClosed(false);
                  } else {
                    alert("Add Failed: " + m);
                  }
                });
            }} 
            styles={{ root: { marginRight: 8 } }}
            >
            Save
            </PrimaryButton>
            <DefaultButton onClick={ () => props.onClosed(false) }>Cancel</DefaultButton>
        </div>
        ),
        [closeAddPanel]
    );

    return (
    <Panel headerText="New Timesheet item" 
        isOpen={ props.isOpen } 
        type={ PanelType.medium }
        onDismiss={ closeAddPanel }
        isFooterAtBottom
        onRenderFooterContent={ onRenderPanelFooter }
        onLoad={ populateProjects }
      >
        <Separator/>
        <Label>Add a new timesheet item by filling up the form below:</Label>
        <Stack tokens={{ childrenGap: 5 }}>
            <Stack.Item>
              <TextField label="Title" rows={ 2 } componentRef={ refs.titleRef } />                   
            </Stack.Item>
            <Stack.Item>
              <Dropdown options={ projectOptions } 
                  componentRef={ refs.projRef }
                  placeholder="Pick a Project" 
                  label='Project'
                  onChange={(ev,option) => {
                    setSelProject(option);
                    populateTasks(parseInt(option.key as string)); 
                  }}/>
            </Stack.Item>
            <Stack.Item>
              <Dropdown options={ taskOptions }
                  componentRef={ refs.taskRef } 
                  onChange={(ev,option)=> {
                    setSelTask(option);
                    console.log("172:SELECTED OPTION : " + JSON.stringify(selTask));

                  }}
                  placeholder="Pick a Project Task" label='Task' />
            </Stack.Item>
            <Stack.Item>
              <DateTimePicker label="From"
                  ref={ refs.fromRef }
                  dateConvention={DateConvention.DateTime}
                  timeConvention={TimeConvention.Hours24}
                  value={fromDate}
                  onChange={ (d: Date)=> { 
                    setFromDate(d);
                  }} />
            </Stack.Item>
            <Stack.Item>
              <DateTimePicker label="To"
                  ref={ refs.toRef }
                  dateConvention={DateConvention.DateTime}
                  timeConvention={TimeConvention.Hours24}
                  value={toDate}
                  onChange={ (d: Date)=> { 
                    setToDate(d);
                  }} />
            </Stack.Item>
            <Stack.Item>
              <Label>Hours: { DateDiffHrs(toDate,fromDate) }</Label>
            </Stack.Item>
            <Stack.Item>
              <TextField label="Notes"
                  componentRef={ refs.notesRef } 
                  multiline rows={ 3 } />
            </Stack.Item>
        </Stack>
      </Panel>
    );
}

// Finds difference between two dates in hours (2 decimal places)
const DateDiffHrs = (from: Date, to: Date) : number => {

  return parseFloat(((from.getTime() - to.getTime())/3600000).toFixed(2));
}
