import * as React from 'react';
import { useState, useEffect, useCallback } from "react";

import styles from './TimeSheet.module.scss';
import { ITimeSheetProps } from './ITimeSheetProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { Stack, IStackStyles, IStackTokens, IStackItemStyles } from "@fluentui/react/lib/Stack";
import { DetailsList, IColumn, DetailsListLayoutMode, Selection, SelectionMode } from "@fluentui/react/lib/DetailsList";
import { CommandBar, ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';
import { Label } from '@fluentui/react/lib/Label';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles } from '@fluentui/react/lib/Dropdown';
import { Separator } from '@fluentui/react/lib/Separator';


import ITimeSheet from '../../../models/ITimeSheet';
import IProject from '../../../models/IProject';
import ITask from '../../../models/ITask';

import { EdgeChromiumHighContrastSelector } from 'office-ui-fabric-react';

import { TextField, PrimaryButton,DefaultButton } from '@fluentui/react';

import { useBoolean } from '@fluentui/react-hooks';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { sp } from "@pnp/sp";
import "@pnp/sp/presets/all";

import PnPTelemetry from "@pnp/telemetry-js";
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/DateTimePicker';



const columns : IColumn[] = [
  {
    key: "ID",
    name: "ID",
    fieldName: "ID",
    minWidth: 25,
    maxWidth: 50,
    isResizable: true,
  },
  {
    key: "Title",
    name: "Title",
    fieldName: "Title",
    minWidth: 100,
    maxWidth: 200,
    isResizable: true,
  },
  {
    key: "From",
    name: "From",
    fieldName: "From",
    minWidth: 75,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: "To",
    name: "To",
    fieldName: "To",
    minWidth: 75,
    maxWidth: 100,
    isResizable: true,
  },
  {
    key: "Hours",
    name: "Hours",
    fieldName: "Hours",
    minWidth: 50,
    maxWidth: 75,
    isResizable: true,
  }
];


export default function TimeSheet(props: ITimeSheetProps) : JSX.Element {
  const {
    wpContext,
    description,
    isDarkTheme,
    environmentMessage,
    hasTeamsContext,
    userDisplayName
  } = props;



  // State
  const [items, setItems] = useState([] as ITimeSheet[]);
  const [projects, setProjects] = useState([] as IProject[]);
  const [tasks,setTasks] = useState([] as ITask[]);

  const [addOpen, { setTrue: openAddPanel, setFalse: closeAddPanel, toggle: toggleAddPanel }] = useBoolean(false);
  const [count, setCount] = useState(0);

  const [projectOptions, setProjectOptions] = useState([] as IDropdownOption[]);
  const [taskOptions, setTaskOptions] = useState([] as IDropdownOption[]);
  
  const [fromDate, setFromDate] = useState(new Date());
  const [toDate, setToDate] = useState(new Date());


  // Selection Object
  const _selection = new Selection({
    onSelectionChanged: () => {
      setCount(_selection.getSelectedCount());
    }
  });

  // Toolbar buttons
  const cbItems : ICommandBarItemProps[] = [
    {
      key: 'addItem',
      text: 'New',
      iconProps: {
        iconName: 'Add'
      },
      onClick: () => {
        openAddPanel();
      }
    },
    {
      key: 'editItem',
      text: 'Edit',
      iconProps: {
        iconName: 'Edit'
      },
      disabled: count == 0 ? true : false
    },
    {
      key: 'delItem',
      text: 'Delete',
      iconProps: {
        iconName: 'Delete'
      },
      disabled: count == 0 ? true : false
    },
  ];


  // Load the Items
  useEffect(()=> {
      // Opt-out of pnp telemetry
      const telemetry = PnPTelemetry.getInstance();
      telemetry.optOut();

      sp.setup({
        spfxContext: wpContext
      });

      getTimeSheetItems();

      populateProjects();

  },[true]);

  // Common Panel Footer Buttons
  const onRenderPanelFooter = React.useCallback(
    () => (
      <div>
        <PrimaryButton onClick={ closeAddPanel } styles={{ root: { marginRight: 8 } }}>
          Save
        </PrimaryButton>
        <DefaultButton onClick={ closeAddPanel }>Cancel</DefaultButton>
      </div>
    ),
    [closeAddPanel]
  );

  const getTimeSheetItems = async () => {
    let data = await sp.web.lists.getByTitle("TimeSheet").items.get();

    setItems(data);
  }


  // Populate Project Dropdown
  const populateProjects = () => {
    (async () => {
      if(projects.length==0) {
        let data : IProject[] = await sp.web.lists.getByTitle("Projects").items.get();
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
      let data : ITask[] = await sp.web.lists.getByTitle("Tasks").items
                            .expand("Project")
                            .select("ID","Title","Project/Title","Project/ID")
                            .filter(`Project/ID eq ${ projID }`)
                            .get();
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

  return (
    <Stack tokens={{ childrenGap: 5 }}>
      <h2>TimeSheets</h2>
      <h3><Label title='User:' />{ wpContext.pageContext.user.displayName } : { count } : { addOpen }</h3>
      <CommandBar items={ cbItems } />
      <DetailsList 
        items={ items } columns={ columns}
        selectionMode={ SelectionMode.single }
        layoutMode={ DetailsListLayoutMode.justified}
        selection = { _selection }
        isHeaderVisible
      >
      </DetailsList>
      <Panel headerText="New Timesheet item" 
        isOpen={ addOpen } 
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
              <Dropdown options={ projectOptions } 
                  placeholder="Pick a Project" 
                  label='Project'
                  onChange={(ev,option) => {
                    console.log("On Change fired!");

                    populateTasks(parseInt(option.key as string)); 
                  }}/>
            </Stack.Item>
            <Stack.Item>
              <Dropdown options={ taskOptions } placeholder="Pick a Project Task" label='Task' />
            </Stack.Item>
            <Stack.Item>
              <DateTimePicker label="From"
                  dateConvention={DateConvention.DateTime}
                  timeConvention={TimeConvention.Hours24}
                  value={fromDate}
                  onChange={ ()=> { console.log('not implemented')} } />
            </Stack.Item>
            <Stack.Item>
              <DateTimePicker label="To"
                  dateConvention={DateConvention.DateTime}
                  timeConvention={TimeConvention.Hours24}
                  value={toDate}
                  onChange={ ()=> { console.log('not implemented')} } />
            </Stack.Item>
            <Stack.Item>
              <Label>Hours:</Label>
            </Stack.Item>
            <Stack.Item>
              <TextField label="Notes" multiline rows={ 3 } />
            </Stack.Item>
        </Stack>
      </Panel>
    </Stack>
  );
} 

