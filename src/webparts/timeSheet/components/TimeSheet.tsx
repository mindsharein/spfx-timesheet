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

import { EdgeChromiumHighContrastSelector } from 'office-ui-fabric-react';

import { TextField, PrimaryButton,DefaultButton } from '@fluentui/react';

import { useBoolean } from '@fluentui/react-hooks';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { initDS, getItems } from '../../../services/DataService';



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

  const [addOpen, { setTrue: openAddPanel, setFalse: closeAddPanel, toggle: toggleAddPanel }] = useBoolean(false);
  const [count, setCount] = useState(0);

  const [projectOptions, setProjectOptions] = useState([] as IDropdownOption[]); 



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
    (async () => {
      initDS(wpContext);

      let items = await getItems("TimeSheet");
      setItems(items);

      populateProjects();
    })();
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


  // Populate Project Dropdown
  const populateProjects = () => {
    (async () => {
      if(projects.length==0) {
        let items : IProject[] = await getItems("Projects");
        setProjects(items);

        if(projectOptions.length==0) {
          for(let i of items) {
            projectOptions.push({
              key: i["ID"].toString(),
              text: `${i.Title} :: ${ i.Status }`
            } as IDropdownOption);
          }
        }
      }
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
        <Separator/>
        <Stack tokens={{ childrenGap: 5 }}>
            <Stack.Item>
              <Dropdown options={ projectOptions } placeholder="Pick a Project" label='Project' />
            </Stack.Item>
            <Stack.Item>
              <TextField label='Task'></TextField>
            </Stack.Item>
            <Stack.Item>
              <TextField label='Task'></TextField>
            </Stack.Item>
        </Stack>
      </Panel>
    </Stack>
  );
} 

