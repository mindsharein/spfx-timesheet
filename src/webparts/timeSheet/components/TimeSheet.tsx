import * as React from 'react';
import { useState, useEffect, useRef } from "react";
import { useBoolean } from '@fluentui/react-hooks';

import styles from './TimeSheet.module.scss';
import { ITimeSheetProps } from './ITimeSheetProps';

import { Stack, IStackStyles, IStackTokens, IStackItemStyles } from "@fluentui/react/lib/Stack";
import { DetailsList, IColumn, DetailsListLayoutMode, Selection, SelectionMode } from "@fluentui/react/lib/DetailsList";
import { CommandBar, ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';
import { Label } from '@fluentui/react/lib/Label';


import ITimeSheet from '../../../models/ITimeSheet';

import { deleteTimeSheetItem, getCurrentUser, getTimeSheetItems } from '../../../common/data';

import NewForm from './NewForm';
import ConfirmDialog from './ConfirmDialog';

import Message, { IMessageProps, MessageType } from './Message';


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
    key: "ProjectTask",
    name: "Project Task",
    fieldName: "ProjectTask",
    minWidth: 100,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: ITimeSheet) => <div>{ item.ProjectTask.Title}</div>
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
    onRender: (item: ITimeSheet) => <div>{ formatDateTime(new Date(item.From.toString()))  }</div>
  },
  {
    key: "To",
    name: "To",
    fieldName: "To",
    minWidth: 75,
    maxWidth: 100,
    isResizable: true,
    onRender: (item: ITimeSheet) => <div>{ formatDateTime(new Date(item.To.toString())) }</div>
  },
  {
    key: "Hours",
    name: "Hours",
    fieldName: "Hours",
    minWidth: 50,
    maxWidth: 75,
    isResizable: true,
  },
  {
    key: "Person",
    name: "Person",
    fieldName: "Person",
    minWidth: 75,
    maxWidth: 100,
    isResizable: true,
    onRender: (item: ITimeSheet) => <div>{ item.Person.FirstName } { item.Person.LastName }</div>
  }
];

function formatDateTime(d: Date) : string {
  return `${pad0(d.getDay())}/${pad0(d.getMonth())}/${d.getFullYear()} ${pad0(d.getHours())}:${pad0(d.getMinutes())}`;
}

function pad0(n: number) : string {
  return '' + (n < 10 ? '0'+n :n);
}

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
  const [count, setCount] = useState(0);
  const [nfOpen,setNfOpen] = useState(false);
  const [currentUser,setCurrentUser] = useState(null);

  const [deletePrompt, { setTrue: showDeletePrompt, setFalse: hideDeletePrompt, toggle: togglePrompt }] = useBoolean(false);

  const [mesgText,setMesgText] = useState<string>("");
  const [mesgType, setMesgType] = useState<MessageType>(MessageType.success);
  const [showMessage,setShowMessage] = useState<boolean>(false);

  const [currentSelection,setCurrentSelection] = useState([] as ITimeSheet[]);

  // Selection Object
  const _selection = new Selection({
    onSelectionChanged: () => {
      setCurrentSelection(_selection.getItems() as ITimeSheet[]);
      setCount(_selection.getSelectedCount());
    }
  });

  const selection = useRef(_selection);

  // Toolbar buttons
  const cbItems : ICommandBarItemProps[] = [
    {
      key: 'addItem',
      text: 'New',
      iconProps: {
        iconName: 'Add'
      },
      onClick: () => {
        setNfOpen(true);
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
      onClick: () => {
        showDeletePrompt();
      },
      disabled: count == 0 ? true : false
    },
  ];

  // Loads timesheet items
  const loadItems = async () => {
    try {
      let user = await getCurrentUser(wpContext);

      setCurrentUser(user);

      let data : ITimeSheet[] = await getTimeSheetItems(wpContext);
      setItems(data);

    } catch(ex) {
      console.log("Error loading TimeSheet items : " + ex.toString());
    }
  }

  const showAlert = (text: string, type: MessageType) : void => {
    setMesgText(text);
    setMesgType(type);
    setShowMessage(true);
  }

  const clearAlert = () => {
    setMesgText("");
    setMesgType(MessageType.success);
    setShowMessage(false);
  }

  // Initial Load
  useEffect(()=> {
    // Load TimeSheet Items for Current User
    loadItems();
  },[true]);


  return (
    <>
      { showMessage && <Message text={ mesgText } type={ mesgType } reset={ clearAlert } /> }
      <Stack tokens={{ childrenGap: 5 }}>
        <h2>TimeSheets</h2>
        <h3><Label title='User:' />{ currentUser!=null ? currentUser.Title : "" } : { count }</h3>
        <CommandBar items={ cbItems } />
        <DetailsList 
          items={ items } columns={ columns}
          selectionMode={ SelectionMode.single }
          layoutMode={ DetailsListLayoutMode.justified}
          selection = { _selection }
          isHeaderVisible
        >
        </DetailsList>

        { nfOpen && <NewForm wpContext={ wpContext } 
                              isOpen={ nfOpen }
                              currentUser={ currentUser }
                              onClosed={ (flag: boolean) => { 
                                            setNfOpen(flag);
                                          } 
                                        }
                              onItemAdded={ (refresh: boolean) => {
                                // Reload items as new item available
                                if(refresh) {
                                  loadItems();
                                }
                              }}
                    /> 
        }
      </Stack>
      <div>
        <h2>Current Selection</h2>
        <div>
          Title : { currentSelection.length > 0 ? currentSelection[0].Title : "" } <br/>
          Task ID : { currentSelection.length > 0 ? currentSelection[0].ProjectTask.Id : "" } <br/>
          Start : { currentSelection.length > 0 ? currentSelection[0].From : "" } <br/>
          End : { currentSelection.length > 0 ? currentSelection[0].To : "" } <br/>
          PersonID: { currentSelection.length > 0 ? currentSelection[0].Person.Id : "" } <br/>
          Person : { currentSelection.length > 0 ? currentSelection[0].Person.Name : "" } <br/>
          Project Task : { currentSelection.length > 0 ? currentSelection[0].ProjectTask.Title : "" } <br/>
          Notes:  { currentSelection.length > 0 ? currentSelection[0].Notes : "" } <br/>
        </div>
      </div>
      <ConfirmDialog show={ deletePrompt } 
              title="Delete Item?" 
              message="Do you want to delete this TimeSheet item?" 
              onClick={(del: boolean)=> {
                hideDeletePrompt();

                (async () => {
                  if(del) {
                    let item = selection.current.getSelection()[0];
          
                    let id = item["ID"] as number;
          
                    deleteTimeSheetItem(id,wpContext)
                      .then(m=> {
                        if(m=="") {
                          showAlert("Item Deleted successfully!",MessageType.success);
                          loadItems();
                        } else {
                          showAlert("Error deleting item : " + m,MessageType.error);
                        }
                      });
                  }
                })();

          }} />
    </>
  );
} 

