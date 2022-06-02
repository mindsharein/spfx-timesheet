import * as React from 'react';
import { useState, useEffect, useRef } from "react";

import styles from './TimeSheet.module.scss';
import { ITimeSheetProps } from './ITimeSheetProps';

import { Stack, IStackStyles, IStackTokens, IStackItemStyles } from "@fluentui/react/lib/Stack";
import { DetailsList, IColumn, DetailsListLayoutMode, Selection, SelectionMode } from "@fluentui/react/lib/DetailsList";
import { CommandBar, ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';
import { Label } from '@fluentui/react/lib/Label';


import ITimeSheet from '../../../models/ITimeSheet';

import getSP from '../../../common/data';

import NewForm from './NewForm';

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
  const [count, setCount] = useState(0);
  const [nfOpen,setNfOpen] = useState(false);

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
      disabled: count == 0 ? true : false
    },
  ];

 
  // Load the Items
  useEffect(()=> {
    // Load TimeSheet Items // TBD:Only for Current User
    (async ()=> {
      const sp = getSP(wpContext);
      let data = await sp.web.lists.getByTitle("TimeSheet").items();
      setItems(data);
    })();
  },[true]);


  return (
    <Stack tokens={{ childrenGap: 5 }}>
      <h2>TimeSheets</h2>
      <h3><Label title='User:' />{ wpContext.pageContext.user.displayName } : { count }</h3>
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
                            onClosed={ (flag: boolean) => { 
                                          setNfOpen(flag);
                                        } 
                                      }
                  /> 
      }

    </Stack>
  );
} 

