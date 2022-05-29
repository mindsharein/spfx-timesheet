import * as React from 'react';
import { useState, useEffect } from "react";

import styles from './TimeSheet.module.scss';
import { ITimeSheetProps } from './ITimeSheetProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { Stack, IStackStyles, IStackTokens, IStackItemStyles } from "@fluentui/react/lib/Stack";
import { DetailsList, IColumn, DetailsListLayoutMode, Selection, SelectionMode } from "@fluentui/react/lib/DetailsList";

import TimeSheetService from '../../../services/TimeSheetService';
import ITimeSheet from '../../../models/ITimeSheet';
import { EdgeChromiumHighContrastSelector } from 'office-ui-fabric-react';

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
  const [dataSvc, setDataSvc] = useState(new TimeSheetService(wpContext));
  const [items, setItems] = useState([] as ITimeSheet[]);
  
  // Load Data 
  useEffect(()=> {
    console.log("Initializing TimeSheet Data Service...");
    dataSvc.init()
      .then(()=> {
          console.log("Getting TimeSheet items...");

          let data : ITimeSheet[] = [];
          
          dataSvc.getItems(100)
            .then((data)=> {
              console.log(`Fetched ${data.length} items!`);
              console.log('Saving to state-calling setItems(data)');
              setItems(data);
            });
      })
      .catch((err)=> {
        console.log("TimeSheet.tsx: Error fetching TimeSheet data: " + err);
      });
  },[true]);

  return (
    <div>
      <h2>TimeSheets</h2>
      <DetailsList 
        items={ items } columns={ columns}
        selectionMode={ SelectionMode.multiple }
        layoutMode={ DetailsListLayoutMode.justified}
        isHeaderVisible
      >

      </DetailsList>
    </div>
  );
} 

