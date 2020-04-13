import React, { useState, useEffect } from 'react';
import { generateTree } from '../lib/generateTree';

import { DialogContent } from 'office-ui-fabric-react/lib/Dialog';
import { RowAccessor, ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';

export interface IDeleteDialogContentProps{
  context: ListViewCommandSetContext;
  selectedRows: ReadonlyArray<RowAccessor>;
  onDismiss: () => void;
}

export function DeleteDialogContent(props: IDeleteDialogContentProps){

  useEffect(() => {
    generateTree(props.context, props.selectedRows);
  }, []);

  return(
    <DialogContent onDismiss={ props.onDismiss } showCloseButton={ true }>
      
    </DialogContent>
  )
}