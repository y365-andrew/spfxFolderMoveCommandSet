import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { BaseDialog } from '@microsoft/sp-dialog';
import { DeleteDialogContent } from '../DeleteDialogContent/DeleteDialogContent';
import { RowAccessor, ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';


export default class DeleteDialog extends BaseDialog{
  private selectedRows: ReadonlyArray<RowAccessor>;
  private context: ListViewCommandSetContext;
  
  public init(context: ListViewCommandSetContext, selectedRows: ReadonlyArray<RowAccessor>){
    this.context = context;
    this.selectedRows = selectedRows;
    
    this.show();
  }

  public render(): void{
    ReactDOM.render(<DeleteDialogContent context={ this.context } onDismiss={ () => this.onDismiss() } selectedRows={ this.selectedRows } />, this.domElement);
  }

  public onDismiss(): void{
    this.close();
  }

}