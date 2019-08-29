import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { RowAccessor, ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';

import MoveDialogContent from '../MoveDialogContent/MoveDialogContent';


export default class MoveDialog extends BaseDialog{
  private selectedRows: ReadonlyArray<RowAccessor>;
  private sourceListTitle: string;
  private context: ListViewCommandSetContext;

  public render(){
    ReactDOM.render(<MoveDialogContent context={ this.context } onDismiss={ this.onDismiss } selectedRows={ this.selectedRows } sourceListTitle={ this.sourceListTitle } />, this.domElement);
  }

  public init(context, selectedRows: ReadonlyArray<RowAccessor>, sourceListTitle: string){
    this.context = context;
    this.selectedRows = selectedRows;
    this.sourceListTitle = sourceListTitle;
    this.show();
  }

  protected onAfterClose(): void {
    super.onAfterClose();
    
    // Clean up the element for the next dialog
    ReactDOM.unmountComponentAtNode(this.domElement);
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: true
    };
  }

  private onDismiss = () => {
    this.close();
  }

}