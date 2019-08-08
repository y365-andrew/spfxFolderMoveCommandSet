import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';

import MoveDialogContent from '../MoveDialogContent/MoveDialogContent';


export default class MoveDialog extends BaseDialog{
  public render(){
    ReactDOM.render(<MoveDialogContent onDismiss={ this.onDismiss } />, this.domElement);
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