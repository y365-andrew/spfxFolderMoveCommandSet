import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { RowAccessor, ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';

import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import ProgressPanel from '../ProgressPanel/ProgressPanel';

export default class ProgressPanelHost{
    private domElement: HTMLElement;
    private isOpen: boolean;
    private isInitialised: boolean;
    private context: ListViewCommandSetContext;

    constructor(context: ListViewCommandSetContext){
        this.domElement = document.createElement('div');
        this.domElement.id = "y365ProgressPanelHost";
        this.isOpen = false;
        this.context = context;
    }
    
    public init(){
        document.body.appendChild(this.domElement);
        this.isInitialised = true;
    }

    public show(){
        if(!this.isInitialised){
            this.init();
        }

        this.isOpen = true;
        this.render();
    }

    public close(){
        this.isOpen = false;
        this.render();
    }

    public render(): void{
        ReactDOM.render(
            <ProgressPanel isOpen={ this.isOpen } onDismissed={ () => this.close() } context={ this.context } />
        , this.domElement);
    }    

}