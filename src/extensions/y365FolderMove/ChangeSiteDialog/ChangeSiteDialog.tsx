import * as React from 'react';
import { Dialog, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { sp, Web, IWeb } from '@pnp/sp/presets/all';
import { graph } from '@pnp/graph';
import { IDrive } from '@pnp/graph/onedrive';
import '@pnp/graph/users';
import '@pnp/graph/onedrive';

import styles from './ChangeSiteDialog.module.scss';

export interface IChangeSiteDialogProps{
    isOpen: boolean;
    onDismiss: () => void;
    onSelectWeb: (destinationType: EDestinationType, Web: IWeb | IDrive) => void;
}

export interface IChangeSiteDialogState{
    destinationType: EDestinationType;
    onedriveUpnValue: string;
    siteUrlFieldValue: string;
    inputIsValid: boolean;
    validationMessage: string;
}

export enum EDestinationType{
    "sharepoint",
    "onedrive"
}

export default class ChangeSiteDialog extends React.Component<IChangeSiteDialogProps, IChangeSiteDialogState>{
    private selectedWeb: IWeb

    constructor(props: IChangeSiteDialogProps){
        super(props);

        this.state = {
            destinationType: EDestinationType.sharepoint,
            onedriveUpnValue: "",
            siteUrlFieldValue: "",
            inputIsValid: false,
            validationMessage: ""
        }
    }

    render(){
        return (
            <Dialog isBlocking={ true } title="Select Destination Site" onDismiss={ this.props.onDismiss } isOpen={ this.props.isOpen }>
                <Dropdown options={[{ key: "onedrive", text: "OneDrive"},{key: "sharepoint", text:"SharePoint" }]} label="Destination Type" onChanged={ (val) => this.onDestinationTypeChange(val) } selectedKey={ EDestinationType[this.state.destinationType] } />
                {
                    this.state.destinationType === EDestinationType.sharepoint && (
                        <TextField label="Site URL" value={ this.state.siteUrlFieldValue } onChange={ (e, val) => this.onTextFieldChange(val) } onGetErrorMessage={ (val) => this.onTextFieldGetErrorMessage(val) } />
                    )
                }{
                    this.state.destinationType === EDestinationType.onedrive && (
                        <TextField label="User Principal Name" value={ this.state.onedriveUpnValue } onChange={ (e, val) => this.onTextFieldChange(val) } onGetErrorMessage={ (val) => this.onTextFieldGetErrorMessage(val) } />
                    )
                }
                <span>{ this.state.validationMessage }</span>
                <DialogFooter>
                    <DefaultButton onClick={ () => this.props.onDismiss() } >Cancel</DefaultButton>
                    <PrimaryButton primary={ true } iconProps={{ iconName: "CheckMark" }} onClick={() => this.onValidateClick() } disabled={ !this.state.inputIsValid }>Validate</PrimaryButton>   
                </DialogFooter>
            </Dialog>
        )
    }

    private onDestinationTypeChange(selectedType: IDropdownOption){
        const destinationType = EDestinationType[selectedType.key];

        this.setState({
            destinationType,
            siteUrlFieldValue: "",
            onedriveUpnValue: "",
            inputIsValid: false
        })
    }

    private onTextFieldChange(newValue: string){
        const state = { ...this.state };
        const changedKey = this.state.destinationType === EDestinationType.sharepoint ? "siteUrlFieldValue" : "onedriveUpnValue";
        state[changedKey] = newValue;

        this.setState(state);
    }

    private onTextFieldGetErrorMessage(value: string){
        const regex: RegExp = this.state.destinationType === EDestinationType.sharepoint ? /(https:\/\/).+(sharepoint\.com).*/g : /(.+)@(.+)\.(.+)/g;

        if(value.match(regex)){
            this.setState({
                inputIsValid: true
            });

            return "";
        }
        else{
            this.setState({
                inputIsValid: true
            });

            return this.state.destinationType === EDestinationType.sharepoint ? "Please enter a valid SharePoint Online site URL" : "Please enter a valid User Principal Name"
        }
    }

    private async onValidateClick(){
        switch(this.state.destinationType){
            case EDestinationType.sharepoint: {
                try {
                    const testWeb = Web(this.state.siteUrlFieldValue)
                    const isValid = await testWeb.get();
                    // ts-lint:disable-next-line
                    console.log(isValid);
        
                    if(isValid){
                        this.props.onSelectWeb(this.state.destinationType, testWeb)
                    }
                    else{
                        this.setState({
                            validationMessage: "Could not access site, please check the Site URL is valid and that you have access."
                        });
                    }
                }
                catch(e){
                    this.setState({
                        validationMessage: "Could not access site, please check the Site URL is valid and that you have access."
                    });
                }

                break;
            }

            case EDestinationType.onedrive: {
                try {
                    //console.log(await graph.me.get());
                    const testWeb = graph.users.getById(this.state.onedriveUpnValue).drive;
                    const isValid = await testWeb.get();
                    // ts-lint:disable-next-line
                    console.log(isValid);

        
                    if(isValid){
                        this.props.onSelectWeb(this.state.destinationType, testWeb)
                    }
                    else{
                        this.setState({
                            validationMessage: "Could not access user OneDrive, please check the user exists and that you have access."
                        });
                    }
                }
                catch(e){
                    this.setState({
                        validationMessage: "Could not access user OneDrive, please check the user exists and that you have access."
                    });
                }

                break;
            }
        }
    }

}