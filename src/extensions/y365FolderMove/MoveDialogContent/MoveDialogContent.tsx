import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';

import { Selection, SelectionMode } from 'office-ui-fabric-react/lib/Utilities';
import { Breadcrumb, IBreadcrumbItem, IBreadCrumbData } from 'office-ui-fabric-react/lib/Breadcrumb';
import { DialogContent, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { DetailsList, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { sp } from '@pnp/sp';

import styles from  './MoveDialogContent.module.scss';

export interface ISelectedItem{
  id: string,
  title: string,
  path?: string
}

export interface IMoveDialogContentProps{
  onDismiss: () => void;
}

export interface IMoveDialogContentState{
  crumbs?: IBreadcrumbItem[];
  currentFolderContents?: any[];
  libraries?: IDropdownOption[];
  selectedFolder?: ISelectedItem;
  selectedLibrary?: ISelectedItem;
}

export default class MoveDialogContent extends React.Component<IMoveDialogContentProps, IMoveDialogContentState>{
  private currentFolderColumns: IColumn[];

  constructor(props: IMoveDialogContentProps){
    super(props);

    this.state = {};

    this.currentFolderColumns = [{
      key: 'icon',
      name: '',
      minWidth: 45,
      maxWidth: 45,
      isIconOnly: true,
      iconName: 'Folder',
      onRender: () => <Icon iconName="Folder" />
    },{
      key: 'name',
      name: 'Name',
      fieldName: 'Name',
      minWidth: 400,
      onRender: this.renderLink
    }]
  }

  public componentDidMount(){
    this.getLibraries();
  }

  public render(): JSX.Element{
    return (
      <DialogContent onDismiss={ this.props.onDismiss } className={ styles.dialogBody } showCloseButton={ true } >
        <h1>Move Item</h1>
        <p>Use the below controls to move this folder.</p>
        <h2>Destination Library</h2>
        {
          this.state.libraries && this.state.libraries.length > 0 && (
            <Dropdown options={ this.state.libraries } onChanged={ this.onLibrarySelected } />
          )
        }
        <h2>Destination Folder</h2>
        <Breadcrumb items={ this.state.crumbs } />
        {
          !this.state.selectedLibrary && (
            <span>Selected a library to continue</span>
          )
        }
        {
          this.state.selectedLibrary && this.state.currentFolderContents && (
            <DetailsList columns={ this.currentFolderColumns } items={ this.state.currentFolderContents } selectionMode={ SelectionMode.none } />
          )
        }
        {
          this.state.selectedLibrary && this.state.currentFolderContents && this.state.currentFolderContents.length <= 0 && (
            <span>This folder has no subfolders. Click move to move the selected item here, otherwise use the breadcrumbs above to navigate back.</span>
          )
        }
        <DialogFooter>
          <DefaultButton text="Cancel" onClick={ this.props.onDismiss } />
          <PrimaryButton text="Move" className={ styles.primaryButton } />
        </DialogFooter>
      </DialogContent>
    )
  }

  private renderLink = (item, i, col) => {
    return (
      <Link onClick={ this.folderClicked.bind(this, item) }>{ item[col.fieldName] }</Link>
    )
  }

  private folderClicked = (folder) => {
    const prevCrumbs = this.state.crumbs || []

    const crumbs = prevCrumbs.concat([{
      key: folder.Id,
      text: folder.Name,
      onClick: () => this.crumbClicked(folder.Id, folder.ServerRelativeUrl, prevCrumbs.length)
    }]);

    this.getCurrentFolderContents(folder.ServerRelativeUrl);

    this.setState({
      crumbs,
      selectedFolder: folder.Id
    })
  }

  private crumbClicked = (folderId, folderPath, crumbIndex) => {
    this.getCurrentFolderContents(folderPath);
    const prevCrumbs = this.state.crumbs || [];
    const crumbs = this.state.crumbs.slice(0, crumbIndex+1);

    this.setState({
      crumbs,
      selectedFolder: folderId
    });
  }

  private getLibraries = async () => {
    const librariesRes = await sp.web.lists.filter(`BaseType eq 1`).select(`Title,Id,RootFolder/ServerRelativeUrl`).expand('RootFolder').get();
    const libraries: IDropdownOption[] = librariesRes.map(v => { return { key: v.Id, text: v.Title, data: v.RootFolder ? v.RootFolder.ServerRelativeUrl : null }} );

    this.setState({
      libraries
    });
  }

  private getCurrentFolderContents = async (folderPath) => {
    const subFolders = await sp.web.getFolderByServerRelativeUrl(folderPath).folders.select(`Name,ServerRelativeUrl`).get();
    const currentFolderContents = subFolders;

    this.setState({
      currentFolderContents
    })
  }

  private onLibrarySelected = async (selectedItem: IDropdownOption) => {
    const selectedLibrary: ISelectedItem = {
      id: selectedItem.key as string,
      title: selectedItem.text,
      path: selectedItem.data
    }

    const crumbs: IBreadcrumbItem[] = [{
      key: selectedItem.key as string,
      text: selectedItem.text,
      onClick: () => this.crumbClicked(selectedItem.key, selectedItem.data, 0)
    }]

    this.setState({
      crumbs,
      selectedLibrary,
      currentFolderContents: null
    });

    if(selectedItem.data){
      this.getCurrentFolderContents(selectedItem.data);
    }
  }
}