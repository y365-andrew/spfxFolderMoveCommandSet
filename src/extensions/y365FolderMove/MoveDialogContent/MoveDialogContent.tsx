import * as React from 'react';
import * as ReactDOM from 'react-dom';
import ChangeSiteDialog, { EDestinationType } from '../ChangeSiteDialog/ChangeSiteDialog';
import { BaseDialog, IDialogConfiguration, Dialog } from '@microsoft/sp-dialog';
import { RowAccessor, ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';

import { Selection, SelectionMode } from 'office-ui-fabric-react/lib/Utilities';
import { Breadcrumb, IBreadcrumbItem, IBreadCrumbData } from 'office-ui-fabric-react/lib/Breadcrumb';
import { DialogContent, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { DetailsList, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { sp, IWeb, Files, Item } from '@pnp/sp/presets/all';

import ConfirmDialog from '../ConfirmDialog/ConfirmDialog';

import styles from  './MoveDialogContent.module.scss';
import { IDrive } from '@pnp/graph/onedrive';

export interface ISelectedRowProps{
  Id: string | number;
  Name: string;
  Type: number;
  sp?: any;
}

export interface ISelectedItem{
  id: string | number;
  title: string;
  path?: string;
}

export interface IMoveDialogContentProps{
  onDismiss: () => void;
  selectedRows: ReadonlyArray<RowAccessor>;
  sourceListTitle: string;
  context: ListViewCommandSetContext;
}

export interface IMoveDialogContentState{
  crumbs?: IBreadcrumbItem[];
  currentFolderContents?: any[];
  changeSiteDialogOpen: boolean;
  destinationSiteName?: string;
  destinationType: EDestinationType;
  filteredFolderContents?: any[];
  searchTerm?: string;
  libraries?: IDropdownOption[];
  selectedFolder?: ISelectedItem;
  selectedLibrary?: ISelectedItem;
  selectedRowsWithProps?: ISelectedRowProps[];
  confirmIsOpen: boolean;
}

export default class MoveDialogContent extends React.Component<IMoveDialogContentProps, IMoveDialogContentState>{
  private currentFolderColumns: IColumn[];
  private destinationWeb: IWeb | IDrive;

  constructor(props: IMoveDialogContentProps){
    super(props);

    this.destinationWeb = sp.web;

    this.state = {
      confirmIsOpen: false,
      changeSiteDialogOpen: false,
      destinationSiteName: "Current site",
      destinationType: EDestinationType.sharepoint
    };

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
    this.getRowProps();
  }

  public render(): JSX.Element{
    return (
      <DialogContent onDismiss={ this.props.onDismiss } showCloseButton={ true } >
        <small>v1.5.0.2</small>
        <div className={ styles.dialogBody }>
          <h1>Move Item</h1>
          <p>Use the below controls to move this folder.</p>
          <div className={ styles.destinationLibraryHeaderFlexContainer }>
            <h2>{ this.state.destinationType === EDestinationType.sharepoint ? "Destination Library" : "OneDrive" }</h2>

            <div className={ styles.destinationSiteContainer }>
              <span><Icon iconName={ this.state.destinationType === EDestinationType.sharepoint ? "SharepointLogo" : "OnedriveLogo" }/> { this.state.destinationSiteName } </span>
              <Link onClick={ () => this.onChangeSiteLinkClicked() } >[Change]</Link>
              <ChangeSiteDialog isOpen={ this.state.changeSiteDialogOpen } onDismiss={ () => this.onChangeSiteDialogDismissed() } onSelectWeb={ (type, web) => this.onDestinationWebSelected(type, web) } />
            </div>
          </div>

          {
            this.state.libraries && this.state.libraries.length > 0 && (
              <Dropdown options={ this.state.libraries } defaultSelectedKey={ this.state.libraries[0] ? this.state.libraries[0].key : 0 } onChanged={ this.onLibrarySelected } />
            )
          }
          <h2>Destination Folder</h2>
          <Breadcrumb items={ this.state.crumbs } />
          <TextField value={ this.state.searchTerm } onChange={ (_, newVal) => this.onSearchTermChange(newVal) } placeholder="Search" iconProps={{iconName: "Search"}} />
          {
            (!this.state.selectedLibrary && this.state.destinationType === EDestinationType.sharepoint) && (
              <span>Select a library to continue</span>
            )
          }
          {
            //(this.state.selectedLibrary && this.state.currentFolderContents) || (this.state.destinationType === EDestinationType.onedrive && this.state.currentFolderContents) && (
              <div className={ styles.folderListContainer }>
                <DetailsList columns={ this.currentFolderColumns } items={ (this.state.filteredFolderContents || this.state.currentFolderContents) || [] } onShouldVirtualize={ () => false } selectionMode={ SelectionMode.none } viewport={{ width: 600, height: 400}}/>
              </div>
            //)
          }
          <span>Move <b>{ this.state.selectedRowsWithProps ? this.state.selectedRowsWithProps.map(v => v.Name).join(', ') : '' }</b> to <b>{this.state.selectedFolder ? this.state.selectedFolder.title : ( this.state.selectedLibrary ? this.state.selectedLibrary.title : 'OneDrive') }</b></span>
        </div>

        <DialogFooter>
          <DefaultButton text="Cancel" onClick={ this.props.onDismiss } />
          <PrimaryButton text="Move" className={ styles.primaryButton } onClick={ this.showConfirmation } />
        </DialogFooter>
        <ConfirmDialog isOpen={ this.state.confirmIsOpen } onDismiss={ () => this.setState({ confirmIsOpen: false }) } onDismissAll={ this.props.onDismiss } selectedRows={ this.state.selectedRowsWithProps } destination={ this.state.selectedFolder ? this.state.selectedFolder : this.state.selectedLibrary } sourceListTitle={ this.props.sourceListTitle } destinationWeb={ this.destinationWeb } context= { this.props.context }/>
      </DialogContent>
    )
  }

  private getRowProps = async () => {
    const { selectedRows } = this.props;

    const selectedRowsWithPropsPromise: Promise<ISelectedRowProps>[] = selectedRows.map(async (row) => {
      const Id = row.getValueByName('ID');
      console.log(Id)
      const res: any = await sp.web.lists.getByTitle(this.props.sourceListTitle).items.getById(Id).select('Title, FieldValuesAsText/FileLeafRef,FileSystemObjectType').expand('FieldValuesAsText').get();

      const Name = res.FieldValuesAsText ? res.FieldValuesAsText.FileLeafRef : res.Title;

      return {
        Id,
        Name,
        Type: res.FileSystemObjectType,
        sp: res
      }
    });

    const selectedRowsWithProps: ISelectedRowProps[] = await Promise.all(selectedRowsWithPropsPromise);
    console.log(selectedRowsWithProps);
    this.setState({
      selectedRowsWithProps
    });
  }

  private onSearchTermChange = (newTerm) => {
    const items = this.state.currentFolderContents;
    
    if([null, undefined, "", " "].indexOf(newTerm) === -1){
      const filteredFolderContents = items.filter((v) => {
        const name = v.Name || v.name;
        return name && name.toLowerCase().indexOf(newTerm.toLowerCase()) >= 0
      });

      this.setState({
        filteredFolderContents,
        searchTerm: newTerm
      });
    }
    else{
      this.setState({
        filteredFolderContents: null,
        searchTerm: null
      })
    }
  }

  private renderLink = (item, i, col) => {
    return (
      <Link onClick={ this.folderClicked.bind(this, item) }>{ item[col.fieldName] }</Link>
    )
  }

  private folderClicked = (folder) => {
    const prevCrumbs = this.state.crumbs || []

    const crumbs = prevCrumbs.concat([{
      key: (folder.ListItemAllFields ? folder.ListItemAllFields.Id : null) ? folder.id : null,
      text: folder.Name || folder.name,
      onClick: () => this.crumbClicked(folder.Id || folder.id, folder.ServerRelativeUrl || null, prevCrumbs.length)
    }]);

    if(this.state.destinationType === EDestinationType.sharepoint){
      this.getCurrentFolderContents(folder.ServerRelativeUrl);
    }
    if(this.state.destinationType === EDestinationType.onedrive){
      this.getCurrentDriveFolderContents(folder.id);
    }

    const selectedFolder: ISelectedItem = {
      id: (folder.ListItemAllFields ? folder.ListItemAllFields.Id : null) ? folder.id : null,
      title: folder.Name || folder.name,
      path: folder.ServerRelativeUrl || folder.webUrl
    };

    this.setState({
      crumbs,
      selectedFolder,
      searchTerm: null,
      filteredFolderContents: null
    });
  }

  private crumbClicked = (folderId, folderPath, crumbIndex) => {
    if(this.state.destinationType === EDestinationType.sharepoint){
      this.getCurrentFolderContents(folderPath);
    }
    if(this.state.destinationType === EDestinationType.onedrive){
      this.getCurrentDriveFolderContents(folderId);
    }
    const prevCrumbs = this.state.crumbs || [];
    const crumbs = prevCrumbs.slice(0, crumbIndex+1);

    this.setState({
      crumbs,
      selectedFolder: folderId,
      searchTerm: null,
      filteredFolderContents: null
    });
  }

  private getLibraries = async () => {
    if(this.state.destinationType === EDestinationType.sharepoint){
      const destWeb = this.destinationWeb as IWeb;
      const librariesRes = await destWeb.lists.filter(`BaseType eq 1 and Hidden eq false and IsCatalog eq false`).select(`Title,Id,RootFolder/ServerRelativeUrl`).expand('RootFolder').get();
      const libraries: IDropdownOption[] = librariesRes.map(v => { return { key: v.Id, text: v.Title, data: v.RootFolder ? v.RootFolder.ServerRelativeUrl : null }} );
      this.setState({
        libraries
      });

      this.getCurrentFolderContents(libraries[0].data);
    }
  }

  private getCurrentFolderContents = async (folderPath) => {
    const destWeb = this.destinationWeb as IWeb;
    const subFolders = await destWeb.getFolderByServerRelativeUrl(folderPath).folders.select(`Name,ServerRelativeUrl,ListItemAllFields/ID`).expand('ListItemAllFields').filter('ListItemAllFields ne null').orderBy('Name').get();
    const currentFolderContents = subFolders;

    this.setState({
      currentFolderContents
    });
  }

  private getCurrentDriveFolderContents = async (folderId ?: string) => {
    const destDrive = this.destinationWeb as IDrive;

    if(folderId){
      const currentFolderContentsRes = await destDrive.getItemById(folderId).children.select("id","name","folder","webUrl").get();
      const currentFolderContents = currentFolderContentsRes.filter((v) => v.folder);

      this.setState({
        currentFolderContents
      });
    }
    else{
      const currentFolderContentsRes = await destDrive.root.children.select("id","name","folder","webUrl").get();
      const currentFolderContents = currentFolderContentsRes.filter((v) => v.folder);

      this.setState({
        currentFolderContents
      });
    }

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
      onClick: () => this.crumbClicked(null, selectedItem.data, 0)
    }]

    this.setState({
      crumbs,
      selectedLibrary,
      currentFolderContents: null,
      searchTerm: null,
      filteredFolderContents: null
    });

    if(selectedItem.data){
      this.getCurrentFolderContents(selectedItem.data);
    }
  }

  private showConfirmation = () => {
    this.setState({
      confirmIsOpen: true
    });
  }

  private onChangeSiteLinkClicked(){
    this.setState({
      changeSiteDialogOpen: true
    });
  }

  private onChangeSiteDialogDismissed(){
    this.setState({
      changeSiteDialogOpen: false
    });
  }

  private async onDestinationWebSelected(destinationType: EDestinationType, newWeb: IWeb | IDrive){

    this.destinationWeb = newWeb;

    if(destinationType === EDestinationType.sharepoint){
      const destWeb: IWeb = newWeb as IWeb;
      const destinationSiteName = (await destWeb.select("Url").get()).Url;

      this.setState({
        changeSiteDialogOpen: false,
        libraries: null,
        currentFolderContents: [],
        filteredFolderContents: [],
        destinationSiteName,
        destinationType: destinationType
      });

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
      }];

      this.getLibraries();
    }
    
    if(destinationType === EDestinationType.onedrive){
      const destDrive: IDrive = newWeb as IDrive;
      const onedrive = await destDrive.select("id","owner").get();
      const destinationSiteName = onedrive.owner.user.displayName;

      const crumbs = [{
        key: onedrive.id,
        text: "OneDrive",
        onClick: () => this.crumbClicked(null, null, 0)
      }];

      this.setState({
        changeSiteDialogOpen: false,
        crumbs,
        libraries: null,
        currentFolderContents: null,
        filteredFolderContents: null,
        destinationSiteName,
        destinationType: destinationType
      });

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
        fieldName: 'name',
        minWidth: 400,
        onRender: this.renderLink
      }];

      this.getCurrentDriveFolderContents();
    }   
  }

}