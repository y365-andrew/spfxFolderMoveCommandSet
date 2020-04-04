import * as React from 'react';
import * as ReactDOM from 'react-dom';
import ChangeSiteDialog from '../ChangeSiteDialog/ChangeSiteDialog';
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
  private destinationWeb: IWeb;

  constructor(props: IMoveDialogContentProps){
    super(props);

    this.destinationWeb = sp.web;

    this.state = {
      confirmIsOpen: false,
      changeSiteDialogOpen: false,
      destinationSiteName: "Current site"
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
        <small>v1.3.3.7 (dev)</small>
        <div className={ styles.dialogBody }>
          <h1>Move Item</h1>
          <p>Use the below controls to move this folder.</p>
          <div className={ styles.destinationLibraryHeaderFlexContainer }>
            <h2>Destination Library</h2>

            <div className={ styles.destinationSiteContainer }>
              <span><Icon iconName="SharepointLogo"/> { this.state.destinationSiteName } </span>
              <Link onClick={ () => this.onChangeSiteLinkClicked() } >[Change]</Link>
              <ChangeSiteDialog isOpen={ this.state.changeSiteDialogOpen } onDismiss={ () => this.onChangeSiteDialogDismissed() } onSelectWeb={ (web) => this.onDestinationWebSelected(web) } />
            </div>
          </div>

          {
            this.state.libraries && this.state.libraries.length > 0 && (
              <Dropdown options={ this.state.libraries } onChanged={ this.onLibrarySelected } />
            )
          }
          <h2>Destination Folder</h2>
          <Breadcrumb items={ this.state.crumbs } />
          <TextField value={ this.state.searchTerm } onChanged={ this.onSearchTermChange } placeholder="Search" iconProps={{iconName: "Search"}} />
          {
            !this.state.selectedLibrary && (
              <span>Select a library to continue</span>
            )
          }
          {
            this.state.selectedLibrary && this.state.currentFolderContents && (
              <div className={ styles.folderListContainer }>
                <DetailsList columns={ this.currentFolderColumns } items={ this.state.filteredFolderContents || this.state.currentFolderContents } onShouldVirtualize={ () => false } selectionMode={ SelectionMode.none } viewport={{ width: 600, height: 400}}/>
              </div>
            )
          }
          {
            this.state.selectedLibrary && this.state.currentFolderContents && this.state.currentFolderContents.length <= 0 && (
              <span>This folder has no subfolders. Click move to move the selected item here, otherwise use the breadcrumbs above to navigate back.</span>
            )
          }
        </div>
        <div>
          <span>Move <b>{ this.state.selectedRowsWithProps ? this.state.selectedRowsWithProps.map(v => v.Name).join(', ') : '' }</b> to <b>{this.state.selectedFolder ? this.state.selectedFolder.title : ( this.state.selectedLibrary ? this.state.selectedLibrary.title : '') }</b></span>
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
    console.log(newTerm);
    if([null, undefined, "", " "].indexOf(newTerm) === -1){
      const filteredFolderContents = items.filter((v) => {
        return v.Name && v.Name.toLowerCase().indexOf(newTerm.toLowerCase()) >= 0
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
      key: folder.ListItemAllFields ? folder.ListItemAllFields.Id : null,
      text: folder.Name,
      onClick: () => this.crumbClicked(folder.Id, folder.ServerRelativeUrl, prevCrumbs.length)
    }]);

    this.getCurrentFolderContents(folder.ServerRelativeUrl);

    const selectedFolder: ISelectedItem = {
      id: folder.ListItemAllFields ? folder.ListItemAllFields.Id : null,
      title: folder.Name,
      path: folder.ServerRelativeUrl
    }

    this.setState({
      crumbs,
      selectedFolder,
      searchTerm: null,
      filteredFolderContents: null
    })
  }

  private crumbClicked = (folderId, folderPath, crumbIndex) => {
    this.getCurrentFolderContents(folderPath);
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
    const librariesRes = await this.destinationWeb.lists.filter(`BaseType eq 1 and Hidden eq false and IsCatalog eq false`).select(`Title,Id,RootFolder/ServerRelativeUrl`).expand('RootFolder').get();
    const libraries: IDropdownOption[] = librariesRes.map(v => { return { key: v.Id, text: v.Title, data: v.RootFolder ? v.RootFolder.ServerRelativeUrl : null }} );

    this.setState({
      libraries
    });
  }

  private getCurrentFolderContents = async (folderPath) => {
    const subFolders = await this.destinationWeb.getFolderByServerRelativeUrl(folderPath).folders.select(`Name,ServerRelativeUrl,ListItemAllFields/ID`).expand('ListItemAllFields').filter('ListItemAllFields ne null').orderBy('Name').get();
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

  private async onDestinationWebSelected(newWeb: IWeb){
    this.destinationWeb = newWeb;
    const destinationSiteName = (await newWeb.select("Url").get()).Url

    this.setState({
      changeSiteDialogOpen: false,
      libraries: null,
      destinationSiteName
    });

    this.getLibraries();
  }

}