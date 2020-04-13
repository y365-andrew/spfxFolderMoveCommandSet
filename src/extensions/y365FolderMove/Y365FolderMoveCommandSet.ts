import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import DeleteDialog from './DeleteDialog/DeleteDialog';
import MoveDialog from './MoveDialog/MoveDialog';
import ProgressPanelHost from './ProgressPanelHost/ProgressPanelHost';
import { setup as pnpSetup } from '@pnp/common';
import * as strings from 'Y365FolderMoveCommandSetStrings';
import { SPPermission } from '@microsoft/sp-page-context';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IY365FolderMoveCommandSetProperties {
  // This is an example; replace with your own properties
}

const LOG_SOURCE: string = 'Y365FolderMoveCommandSet';

export default class Y365FolderMoveCommandSet extends BaseListViewCommandSet<IY365FolderMoveCommandSetProperties> {
  private progressPanelHost: ProgressPanelHost;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized Y365FolderMoveCommandSet');

    pnpSetup({
      spfxContext: this.context
    });

    this.progressPanelHost = new ProgressPanelHost(this.context);

    const showProgressCommand: Command = this.tryGetCommand("SHOW_PROGRESS");
    const isAdmin = this.context.pageContext.list.permissions.hasPermission(SPPermission.manageWeb);

    if(showProgressCommand){
      showProgressCommand.visible = isAdmin;
    }

    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const moveFolderCommand: Command = this.tryGetCommand('MOVE_FOLDER');
    const deleteFolderCommand: Command = this.tryGetCommand("DELETE_FOLDER");
    const isAdmin = this.context.pageContext.list.permissions.hasPermission(SPPermission.manageWeb);

    if (moveFolderCommand) {
      const hasPermission = this.context.pageContext.list.permissions.hasPermission(SPPermission.editListItems);
      const moreThanOneSelected = event.selectedRows.length >= 1;
      // This command should be hidden unless move than one row is selected and the user has edit items permission on the list
      moveFolderCommand.title = "Shift"
      moveFolderCommand.visible = hasPermission && moreThanOneSelected;
    }
    if(deleteFolderCommand){
      const moreThanOneSelected = event.selectedRows.length >= 1;
      deleteFolderCommand.visible = isAdmin && moreThanOneSelected;
    }
  }

  @override
  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    switch (event.itemId) {
      case 'MOVE_FOLDER':
        const dialog = new MoveDialog();
        //console.log(this.context.pageContext.legacyPageContext)
        //const web = await sp.web.select('Url').get();
        //const url = window.location.href;
        const list = this.context.pageContext.list.title;
        // console.log(list);
        // Normally we'd use the below to get the list name however the context is not kept up to date when navigating across lists. https://github.com/SharePoint/sp-dev-docs/issues/1743
        // this.context.pageContext.list.title
        // Except this causes issues loading the list if it's title is different to it's path
        // const list = url.replace(web.Url, '').split('/')[1];

        dialog.init(this.context, event.selectedRows, list);
        break;
      case 'SHOW_PROGRESS':
        this.progressPanelHost.show();
        break;
      case 'DELETE_FOLDER':
        const deleteDialog = new DeleteDialog();
        deleteDialog.init(this.context, event.selectedRows);

        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
