import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import MoveDialog from './MoveDialog/MoveDialog';
import { sp } from '@pnp/sp';

import * as strings from 'Y365FolderMoveCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IY365FolderMoveCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'Y365FolderMoveCommandSet';

export default class Y365FolderMoveCommandSet extends BaseListViewCommandSet<IY365FolderMoveCommandSetProperties> {

  @override
  public onInit(): Promise<void> {    
    Log.info(LOG_SOURCE, 'Initialized Y365FolderMoveCommandSet');

    sp.setup({
      spfxContext: this.context
    });

    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('MOVE_FOLDER');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.title = "Move Me"
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'MOVE_FOLDER':
        console.log("Clicked");
        const dialog = new MoveDialog();
        dialog.show().then(() => {
          console.log("dialog shown");
        });
        break;
      case 'COMMAND_2':
        Dialog.alert(`${this.properties.sampleTextTwo}`);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
