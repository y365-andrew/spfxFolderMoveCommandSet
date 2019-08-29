import * as React from 'react';
import { from, Observable } from 'rxjs';
import { Dialog, DialogType, DialogFooter, DialogContent } from 'office-ui-fabric-react/lib/Dialog';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { DetailsList, IColumn, SelectionMode,DetailsRow } from 'office-ui-fabric-react/lib/DetailsList';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { sp, Folder, SPBatch } from '@pnp/sp';

import { ISelectedItem, ISelectedRowProps } from '../MoveDialogContent/MoveDialogContent';
import { moveFile, moveFolder } from '../moveFunctions/move.function';
import styles from './ConfirmDialog.module.scss';
import { Label } from 'office-ui-fabric-react/lib/Label';

export interface IConfirmSelectedRowProps extends ISelectedRowProps{
  Exists: boolean;
  Rename: boolean;
  NewName?: string;
}

export interface IConfirmDialogProps{
  isOpen: boolean;
  onDismiss: () => void;
  onDismissAll: () => void;
  selectedRows: ISelectedRowProps[];
  destination: ISelectedItem;
  sourceListTitle: string;
}

export interface IConfirmDialogState{
  selectedRowsWithProps?: IConfirmSelectedRowProps[];
  // Is errored dictates if the move button should be disabled
  isErrored: boolean;
  isWorking: boolean;
  // Has errored indicates there was an error moving the files
  hasErrored: boolean;
  hasCompleted: boolean;
  log: string[];
}

export default class ConfirmDialog extends React.Component<IConfirmDialogProps, IConfirmDialogState>{
  private columns: IColumn[];
  private log$: Observable<string>;

  constructor(props: IConfirmDialogProps){
    super(props);

    this.state = {
      isErrored: false,
      isWorking: false,
      hasCompleted: false,
      hasErrored: false,
      log: []
    };

    this.columns = [{
      key: 'name',
      fieldName: 'Name',
      minWidth: 200,
      name: 'Name',
      onRender: this.renderNameColumn
    },{
      key: 'rename',
      minWidth: 50,
      name: 'Rename',
      onRender: this.renderToggleColumn
    }];
  }

  public componentWillReceiveProps(nextProps: IConfirmDialogProps){
    if(nextProps.isOpen === true && this.props.isOpen === false){
      this.folderExistsInDestination();

      this.setState({
        isErrored: false,
        isWorking: false,
        hasCompleted: false,
        hasErrored: false,
        log: []
      });
    }
  }

  public render(){
    return (
      <Dialog maxWidth="700" isOpen={ this.props.isOpen } title="Are you sure?" type={ DialogType.largeHeader } onDismiss={ this.props.onDismiss }>
          <p>The following items will be moved to { this.props.destination && this.props.destination.path }:</p>
          {
            this.state.selectedRowsWithProps && this.state.selectedRowsWithProps.length > 0 && (
              <DetailsList columns={ this.columns } items={ this.state.selectedRowsWithProps } selectionMode={ SelectionMode.none } />
            )
          }
        <DialogFooter>
          <DefaultButton text="Cancel" onClick={ this.props.onDismiss } />
          <PrimaryButton text="Move" disabled={ this.state.isErrored } onClick={ this.move } />
        </DialogFooter>

        <Dialog isOpen={ this.state.isWorking || this.state.hasCompleted || this.state.hasErrored } isBlocking={ true } >
          {
            this.state.isWorking && (
              <Spinner label={ this.state.log[this.state.log.length -1] } />
            )
          }{
            this.state.hasCompleted && (
              <div className={ styles.progressDialog }>
                <Icon iconName="Completed" className={ styles.successIcon } />
                <Label>Items moved successfully!</Label>
                <DialogFooter>
                  <DefaultButton text="Close" onClick={ this.props.onDismissAll } />
                </DialogFooter>
              </div>
            )
          }{
            this.state.hasErrored && (
              <div className={ styles.progressDialog }>
                <Icon iconName="ErrorBadge" className={ styles.errorIcon } />
                <Label>An error ocurred, please refresh the page try again or contact IT.</Label>
                <DialogFooter>
                  <DefaultButton text="Close" onClick={ this.props.onDismissAll } />
                </DialogFooter>
              </div>
            )
          }
        </Dialog>
      </Dialog>
    );
  }

  private renderToggleColumn = (item, i, col) => {
    return(
      <div>
        <Toggle onChange={ (ev, val) => { this.onToggleRename(val, i); }} checked={ item.Rename || false } />
      </div>
    );
  }

  private renderNameColumn = (item, i, col) => {
    return(
      <div className={ styles.nameColumnContainer }>
        { item[col.fieldName] }
        {
          item.Rename && (
            <TextField prefix="New name" value={ item.NewName || "" } onGetErrorMessage={ this.validateRenameText } onChange={ (ev, newVal) => { this.onRenameTextChange(i, newVal); } } />
          )
        }
        {
          item.Exists && !item.Rename && (
            <p className={ styles.warningText }><Icon iconName="Warning" />&nbsp;A folder with the same name exists in the destination, if you do not rename it the contents will be merged.</p>
          )
        }
      </div>
    );
  }

  private validateRenameText = (value) => {
    if(value === "" || value === null || value === undefined){
      this.setState({
        isErrored: true
      });

      return "This field cannot be blank";
    }
    else if(value.length >= 255){
      this.setState({
        isErrored: true
      });

      return "The specified folder name is too long.";
    }

    this.setState({
      isErrored: false
    });

    return "";
  }

  private onRenameTextChange = (itemIndex: number, newVal: string) => {
    const selectedRowsWithProps = [...this.state.selectedRowsWithProps];
    selectedRowsWithProps[itemIndex] = {
      ...selectedRowsWithProps[itemIndex],
      NewName: newVal
    };

    this.setState({
      selectedRowsWithProps
    });
  }

  private onToggleRename = (newValue, itemIndex) => {
    const selectedRowsWithProps = [...this.state.selectedRowsWithProps];
    selectedRowsWithProps[itemIndex] = {
      ...selectedRowsWithProps[itemIndex],
      Rename: newValue
    };

    this.setState({
      selectedRowsWithProps,
      isErrored: false
    });
  }

  private folderExistsInDestination = async () => {
    const itemsToMove = this.props.selectedRows;
    const destination = this.props.destination;

    const selectedRowsWithPropsPromise = itemsToMove.map(async (item) => {

      try{
        const exists = item.Type === 1 ? await sp.web.getFolderByServerRelativeUrl(`${destination.path}/${item.Name}`).get() : await sp.web.getFileByServerRelativeUrl(`${destination.path}/${item.Name}`).get();
        
        return {
          ...item,
          Exists: exists.Exists,
          Rename: true
        }
      }
      catch(e){
        return {
          ...item,
          Exists: false,
          Rename: false
        }
      }

    });

    const selectedRowsWithProps = await Promise.all(selectedRowsWithPropsPromise);

    this.setState({
      selectedRowsWithProps
    });
  }

  private subscribeLog = () => {
    this.log$.subscribe({
      next: (nextVal) => {
        console.log(nextVal);

        const log = [...this.state.log, nextVal];
        this.setState({
          log
        });
      },
      complete: () => {
        const log = [...this.state.log, "Items move successfully!!"];
        this.setState({
          hasCompleted: true,
          isWorking: false,
          log
        });
      },
      error: (err) => {
        console.log(err);
        const log = [...this.state.log, "Error: an error occurred with the observable log, please contact IT."];

        this.setState({
          hasErrored: true,
          isWorking: false,
          log
        });
      }
    });
  }

  private move = async () => {
    this.setState({
      isWorking: true
    });

    this.log$ = Observable.create((observer) => {
      observer.next("Initialising item shift.");

      const promises = this.state.selectedRowsWithProps.map(async (row) => {
        // FOLDERS
        if(row.Type === 1){
          const folderName = row.Rename === true ? row.NewName : row.Name;
          const res = await sp.web.lists.getByTitle(this.props.sourceListTitle).items.getById(row.Id as number).folder.get();
          return moveFolder(res.ServerRelativeUrl, `${this.props.destination.path}/${folderName}`, observer)
        }
        //FILES
        else if(row.Type === 0){
          const fileName = row.Rename === true ? row.NewName : row.Name;
          const res = await sp.web.lists.getByTitle(this.props.sourceListTitle).items.getById(row.Id as number).file.get();
  
          return moveFile(res.ServerRelativeUrl, `${this.props.destination.path}/${fileName}`, observer)
        }
      });
  
      Promise.all(promises).then((res) => {
        console.log("promises finished");
        observer.complete("Items moved successfully!");
        console.log(res);
      }).catch((err) => {
        console.log("promises errored");
        observer.error("Item move errored, please contact IT.");
        console.log(err);
      });
    });

    this.subscribeLog();
  }

}