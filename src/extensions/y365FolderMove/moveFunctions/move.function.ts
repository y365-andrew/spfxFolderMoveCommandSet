import { sp, FolderAddResult, Files } from '@pnp/sp';
import { Observable, of } from 'rxjs';
import { concat } from 'rxjs/operators';

export interface ILogObserver{
  next: (msg: string) => void;
  error: (err: any) => void;
  complete: (msg: string) => void;
}

export class MoveLog{
  private observer: ILogObserver;
  constructor(observer?: ILogObserver){
    this.observer = observer;
  }

  public write = (msg) => {
    if(this.observer){ 
      this.observer.next(msg); 
    }
  }
}

export function moveFolder(sourceServerRelativeUrl: string, destinationServerRelativeUrl: string, logObserver?: ILogObserver): Promise<any>{
  const log = new MoveLog(logObserver);
  // Create result object for debugging
  return new Promise((resolve, reject) => {
    let result: object = {};

    log.write(`Moving folder ${sourceServerRelativeUrl}`);
    // Create folder in destination
    sp.web.folders.add(destinationServerRelativeUrl).then((folderAddResult: FolderAddResult) => {
      // Add result
      result = {
        ...result, 
        folderAddResult: folderAddResult.data 
      };

      const destinationFolderUrl = folderAddResult.data.ServerRelativeUrl;
      const sourceFolder = sp.web.getFolderByServerRelativeUrl(sourceServerRelativeUrl);
      // Get contents
      // Folders
      sourceFolder.folders.select('ServerRelativeUrl,Name').get().then((subFolders) => {
        let subFolderPromises: Promise<any>[] = subFolders.map((subFolder) => {
          return moveFolder(subFolder.ServerRelativeUrl, `${destinationFolderUrl}/${subFolder.Name}`, logObserver);
        });
        // Files
        sourceFolder.files.select('ServerRelativeUrl,Name').get().then((subFiles) => {
          let subFilePromises = subFiles.map((subFile) => {
            return moveFile(subFile.ServerRelativeUrl, `${destinationFolderUrl}/${subFile.Name}`, logObserver);
          });

          const subPromises = [...subFolderPromises, ...subFilePromises];

          // Sub folder/files have sucessfully completed
          Promise.all(subPromises).then((subResult) => {
            result = {
              ...result,
              subResult
            };
            // Delete Folder
            sourceFolder.delete().then(() => {
              result = {
                ...result, 
                folderDeleted: true
              };
              // Resolve promise
              resolve(result);
            }).catch((folderDeleteError) => {
              result = {
                ...result,
                folderDeleted: false,
                folderDeleteError
              };

              reject(result);
            });
          }).catch((subError) => {
            result = {
              ...result,
              subError
            };

            reject(result);
          });
        }).catch((subFilesError) => {
          result = {
            ...result,
            subFilesError
          };

          reject(result);
        });
      }).catch((subFoldersError) => {
        result = {
          ...result,
          subFoldersError
        };

        reject(result);
      });
      
    }).catch((folderAddError) => {
      result = {
        ...result,
        folderAddError
      };

      reject(result);
    });
  });
}

export function moveFile(sourceServerRelativeUrl: string, destServerRelativeUrl: string, logObserver?: ILogObserver): Promise<any>{
  const log = new MoveLog(logObserver);

  return new Promise((resolve, reject) => {
    // Move file
    log.write(`Moving file ${sourceServerRelativeUrl}`)
    sp.web.getFileByServerRelativeUrl(sourceServerRelativeUrl).moveTo(destServerRelativeUrl).then(() => {
      const res = {
        file: sourceServerRelativeUrl,
        destination: destServerRelativeUrl,
        moved: true
      }

      resolve(res);
    }).catch((fileMoveErr) => {
      const res = {
        file: sourceServerRelativeUrl,
        destination: destServerRelativeUrl,
        moved: false,
        fileMoveErr
      }

      reject(res);
    })
  });

  // Handle errors + retry policy
}