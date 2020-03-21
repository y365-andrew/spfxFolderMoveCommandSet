import { sp, FolderAddResult, Files, SPBatch } from '@pnp/sp';
import { Observable, of } from 'rxjs';
import { concat, retry } from 'rxjs/operators';
import { Queryable } from '@pnp/odata';

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

export function moveFolder(sourceId: string, sourceServerRelativeUrl: string, destinationServerRelativeUrl: string, logObserver?: ILogObserver): Promise<any>{
  const log = new MoveLog(logObserver);
  // Create result object for debugging
  return new Promise((resolve, reject) => {
    let result: object = {};

    log.write(`Moving folder ${sourceServerRelativeUrl}`);
    // Create folder in destination
    const encodedDestination = destinationServerRelativeUrl.split('/').map(v => encodeURIComponent(v).replace(/\'/g, "%27%27")).join('/');
    console.log(encodedDestination);
    sp.web.folders.add(`!@p1::${encodedDestination}`).then((folderAddResult: FolderAddResult) => {
      // Add result
      result = {
        ...result, 
        folderAddResult: folderAddResult.data 
      };

      const destinationFolderUrl = folderAddResult.data.ServerRelativeUrl;
      const sourceFolder = sp.web.getFolderById(sourceId);
      // Get contents
      // Folders
      sourceFolder.folders.select('ServerRelativeUrl,Name,UniqueId').get().then((subFolders) => {
        let subFolderPromises: Promise<any>[] = subFolders.map((subFolder) => {
          return moveFolder(subFolder.UniqueId, subFolder.ServerRelativeUrl, `${destinationFolderUrl}/${subFolder.Name}`, logObserver);
        });
        // Files
        sourceFolder.files.select('ServerRelativeUrl,Name,UniqueId').get().then((subFiles) => {
          let subFilePromises = subFiles.map((subFile) => {
            return moveFile(subFile.UniqueId, subFile.ServerRelativeUrl, `${destinationFolderUrl}/${subFile.Name}`, logObserver);
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

export async function moveFile(sourceId: string, sourceServerRelativeUrl: string, destServerRelativeUrl: string, logObserver: ILogObserver, retryCount?: number): Promise<any>{
  const log = new MoveLog(logObserver);
  // Move file
  log.write(`Moving file ${sourceServerRelativeUrl}`)
  const encodedDestination = destServerRelativeUrl.split('/').map(v => encodeURIComponent(v).replace(/\'/g, "%27%27")).join('/');
  console.log(encodedDestination);
  
  try{
    const fileMoveRes = sp.web.getFileById(sourceId).moveTo(`!@p1::${encodedDestination}`);
    const res = {
      file: sourceServerRelativeUrl,
      destination: destServerRelativeUrl,
      moved: true,
      fileMoveRes
    }
    return res;
  }
  catch(err){
    const newRetryCount = (retryCount || 0) + 1;
    console.log(err);
    if(newRetryCount > 5){
      throw new Error(`Retry count exceeded, error moving folder. Error message: ${err && err.Message ? err.Message : "Unknown"}`)
    }
    else{
      if(err && err.status){
        switch(err.status){
          //Server error
          case "500":{
            console.log("Retrying due to throttling", err);
            const retryRes = await new Promise((resolve, reject) => {
              setTimeout(() => {
                moveFile(sourceId, sourceServerRelativeUrl ,destServerRelativeUrl, logObserver, newRetryCount).then(res => { resolve(res) }).catch(err => { reject(err) })
              }, 100)
            });
            return retryRes;
          }
          //Throttle cases
          case "503":
          case "429":{
            console.log("Retrying due to throttling", err);
            const retryRes = await new Promise((resolve, reject) => {
              setTimeout(() => {
                moveFile(sourceId, sourceServerRelativeUrl ,destServerRelativeUrl, logObserver, newRetryCount).then(res => { resolve(res) }).catch(err => { reject(err) })
              }, 10000)
            });
            return retryRes;
          }
          //Other
          default:{
            throw new Error(err)
          }
        }
      }
      else{
        log.write(`Moving file: ${sourceServerRelativeUrl}`);
        console.log("Retrying due to unknown error code", err);
        const retryRes = await new Promise((resolve, reject) => {
          setTimeout(() => {
            moveFile(sourceId, sourceServerRelativeUrl ,destServerRelativeUrl, logObserver, newRetryCount).then(res => { resolve(res) }).catch(err => { reject(err) })
          }, 500)
        });

        return retryRes;
      }
    }
  }
}

export async function moveFolder2(sourceFolderId: string, sourceServerRelativeUrl: string, destinationServerRelativeUrl: string, logObserver: ILogObserver, requestCount?: number): Promise<any>{
  const reqCount = requestCount || 1;
  const log: MoveLog = new MoveLog(logObserver);
  let result: object = {};

  if(reqCount >= 100){
    console.log("Waiting")
    console.timeStamp(`Waiting 10 seconds.${ sourceServerRelativeUrl }`)
    setTimeout(() => {
      console.timeStamp(`Timeout Finished: ${ sourceServerRelativeUrl }`)
      moveFolder2(sourceFolderId, sourceServerRelativeUrl, destinationServerRelativeUrl, logObserver, 0);
    }, 10000)
  }
  else{
    log.write(`Moving folder ${sourceServerRelativeUrl}`);
    
    try{
      const encodedDestination = destinationServerRelativeUrl.split('/').map(v => encodeURIComponent(v).replace(/\'/g, "%27%27")).join('/');
      const { data: folderAddResult } = await sp.web.folders.add(`!@p1::${encodedDestination}`);
  
      result = {
        ...result,
        folderAddResult
      }
  
      const destinationFolderUrl = folderAddResult.ServerRelativeUrl;
      const sourceFolder = sp.web.getFolderById(sourceFolderId);
  
      const subFolders: any = await sourceFolder.folders.select('ServerRelativeUrl,Name,UniqueId').get();
  
      const subFiles = await sourceFolder.files.select('ServerRelativeUrl,Name,UniqueId').get();
  
      const subFolderResults: any = await Promise.all(subFolders.map((subFolder, i) => {
        return moveFolder2(subFolder.UniqueId, subFolder.ServerRelativeUrl, `${destinationFolderUrl}/${subFolder.Name}`, logObserver, (reqCount + i + 1));
      }));
  
      const batches: SPBatch[] = [];
  
      const subFileResultsPromises = subFiles.map((subFile, i) => {
        if(i === 0 || i % 20 === 0){
          const batch = sp.web.createBatch();
          batches.push(batch);
        }
        
        return moveFilesAsBatch(subFile.UniqueId, subFile.ServerRelativeUrl, `${destinationFolderUrl}/${subFile.Name}`, logObserver, batches[batches.length -1])
      });
  
      const batchResults = await Promise.all(batches.map(batch => {
        return batch.execute();
      }));
      
      const subFileResults: any = await Promise.all(subFileResultsPromises);
  
      let folderDeleteResult
  
      if(folderAddResult && (!subFileResults.err && !subFolderResults.err)){
        folderDeleteResult = await sourceFolder.delete();
      }
      
      result = {
        ...result,
        folderDeleteResult,
        subFolderResults,
        subFileResults
      }
  
      return result;
    }
    catch(err){
      console.log(err);
      log.write(`Error moving folder: ${sourceServerRelativeUrl}`);
  
      result = {
        ...result,
        err
      }
      console.log(result);
      throw new Error(err);
      //return result;
    }
  }

}

export async function moveFilesAsBatch(sourceId: string, sourceServerRelativeUrl: string, destServerRelativeUrl: string, logObserver: ILogObserver, batch: SPBatch, retryCount?: number){
  const log = new MoveLog(logObserver);

  try{
    log.write(`Moving file ${sourceServerRelativeUrl}`);
    const encodedDestination = destServerRelativeUrl.split('/').map(v => encodeURIComponent(v).replace(/\'/g, "%27%27")).join('/');
    const fileMoveResult = await sp.web.getFileById(sourceId).inBatch(batch).moveTo(`!@p1::${encodedDestination}`);

    return {
      file: sourceServerRelativeUrl,
      fileMoveResult,
      destination: destServerRelativeUrl,
      moved: true
    }
  }
  catch(err){
    const newRetryCount = (retryCount || 0) + 1;
    console.log(newRetryCount, err);
    console.log(destServerRelativeUrl);
    if(newRetryCount > 5){
      throw new Error(`Retry count exceeded, error moving folder. Error message: ${err && err.Message ? err.Message : "Unknown"}`)
    }
    else{
      if(err && err.status){
        switch(err.status){
          //Server error
          case "500":{
            console.log("Retrying due to throttling", err);
            const retryRes = await new Promise((resolve, reject) => {
              setTimeout(() => {
                moveFile(sourceId, sourceServerRelativeUrl ,destServerRelativeUrl, logObserver, newRetryCount).then(res => { resolve(res) }).catch(err => { reject(err) })
              }, 100)
            });
            return retryRes;
          }
          //Throttle cases
          case "503":
          case "429":{
            console.log("Retrying due to throttling", err);
            const retryRes = await new Promise((resolve, reject) => {
              setTimeout(() => {
                moveFile(sourceId, sourceServerRelativeUrl ,destServerRelativeUrl, logObserver, newRetryCount).then(res => { resolve(res) }).catch(err => { reject(err) })
              }, 10000)
            });
            return retryRes;
          }
          //Other
          default:{
            throw new Error(err)
          }
        }
      }
      else{
        log.write(`Moving file: ${sourceServerRelativeUrl}`);
        console.log("Retrying due to unknown error code", err);
        const retryRes = await new Promise((resolve, reject) => {
          setTimeout(() => {
            moveFile(sourceId, sourceServerRelativeUrl ,destServerRelativeUrl, logObserver, newRetryCount).then(res => { resolve(res) }).catch(err => { reject(err) })
          }, 500)
        });

        return retryRes;
      }
    }
/*
    return {
      file: sourceServerRelativeUrl,
      destination: destServerRelativeUrl,
      moved: false,
      err
    }*/
  }
}

export async function moveOrchestrator(sourceFolderId: string, sourceServerRelativeUrl: string, destinationServerRelativeUrl: string, logObserver: ILogObserver, retryCount?: number){
  const log: MoveLog = new MoveLog(logObserver);
  let newRetryCount = (retryCount || 0) + 1;
  const retryResetTimer = setTimeout(() => {
    newRetryCount = 0;
  }, 18000);

  log.write(`Creating destination: ${ destinationServerRelativeUrl }`);

  try{
    const encodedDestination = destinationServerRelativeUrl.split('/').map(v => encodeURIComponent(v).replace(/\'/g, "%27%27")).join('/');
    const { data: folderAddResult } = await sp.web.folders.add(`!@p1::${encodedDestination}`);
    const destinationFolderUrl = folderAddResult.ServerRelativeUrl;

    const subFolders = await sp.web.getFolderById(sourceFolderId).folders.select('ServerRelativeUrl,Name,UniqueId').get();

    for(let i = 0; i < subFolders.length; i++){
      const res = await createFolderInDestination(subFolders[i].UniqueId, `${destinationFolderUrl}/${subFolders[i].Name}`, logObserver);
      console.log(res);
    }

    const fileMoveRes = await orchestrateFileBatches(sourceFolderId, destinationServerRelativeUrl, logObserver);
    console.log(fileMoveRes);

    const folderDeleteRes = await sp.web.getFolderById(sourceFolderId).delete();
    console.log(folderDeleteRes);

    return {
      ...fileMoveRes
    }
  }
  catch(err){
    clearTimeout(retryResetTimer);
    
    if(newRetryCount > 3){
      console.log(err);
      return err;
    }
    else{
      console.log(`An error ocurred, retrying move (retry ${ newRetryCount } of 3)`);
      return await new Promise((resolve, reject) => {
        setTimeout(() => {
           moveOrchestrator(sourceFolderId, sourceServerRelativeUrl, destinationServerRelativeUrl, logObserver, (newRetryCount + 1)).then((res) => { resolve(res); }).catch((err) => { reject(err)})
        }, 5000);
      });
    }
  }
}

export async function orchestrateFileBatches(folderId, destinationPath: string, logObserver){
  const log: MoveLog = new MoveLog(logObserver);
  const files = await sp.web.getFolderById(folderId).files.select('ServerRelativeUrl,Name,UniqueId').get();

  if(files){
    const batches: SPBatch[] = [];
    const fileMoveResults = Promise.all(files.map((file, i) => {
      if(i % 100 === 0){
        const batch = sp.web.createBatch();
        batches.push(batch);
      }

      return moveFilesAsBatch(file.UniqueId, file.ServerRelativeUrl, `${destinationPath}/${file.Name}`, logObserver, batches[batches.length -1]);
    }));

    await Promise.all(batches.map((batch, i, batchArr) => {
      log.write(`Moving ${ destinationPath } files in batch (${ i+1 } of ${ batchArr.length }).`);
      return batch.execute();
    }));

    return fileMoveResults;
  }
  else{
    return [];
  }
}

export async function createFolderInDestination(folderId, destinationPath: string, logObserver: ILogObserver, retryCount?: number){
    const log: MoveLog = new MoveLog(logObserver);
    const encodedDestination = destinationPath.split('/').map(v => encodeURIComponent(v).replace(/\'/g, "%27%27")).join('/');

    try{
      const { data: folderAddResult } = await sp.web.folders.add(`!@p1::${encodedDestination}`);
      log.write(`Folder created: ${destinationPath}`);
      const subFolders = await sp.web.getFolderById(folderId).folders.select('ServerRelativeUrl,Name,UniqueId').get();

      for(let i = 0; i < subFolders.length; i++){
        const subFolderPath = `${destinationPath}/${subFolders[i].Name}`;
        const res = await createFolderInDestination(subFolders[i].UniqueId, subFolderPath, logObserver);
        continue;
      }

      log.write(`Moving files into ${ destinationPath }`)
      const fileMoveRes = await orchestrateFileBatches(folderId, destinationPath, logObserver);
      console.log(fileMoveRes);
      
      log.write(`Branch Complete: ${ destinationPath }`);
      
      const folderDeleteRes = await sp.web.getFolderById(folderId).delete();
      console.log(folderDeleteRes);

      return `Branch Complete: ${ destinationPath }`;
    }
    catch(err){
      console.log(err);
      console.log(destinationPath);
      const newRetryCount = (retryCount || 0) + 1;

      if(newRetryCount > 5){
        throw new Error(`Retry count exceeded, error moving folder. Error message: ${err && err.Message ? err.Message : "Unknown"}`)
      }
      else{
        if(err && err.status){
          switch(err.status){
            //Server error
            case "500":{
              console.log("Retrying due to throttling", err);
              const retryRes = await new Promise((resolve, reject) => {
                setTimeout(() => {
                  createFolderInDestination(folderId, destinationPath, logObserver, newRetryCount).then(res => { resolve(res) }).catch(err => { reject(err) });
                }, 100)
              });
              return retryRes;
            }
            //Throttle cases
            case "503":
            case "429":{
              console.log("Retrying due to throttling", err);
              const retryRes = await new Promise((resolve, reject) => {
                setTimeout(() => {
                  createFolderInDestination(folderId, destinationPath, logObserver, newRetryCount).then(res => { resolve(res) }).catch(err => { reject(err) });
                }, 10000)
              });
              return retryRes;
            }
            //Other
            default:{
              throw new Error(err)
            }
          }
        }
        else{
          console.log("Retrying due to unknown error code", err);
          console.log(destinationPath);
          const retryRes = await new Promise((resolve, reject) => {
            setTimeout(() => {
              createFolderInDestination(folderId, destinationPath, logObserver,newRetryCount).then(res => { resolve(res) }).catch(err => { reject(err) })
            }, 500)
          });

          return retryRes;
        }
      }
    }

}