import { sp, IFolderAddResult, Files, SPBatch, IWeb } from '@pnp/sp/presets/all';
import { Observable, of } from 'rxjs';
import { concat, retry, timeInterval } from 'rxjs/operators';
import { Queryable } from '@pnp/odata';
import { BearerTokenFetchClient } from '@pnp/common'
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { SPHttpClient, SPHttpClientConfiguration, ISPHttpClientOptions } from '@microsoft/sp-http';

export interface ILogObserver{
  next: (msg: ILogNextData ) => void;
  error: (err: any) => void;
  complete: (msg: string) => void;
}

export interface ILogNextData{
  msgType: ELogNextDataMsgType,
  msg?: string,
  objectsProcessed?: string;
  objectsExpected?: string;
}

export enum ELogNextDataMsgType{
  "Log",
  "Progress",
  "Warning",
  "Error"
}

export class MoveLog{
  private observer: ILogObserver;
  constructor(observer?: ILogObserver){
    this.observer = observer;
  }

  public write = (msg: string) => {
    if(this.observer){
      this.observer.next({ msgType: ELogNextDataMsgType.Log, msg }); 
    }
  }

  public writeError = (msg: string) => {
    if(this.observer){
      this.observer.next({ msgType:ELogNextDataMsgType.Warning, msg })
    }
  }

  public writeProgress = (objectsProcessed: string, objectsExpected: string) => {
    if(this.observer){
      this.observer.next({ 
        msgType:ELogNextDataMsgType.Progress,
        objectsExpected,
        objectsProcessed 
      });
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
    sp.web.folders.add(`!@p1::${encodedDestination}`).then((folderAddResult: IFolderAddResult) => {
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
    const fileMoveRes = sp.web.getFileById(sourceId).moveByPath(encodedDestination, false);
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
        
        //return moveFilesAsBatch(context, subFile.UniqueId, subFile.ServerRelativeUrl, `${destinationFolderUrl}/${subFile.Name}`, logObserver, batches[batches.length -1])
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

export async function moveFilesAsBatch(context: ListViewCommandSetContext, sourceId: string, sourceServerRelativeUrl: string, destServerRelativeUrl: string, logObserver: ILogObserver, batch: SPBatch, retryCount?: number){
  const log = new MoveLog(logObserver);

  try{
    log.write(`Moving file ${sourceServerRelativeUrl}`);
    const encodedDestination = destServerRelativeUrl.split('/').map(v => encodeURIComponent(v).replace(/\'/g, "%27%27")).join('/');

    const fileMoveResult = await sp.web.getFileById(sourceId).inBatch(batch).moveByPath(destServerRelativeUrl, false);

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

export async function moveOrchestrator(context: ListViewCommandSetContext, sourceFolderId: string, sourceServerRelativeUrl: string, destinationServerRelativeUrl: string, destinationWeb: IWeb, logObserver: ILogObserver, retryCount?: number){
  const log: MoveLog = new MoveLog(logObserver);
  let newRetryCount = (retryCount || 0) + 1;
  const retryResetTimer = setTimeout(() => {
    newRetryCount = 0;
  }, 18000);

  log.write(`Creating destination: ${ destinationServerRelativeUrl }`);

  try{
    const encodedDestination = destinationServerRelativeUrl.split('/').map(v => encodeURIComponent(v).replace(/\'/g, "%27%27")).join('/');
    const { data: folderAddResult } = await destinationWeb.folders.add(`!@p1::${encodedDestination}`);
    const destinationFolderUrl = folderAddResult.ServerRelativeUrl;

    const subFolders = await sp.web.getFolderById(sourceFolderId).folders.select('ServerRelativeUrl,Name,UniqueId').get();

    for(let i = 0; i < subFolders.length; i++){
      const res = await createFolderInDestination(context, subFolders[i].UniqueId, `${destinationFolderUrl}/${subFolders[i].Name}`, destinationWeb, logObserver);
      console.log(res);
    }

    const fileMoveRes = await orchestrateFileBatches(context, sourceFolderId, destinationServerRelativeUrl, logObserver);
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
           moveOrchestrator(context, sourceFolderId, sourceServerRelativeUrl, destinationServerRelativeUrl, destinationWeb, logObserver, (newRetryCount + 1)).then((res) => { resolve(res); }).catch((err) => { reject(err)})
        }, 5000);
      });
    }
  }
}

export async function orchestrateFileBatches(context:ListViewCommandSetContext, folderId, destinationPath: string, logObserver){
  const log: MoveLog = new MoveLog(logObserver);
  const files = await sp.web.getFolderById(folderId).files.select('ServerRelativeUrl,Name,UniqueId').get();

  if(files){
    /*const batches: SPBatch[] = [];
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

    return fileMoveResults;*/
    const siteUrl = context.pageContext.web.absoluteUrl.replace(/\/sites\/.+/gmi, "")

    const exportObjectUris =  files.map((file) => `${siteUrl}${file.ServerRelativeUrl}`)

    const body = {
      exportObjectUris,
      destinationUri: `${siteUrl}${destinationPath}`,
      options: {
        IgnoreVersionHistory: false,
        IsMoveMode: true,
        AllowSchemaMismatch: true
      }
    }
    const options: ISPHttpClientOptions = {
      method: "POST",
      headers: {
        "Accept":"application/json",
        "Content-Type": "application/json"
      },
      body: JSON.stringify(body)
    }

    console.log(siteUrl);
    const createRes = await context.spHttpClient.fetch(`${siteUrl}/_api/site/CreateCopyJobs`, SPHttpClient.configurations.v1, options)
    console.log(createRes);
    return context.spHttpClient.fetch(`${siteUrl}/_api/site/GetCopyJobProgress`, SPHttpClient.configurations.v1, {method: "POST", headers:{"Acccept":"application/json"}})
  }
  else{
    return [];
  }
}

export async function createFolderInDestination(context: ListViewCommandSetContext, folderId, destinationPath: string, destinationWeb: IWeb, logObserver: ILogObserver, retryCount?: number){
    const log: MoveLog = new MoveLog(logObserver);
    const encodedDestination = destinationPath.split('/').map(v => encodeURIComponent(v).replace(/\'/g, "%27%27")).join('/');

    try{
      const { data: folderAddResult } = await destinationWeb.folders.add(`!@p1::${encodedDestination}`);
      log.write(`Folder created: ${destinationPath}`);
      const subFolders = await sp.web.getFolderById(folderId).folders.select('ServerRelativeUrl,Name,UniqueId').get();

      for(let i = 0; i < subFolders.length; i++){
        const subFolderPath = `${destinationPath}/${subFolders[i].Name}`;
        const res = await createFolderInDestination(context, subFolders[i].UniqueId, subFolderPath, destinationWeb, logObserver);
        continue;
      }

      log.write(`Moving files into ${ destinationPath }`)
      const fileMoveRes = await orchestrateFileBatches(context, folderId, destinationPath, logObserver);
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
                  createFolderInDestination(context, folderId, destinationPath, destinationWeb, logObserver, newRetryCount).then(res => { resolve(res) }).catch(err => { reject(err) });
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
                  createFolderInDestination(context, folderId, destinationPath, destinationWeb, logObserver, newRetryCount).then(res => { resolve(res) }).catch(err => { reject(err) });
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
              createFolderInDestination(context, folderId, destinationPath, destinationWeb, logObserver, newRetryCount).then(res => { resolve(res) }).catch(err => { reject(err) })
            }, 500)
          });

          return retryRes;
        }
      }
    }

}


// *** 2.0 *** //
// These new functions use the MoveCopyJob endpoint to schedule and manage the jobs //

// This function traverses the folder structure and moves only items which do not exist in the destination as no merge function exists currently //
export async function moveFilesAndMerge(context: ListViewCommandSetContext, sourceFolderId: string, destinationPath: string, destinationWeb: IWeb, logObserver: ILogObserver){
  const sourceFolder = sp.web.getFolderById(sourceFolderId);
  const sourceFolderName = await sourceFolder.select("Name").get();
  const sourceSubFolders = await sourceFolder.folders.select('Name','ServerRelativeUrl','UniqueId').get();
  const sourceSubFiles = await sourceFolder.files.select('Name','ServerRelativeUrl','UniqueId').get();

  const destFolder = destinationWeb.getFolderByServerRelativeUrl(destinationPath);
  const destSubFolders = await destFolder.folders.select('Name','ServerRelativeUrl','UniqueId').get();
  const destSubFiles = await destFolder.files.select('Name','ServerRelativeUrl','UniqueId').get();

  const sourceSubFolderPromises = await sourceSubFolders.map(async (sourceSubFolder) => {
    // Check if folder exists already
    const exists = destSubFolders.map(f => f.Name).indexOf(sourceSubFolder.Name) >= 0;
    const newDestinationPath = `${destinationPath}/${sourceFolderName.Name}`;
    console.log(newDestinationPath)
    console.log(`Exists: ${exists}`);
    if(exists){
      const res = await moveFilesAndMerge(context, sourceSubFolder.UniqueId, newDestinationPath, destinationWeb, logObserver);
      return res;
    }
    else{
      const res = await moveFilesAsCopyJob(context, sourceSubFolder.UniqueId, false, newDestinationPath, logObserver);
      return res;
    }
  });

  const sourceSubFilePromises = await sourceSubFiles.map(async (sourceSubFile) => {
    // Check if file exists already
    const exists = destSubFiles.map(f => f.Name).indexOf(sourceSubFile.Name) >= 0
    const newDestinationPath = `${destinationPath}/${sourceFolderName.Name}`;

    if(exists){
      return;
    }
    else{
      const res = await moveFilesAsCopyJob(context, sourceSubFile.UniqueId, true, newDestinationPath, logObserver);
      return res;
    }
  });

  return Promise.all([sourceSubFolderPromises, sourceSubFilePromises]);
}

export async function moveFilesAsCopyJob(context: ListViewCommandSetContext, sourceObjectId: string, isFile: boolean, destinationPath: string, logObserver: ILogObserver){
  const log: MoveLog = new MoveLog(logObserver);
  const siteUrl = context.pageContext.web.absoluteUrl.replace(/\/sites\/.+/gmi, "");
  const sourceObject = isFile ? await sp.web.getFileById(sourceObjectId).select('ServerRelativeUrl,Name,UniqueId').get() : await sp.web.getFolderById(sourceObjectId).select('ServerRelativeUrl,Name,UniqueId').get();
  const exportObjectUri =  `${siteUrl}${sourceObject.ServerRelativeUrl}`; //subFolders.map((folder) => `${siteUrl}${folder.ServerRelativeUrl}`);

  const destinationUri = (destinationPath.match(/(https:\/\/).+/g) && destinationPath.match(/(https:\/\/).+/g).length >= 0) ? destinationPath : `${siteUrl}${destinationPath}`;

  const body = {
    exportObjectUris: [exportObjectUri],
    destinationUri,
    options: {
      IgnoreVersionHistory: false,
      IsMoveMode: true,
      AllowSchemaMismatch: true,
      NameConflictBehavior: 0 // 0 = Fail, 1 = Replace, 2 = Rename -- add a UI element for this later
    }
  }

  const options: ISPHttpClientOptions = {
    method: "POST",
    headers: {
      "Accept":"application/json",
      "Content-Type": "application/json"
    },
    body: JSON.stringify(body)
  }

  console.log(siteUrl);
  const createRes = await (await context.spHttpClient.fetch(`${siteUrl}/_api/site/CreateCopyJob`, SPHttpClient.configurations.v1, options)).json()
  console.log(createRes);

  if(createRes.error){
    throw new Error(createRes.error.message)
  }

  let progress = await getCopyJobProgress(context, siteUrl, createRes, logObserver);
  console.log(progress);

  return progress;

}

export async function getCopyJobProgress(context: ListViewCommandSetContext, siteUrl: string, job: any, logObserver: ILogObserver){
  const log: MoveLog = new MoveLog(logObserver);
  let nextCheckInMs = 5000;

  const runAgain =  (timerValue) => new Promise((resolve, reject) => {
    setTimeout(() => {
      resolve(getCopyJobProgress(context, siteUrl, job, logObserver));
    }, timerValue);
  });

  const body = {
    copyJobInfo:{
      EncryptionKey: job.EncryptionKey,
      JobId: job.JobId,
      JobQueueUri: job.JobQueueUri
    }
  };

  const options: ISPHttpClientOptions = {
    method: "POST",
    headers: {
      "Accept":"application/json",
      "Content-Type": "application/json"
    },
    body: JSON.stringify(body)
  }
  
  const jobProgress = await (await context.spHttpClient.fetch(`${siteUrl}/_api/site/GetCopyJobProgress`, SPHttpClient.configurations.v1, options)).json()
  console.log(jobProgress);

  // Check if queued or active
  if(jobProgress.JobState === 2){
    log.write("Job queued");

    return runAgain(nextCheckInMs)
  }

  if(jobProgress  && jobProgress.Logs && jobProgress.Logs.length > 0){
    console.log(jobProgress.Logs.map(log => JSON.parse(log).Event))
    
    for(let i = 0; i < jobProgress.Logs.length; i++){
      const logObj = JSON.parse(jobProgress.Logs[i]);
      switch(logObj.Event){
        case "JobFatalError":{
          throw new Error(logObj.Message)
        }
        case "JobError":{
          log.write(logObj.Message);
          break;
        }
        case "JobQueued": {
          log.write("Job queued");

          const storedJobsString = localStorage.getItem("y365RunningJobs");
          const storedJobs = storedJobsString ? JSON.parse(storedJobsString) : false;
          localStorage.setItem("y365RunningJobs", JSON.stringify( storedJobs ? [...storedJobs, body] : [body] ));
          break;
        }
        case "JobLogFileCreate": {
          log.write("Job log file created");
          break;
        }
        case "JobProgress":{
          log.writeProgress(logObj.ObjectsProcessed, logObj.TotalExpectedSPObjects);
          break;
        }
        case "JobStart": {
          nextCheckInMs = 1000;

          switch(logObj.MigrationDirection){
            case "Export": {
              log.write("Exporting files from source");
              break;
            }
            case "Import": {
              log.write("Importing files to destination");
              break;
            }
            case "Cleanup": {
              log.write("Cleaning up");
              break;
            }
          }

          break;
        }
        case "JobFinishedObjectInfo": {
          nextCheckInMs = 1000;

          const storedJobsString = localStorage.getItem("y365RunningJobs");
          const storedJobs = storedJobsString ? JSON.parse(storedJobsString) : false;
          if(storedJobs){
            const newStoredJobs = storedJobs.filter((j) => j != body);
            if(newStoredJobs.length > 0){
              localStorage.setItem("y365RunningJobs", JSON.stringify(storedJobs))
            }
            else{
              localStorage.removeItem("y365RunningJobs")
            }
          }
          else{
            localStorage.removeItem("y365RunningJobs")
          }
          
          break;
        }
        case "JobEnd": {
          nextCheckInMs = 1000;

          switch(logObj.MigrationDirection){
            case "Export": {
              log.write(`Export completed. Processed ${ logObj.ObjectsProcessed } objects (${ logObj.BytesProcessed } bytes) in ${ logObj.TotalDurationInMs }`);
              break;
            }
            case "Import": {
              log.write(`Import completed. Processed ${ logObj.ObjectsProcessed } objects (${ logObj.BytesProcessed } bytes) in ${ logObj.TotalDurationInMs }`);
              break;
            }
            case "MoveCleanup":
            case "Cleanup": {
              log.write("Cleaning up");

              return jobProgress
            }
          }

          break;
        }
      }

    };

  }

  return runAgain(nextCheckInMs);
}

// The below aren't used, they exist to help understand the copy job logging //

export enum ICopyJobLogEventJobType{
  "JobQueued",
  "JobLogFileCreate",
  "JobStart",
  "JobFinishedObjectInfo",
  "JobEnd",
  "JobProgress",
  "JobError",
  "JobWarning",
  "JobFatalError"
}

export enum ICopyJobLogEventMigrationDirection{
  "Export",
  "Import",
  "Cleanup",
  "MoveCleanup"
}

export enum EJobState{
  "Queued" = 2,
  "Running" = 4,
  "Other" = 0
}

export interface ICopyJobLog{
  BytesProcessed: string;
  CorrelationId: string;
  CpuDurationInMs: string;
  CreatedOrUpdatedFileStatsBySize: string;
  Event: ICopyJobLogEventJobType;
  FilesCreated: string;
  JobId: string;
  MigrationDirection: ICopyJobLogEventMigrationDirection;
  MigrationType: string;
  ObjectsProcessed: string;
  ObjectsStatsByType: string;
  //"{\"SPUser\":{\"Count\":1,\"TotalTime\":0,\"AccumulatedVersions\":0,\"ObjectsWithVersions\":0},\"SPFolder\":{\"Count\":1,\"TotalTime\":462,\"AccumulatedVersions\":0,\"ObjectsWithVersions\":0},\"SPListItem\":{\"Count\":1,\"TotalTime\":1248,\"AccumulatedVersions\":0,\"ObjectsWithVersions\":0}}";
  SqlDurationInMs: string;
  SqlQueryCount: string;
  Time: string;
  TotalDurationInMs: string;
  TotalErrors: string;
  TotalExpectedBytes: string;
  TotalExpectedSPObjects: string;
  TotalRetryCount: string;
  TotalWarnings: string;
  WaitTimeOnSqlThrottlingMilliseconds: string;
}
