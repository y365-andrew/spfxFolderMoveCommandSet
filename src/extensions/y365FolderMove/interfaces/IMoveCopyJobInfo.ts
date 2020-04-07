
export interface IMoveCopyJobInfoContainer{
    copyJobInfo: IMoveCopyJobInfo;
    [other: string]: any;
}

export interface IMoveCopyJobInfo{ 
    EncryptionKey: string;
    JobId: string;
    JobQueueId: string;
}