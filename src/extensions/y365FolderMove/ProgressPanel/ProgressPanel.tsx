import React, { useState, useEffect } from 'react';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { ISPHttpClientOptions, SPHttpClient } from '@microsoft/sp-http';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { sp } from '@pnp/sp';

import { IMoveCopyJobInfoContainer } from '../interfaces/IMoveCopyJobInfo';

import styles from './ProgressPanel.module.scss';

export interface IProgressPanelProps{
  context: ListViewCommandSetContext;
  isOpen: boolean;
  onDismissed: () => void;
}

export default function ProgressPanel(props: IProgressPanelProps) {
  const [jobs, setJobs]: [IMoveCopyJobInfoContainer[], (_: any) => void] = useState([]);

  useEffect(() => {
    getJobs();
  }, []);

  const getJobs = () => {
    setJobs([]);

    const storedJobsString = localStorage.getItem("y365RunningJobs");
    const storedJobs = JSON.parse(storedJobsString);

    if(storedJobs){
      setJobs(storedJobs);
      getJobLogs(storedJobs);
    }
  }

  const getJobLogs = async (jobs: IMoveCopyJobInfoContainer[]) => {
    const jobsWithLogs$ = jobs.map((job) => {
      return new Promise(async (resolve, reject) => {
        const messages = await getJobLog(job);
        resolve({
          ...job,
          messages
        });
      });
    });

    const jobsWithLogs = await Promise.all(jobsWithLogs$);
    console.log(jobsWithLogs);
    setJobs(jobsWithLogs);
  }

  const getJobLog = async (copyJobInfo: IMoveCopyJobInfoContainer) => {
    const { context } = props;
    const siteUrl = context.pageContext.web.absoluteUrl.replace(/\/sites\/.+/gmi, "");
    const body = copyJobInfo;

    const options: ISPHttpClientOptions = {
      method: "POST",
      headers: {
        "Accept":"application/json",
        "Content-Type": "application/json"
      },
      body: JSON.stringify(body)
    }

    try{
      const res = await context.spHttpClient.fetch(`${siteUrl}/_api/site/GetCopyJobProgress`, SPHttpClient.configurations.v1, options);
      const jobStatus = await res.json();
      return jobStatus.Logs;
    }
    catch(e){
      console.log(e);
      return ["ERROR"]
    }
  }

  const clearJobs = () => {
    localStorage.removeItem("y365RunningJobs");

    getJobs();
  }

  return (
    <Panel isOpen={ props.isOpen } type={ PanelType.medium } headerText="Shift Job Progress" className={ styles.shiftProgressPanel } onDismiss={ props.onDismissed }>
      <IconButton iconProps={{ iconName: "refresh" }} onClick={ getJobs } />
      <IconButton iconProps={{ iconName: "clear" }} onClick={ clearJobs } />
      <ul className={ styles.jobList }>
        {
          jobs.map((job) => {
            return (
              <li>
                <span>{ job.copyJobInfo.JobId }</span>
                {
                  job.messages && Array.isArray(job.messages) ? (
                    <ul className={ styles.jobList }>
                      {
                        job.messages.map((msg) => <li><pre>{ JSON.stringify(msg, null, " ") }</pre></li>)
                      }
                    </ul>
                  ) : <Spinner />
                }
              </li>       
            )
          })
        }
      </ul>
    </Panel>
  );
}