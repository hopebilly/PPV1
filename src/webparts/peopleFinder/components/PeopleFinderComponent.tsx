import * as React from 'react';
import styles from '../components/PeopleFinderComponent.module.scss';
import { ServiceScope } from "@microsoft/sp-core-library";
//import { IEmployee } from "../../../model";
//import { makeStyles } from '@material-ui/core';
import { PeopleFinder } from './PeopleFinder';

export interface IPeopleFinderComponentProps {
  description: string;
  serviceScope: ServiceScope;
  customText: string ;
}

export interface IPeopleFinderComponentState {
  
}

export default class PeopleFinderComponent extends React.Component<IPeopleFinderComponentProps, IPeopleFinderComponentState> {

  private styleTagId = 'cf-fullwidth-style';

  constructor(props: IPeopleFinderComponentProps) {
    super(props);
  }

  public componentDidMount(): void {
    const stale = document.getElementById(this.styleTagId);
    if (stale) stale.remove();
  
    const style = document.createElement('style');
    style.id = this.styleTagId;
    style.innerHTML = `
      /* ===== PeopleFinderComponent Full Width ===== */
  
      /* 1. Hide left nav */
      .sp-pageLayout-sideNav [class^='deferredLeftNav'],
      #spLeftNav {
        display: none !important;
        max-width: 0 !important;
      }
  
      /* 2. Expand layout containers */
      .sp-pageLayout-sideNav .SPCanvas-canvas,
      .CanvasZone.row.CanvasZone--alignment.CanvasZone--read.CanvasZone--noMargin,
      .CanvasZoneContainer:first-child .CanvasZone.row,
      .CanvasSection,
      .ControlZone {
        max-width: 100% !important;
        width: 100% !important;
        margin: 0 !important;
        padding: 0 25px !important;
      }
 
      .ControlZone--control{
        max-width: 100% !important;
        width: 100% !important;
        margin: 0 !important;
        padding: 35px 0 !important; /* add vertical padding for better breathing room */
      }
  
      /* 3. Fix upper layout compression (important) */
      .CanvasZoneContainer:first-child .CanvasZone.row:first-of-type {
        padding-top: 24px !important;
        padding-bottom: 12px !important;
      }
  
      /* 4. Remove canvas limits */
      div[data-automation-id='CanvasLayout'] > div > div,
      div[data-automation-id='CanvasZone'] > div {
        max-width: none !important;
      }
  
      /* 5. Web part container styling */
      .webPartContainer > div {
        padding-left: 25px !important;
      }
  
      .webPartContainer > div > div {
        min-height: 50px !important;
        padding-top: 15px !important;
      } */
    `;
  
    document.head.appendChild(style);
    window.dispatchEvent(new Event('resize'));
  }
  
  public componentWillUnmount(): void {
    const style = document.getElementById(this.styleTagId);
    if (style) {
      style.remove();
      window.dispatchEvent(new Event('resize'));
    }
  }

  public render(): React.ReactElement<IPeopleFinderComponentProps> {
    
    return (
      <>
         <div className={`${styles.peopleFinderComponent}`}>
           <div> 
              <PeopleFinder serviceScope={this.props.serviceScope} />
          </div>
        </div>
      </>
    );
  }
}
