import * as React from 'react';
import {OrganizationChartWrapper} from './OrganizationChartWrapper';
import { ServiceScope } from "@microsoft/sp-core-library";
//import { SharedProvider } from '../../../contexts/SharedContext';
require("../../../../node_modules/orgchart/src/css/jquery.orgchart.css");
require("../../../../node_modules/font-awesome/css/font-awesome.min.css");

export interface IOrganizationChartComponentProps {
  serviceScope: ServiceScope; 
}


export default class OrganizationChartComponent extends React.Component<IOrganizationChartComponentProps> {
  public render(): React.ReactElement<IOrganizationChartComponentProps> {
      return (
            <OrganizationChartWrapper serviceScope={this.props.serviceScope} Directorate='' />
    );
  }
}