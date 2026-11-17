import * as React from "react";
//import styles from "./PeopleorgchartComponent.module.scss";
import { PeopleOrgChart } from "./PeopleOrgChart";
import { ServiceScope } from "@microsoft/sp-core-library";

export interface IPeopleorgchartComponentProps {
  serviceScope: ServiceScope;
}

export default class PeopleorgchartComponent extends React.Component<IPeopleorgchartComponentProps> {
  public render(): React.ReactElement<IPeopleorgchartComponentProps> {
    return (
      <PeopleOrgChart serviceScope={this.props.serviceScope} />
    );
  }
}
