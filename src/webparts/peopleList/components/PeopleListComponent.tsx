import * as React from "react";
import styles from "./PeopleListComponent.module.scss";
import { ServiceScope } from "@microsoft/sp-core-library";
import { PeopleList } from "./PeopleList";
import { IEmployee } from "../../../model";
import { EmployeeService, IEmployeeService } from "../../../service";

export interface IPeopleListComponentProps {
  description: string;
  serviceScope: ServiceScope;
}

export interface IPeopleListComponentState {
  employees: IEmployee[];
}

export default class PeopleListComponent extends React.Component<IPeopleListComponentProps, IPeopleListComponentState> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  constructor(props: any) {
    super(props);
    this.state = { employees: [] };
  }

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  async componentDidMount() {
    // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
    const getEmployees = async () => {
      try {
        const employeeService: IEmployeeService = this.props.serviceScope.consume(EmployeeService.serviceKey);
        const empsResponse = await employeeService.getEmployees();
        if (empsResponse && empsResponse) {
          this.setState({ employees: empsResponse });

          //await employeeService.populateEmployeePhoto(empsResponse);
        }
      } catch (error) {
        console.error("Error getting employees: ", error);
      }
    };

    // Call the async function
    // eslint-disable-next-line no-void
    void getEmployees();
  }

  public render(): React.ReactElement<IPeopleListComponentProps> {
    return (
      <div className={`${styles.peopleListComponent}`}>
        <div className={styles.title}>Search People and View Org Chart</div>
        {this.state.employees && this.state.employees.length > 0 && <PeopleList serviceScope={this.props.serviceScope} employees={this.state.employees} />}
      </div>
    );
  }
}
