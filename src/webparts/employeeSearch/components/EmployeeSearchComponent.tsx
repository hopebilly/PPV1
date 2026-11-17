import * as React from 'react';
//import styles from './EmployeeSearchComponent.module.scss';
import { FilterableTable }  from './EmployeeSearch';
import { ServiceScope } from "@microsoft/sp-core-library";
import { IEmployee } from "../../../model";
import { EmployeeService, IEmployeeService } from "../../../service";
//import {sampleEmployees} from "../../../service/EmployeeSampleData2";

export interface IEmployeeSearchComponentProps {
  serviceScope: ServiceScope;
}

export interface IEmployeeSearchComponentState {
  employees: IEmployee[];
}

export default class EmployeeSearchComponent extends React.Component<IEmployeeSearchComponentProps, IEmployeeSearchComponentState> {
    
  constructor(props: IEmployeeSearchComponentProps){
      super(props);
      this.state = { employees: [] };
  }

  getEmployees = async (): Promise<void> => {

    try {
      const employeeService: IEmployeeService = this.props.serviceScope.consume(EmployeeService.serviceKey);
      const empsResponse = await employeeService.getEmployees();
      //Sort employee list based on Employee name
      //empsResponse?.sort((a, b) => (a.EmployeeName?a.EmployeeName:" ").localeCompare(b.EmployeeName?b.EmployeeName:" ")); 
      //Update employee name column display in FN-LN format
      
      if (empsResponse && empsResponse.length > 0) {
        this.setState({ employees: empsResponse });

        //await employeeService.populateEmployeePhoto(empsResponse);
        //Sort employee list based on Employee name
        empsResponse?.sort((a, b) => (a.EmployeeName?a.EmployeeName:" ").localeCompare(b.EmployeeName?b.EmployeeName:" ")); 
        
        //Update employee name column display in FN-LN format
        const formattedEmpsResponse: IEmployee[] = [];
        empsResponse.forEach((value: IEmployee) => {
          const item: IEmployee = value;
          item.EmployeeName = value.EmployeeName? value.EmployeeName.split(",", 2).reverse().join(" "): "";
          formattedEmpsResponse.push(item);
        }); 
     
        this.setState({ employees: formattedEmpsResponse });
      }
    } catch (error) {
      console.error("Error getting employees: ", error);
    }
  };

  async componentDidMount(): Promise<void> {

    await this.getEmployees();
  
  }
  
  public render(): React.ReactElement<IEmployeeSearchComponentProps> {
      
       return (
          <>
            { this.state.employees && this.state.employees.length > 0 && <FilterableTable serviceScope={this.props.serviceScope} employees={this.state.employees} />}
          </>
       );
  }   
}