/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import { FC } from "react";
import { IEmployee } from "../../../model";
import { EmployeeService, IEmployeeService } from "../../../service";
import { ServiceScope } from "@microsoft/sp-core-library";
import { CoreOrgChart } from "./CoreOrgChart";

export interface IPeopleOrgChartProps {
  serviceScope: ServiceScope;
}

export const PeopleOrgChart: FC<IPeopleOrgChartProps> = (props) => {
  //const [employeesForOrgChart, setemployeesForOrgChart] = React.useState<IEmployee[]>([]);
  //const{employeesForOrgChart, setEmployeesForOrgChart} = useSharedContext();
  const [employees, setEmployees] = React.useState<IEmployee[]>([]);

      React.useEffect(() => {
        // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
      
        const getEmployees = async (): Promise<void> => {
          try {
            const employeeService: IEmployeeService = props.serviceScope.consume(EmployeeService.serviceKey);
            const empsResponse = await employeeService.getEmployees();
            if (empsResponse && empsResponse) {
              setEmployees(empsResponse);
              //await employeeService.populateEmployeePhoto(empsResponse);
            }
          } catch (error) {
            console.error("Error getting employees: ", error);
          }
        };

        // Call the async function
        // eslint-disable-next-line no-void
        void getEmployees();
      }, []);
  

  return <>
  {employees && <CoreOrgChart data={employees} serviceScope={props.serviceScope} showHighlighers={true} />}
  {/*{employees && <CoreOrgChart data={employees} selectedEmployee={undefined} serviceScope={props.serviceScope} showHighlighers={true} />} */}
  </>;
};
