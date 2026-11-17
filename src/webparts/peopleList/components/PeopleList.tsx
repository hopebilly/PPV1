// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as d3 from "d3";
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as flextree from "d3-flextree";
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import { OrgChart } from "d3-org-chart";
import { DefaultButton, DetailsList, DetailsListLayoutMode, IColumn, Link, MessageBar, MessageBarType, Panel, PanelType, SelectionMode, TextField } from "@fluentui/react";
import { IEmployee } from "../../../model";
import * as React from "react";
import { FC } from "react";
import { ServiceScope } from "@microsoft/sp-core-library";
import { cloneDeep } from "@microsoft/sp-lodash-subset";
import { CoreOrgChart } from "../../peopleorgchart/components/CoreOrgChart";
import styles from "./PeopleList.module.scss";

export interface IPeopleListProps {
  serviceScope: ServiceScope;
  employees: IEmployee[];
}

export const PeopleList: FC<IPeopleListProps> = (props) => {
  const [employees, setEmployees] = React.useState<IEmployee[]>(props.employees);
  const [filteredEmployees, setFilteredEmployees] = React.useState<IEmployee[]>(props.employees);
  const [employeesForOrgChart, setEmployeesForOrgChart] = React.useState<IEmployee[]>(props.employees);
  const [selectedEmployee, setSelectedEmployee] = React.useState<IEmployee>();
  const [message, setMessage] = React.useState<string>("");
  const [searchFilter, setSearchFilter] = React.useState<string>("");
  const initialColumns: IColumn[] = [
    {
      key: "OrgChart",
      name: "",
      fieldName: "OrgChart",
      minWidth: 50,
      maxWidth: 50,
    },
    {
      key: "EmployeeName",
      name: "Employee Name",
      fieldName: "EmployeeName",
      minWidth: 150,
      maxWidth: 200,
      isRowHeader: true,
      isResizable: true,
      isSorted: true,
      isSortedDescending: false,
      sortAscendingAriaLabel: "Sorted A to Z",
      sortDescendingAriaLabel: "Sorted Z to A",
      onColumnClick: _onColumnClick,
      data: "string",
      isPadded: true,
    },
    {
      key: "EmpO365Email",
      name: "Email",
      fieldName: "EmpO365Email",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onColumnClick: _onColumnClick,
      data: "number",
      isPadded: true,
    },
    {
      key: "EmpPhone",
      name: "Phone Number",
      fieldName: "EmpPhone",
      minWidth: 70,
      maxWidth: 90,
      isResizable: true,
      isCollapsible: true,
      data: "string",
      onColumnClick: _onColumnClick,
      onRender: (item: IEmployee) => {
        return <span>{item.EmpPhone}</span>;
      },
      isPadded: true,
    },
    {
      key: "SupervisorName",
      name: "Supervisor Name",
      fieldName: "SupervisorName",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      data: "string",
      onColumnClick: _onColumnClick,
      isPadded: true,
    },
    {
      key: "Directorate",
      name: "Directorate",
      fieldName: "Directorate",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      data: "string",
      onColumnClick: _onColumnClick,
      isPadded: true,
    },
  ];
  const [columns, setColumns] = React.useState<IColumn[]>(initialColumns);
  const [isOpen, setIsOpen] = React.useState(false);
  function _onColumnClick(ev: React.MouseEvent<HTMLElement>, column: IColumn): void {
    const newColumns: IColumn[] = [...columns];
    const currColumn: IColumn = newColumns.filter((currCol) => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
    const newItems = _copyAndSort(employees, currColumn.fieldName!, currColumn.isSortedDescending);
    // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
    const newFilteredItems = _copyAndSort(filteredEmployees, currColumn.fieldName!, currColumn.isSortedDescending);
    setEmployees(newItems);
    setFilteredEmployees(newFilteredItems);
    setColumns(newColumns);
  }

  const handleGlobalSearchChange = (newSearchValue : string): void => {

    setSearchFilter(newSearchValue.trim());
    const result = employees.filter((i) => i.EmployeeName && i.EmployeeName.toLowerCase().indexOf(newSearchValue.toLowerCase()) > -1);
    
    if (!result || result.length === 0) {
      setMessage(`No Employees found with name '${newSearchValue}'.`);
    } else {
      setMessage("");
    }
    setFilteredEmployees([...result]);
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function _onItemInvoked(item: IEmployee): void {
    setIsOpen(true);
    setSelectedEmployee(item);
    const results: IEmployee[] = employees.filter((e) => e.EmpO365Email === item.EmpO365Email || e.SupervisorO365Email === item.EmpO365Email || e.EmpO365Email === item.SupervisorO365Email);
    const orgCharEmployees = cloneDeep(results);
    const match = orgCharEmployees.find((e) => e.EmpO365Email === item.SupervisorO365Email);
    if (match) {
      match.parentId = "";
    }
    setEmployeesForOrgChart(orgCharEmployees);
  }
  function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
  }

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  function renderItem(item: IEmployee, index?: number, column?: IColumn) {
    if (column && column.fieldName) {
      switch (column?.key) {
        case "OrgChart":
          return (
            <DefaultButton
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              onClick={(e: any) => {
                _onItemInvoked(item);
              }}
              iconProps={{ iconName: "Org" }}
              className={styles.orgChartBtn}
              title="View Org Chart"
            />
          );
        case "EmployeeName":
          return <Link>{item.EmployeeName}</Link>;
        default:
          return <span>{item[column.fieldName as keyof IEmployee]}</span>;
      }
    }
  }

  return (
    <div className={styles.peopleList}>
      {employees && employees.length > 0 && (
        <div className={styles.container}>

          <TextField
            label="Filter by Employee Name:"
            value={searchFilter}
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            onChange={(_, newValue) => handleGlobalSearchChange(newValue || '')} 
            styles={{
              root: {
                margin: "0 30px 20px 0",
                maxWidth: "300px",
              }
            }}
          />

          {message && message.length > 0 && (
              <>
                <br />
                <MessageBar messageBarType={MessageBarType.info}>{message}</MessageBar>
              </>
            )
          }

          {filteredEmployees && filteredEmployees.length > 0 && (
            <DetailsList
              items={filteredEmployees}
              compact={false}
              columns={columns}
              selectionMode={SelectionMode.none}
              getKey={(item) => {
                return item.key;
              }}
              setKey="none"
              layoutMode={DetailsListLayoutMode.justified}
              isHeaderVisible={true}
              onItemInvoked={_onItemInvoked}
              onRenderItemColumn={renderItem}
            />
          )}
        </div>
      )}
      <Panel isOpen={isOpen} onDismiss={() => setIsOpen(false)} type={PanelType.medium} closeButtonAriaLabel="Close" headerText={selectedEmployee ? `${selectedEmployee.EmployeeName} Org Chart` : "Org Chart"}>
        {/*<CoreOrgChart serviceScope={props.serviceScope} selectedEmployee={selectedEmployee? selectedEmployee: undefined} data={employeesForOrgChart} showHighlighers={false} /> */}
        <CoreOrgChart serviceScope={props.serviceScope} data={employeesForOrgChart} showHighlighers={false} />
      </Panel>
    </div>
  );
};

export class PeopleListConstants {}

export interface IPeopleListState {
  columns: IColumn[];
  items: IEmployee[];
  selectionDetails: string;
  isModalSelection: boolean;
  announcedMessage?: string;
}
