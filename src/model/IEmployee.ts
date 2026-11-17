export interface IEmployee {
  ItemId: number;
  id: string;
  parentId: string;
  EmployeeName?: string;
  Directorate?: string;
  JobTitle?: string;
  Department?: string;
  HireDate?: Date;
  SupervisorEmail?: string;
  SupervisorName?: string;
  EmpO365Email?: string;
  SupervisorO365Email?: string;
  EmpPhone?: string;
  DirectorateName?: string;
  City?: string;
  //Last_x0020_Name?: string;
  FirstName?: string;
  LastName?: string;
  KeyStaff?: string;
  DirectorateHead?: string;
  Tier?: number;
  children?:IEmployee[];

  // New optional properties for chart interactions
  _expanded?: boolean;
  _highlighted?: boolean;
}