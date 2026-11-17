import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { spfi, SPFI, SPFx as spSPFx } from "@pnp/sp/presets/all";
import { PageContext } from "@microsoft/sp-page-context";
//import "@pnp/sp/items/get-all";
import { ICached, IEmployee } from "../model";
import { ListCacheHelper, ListTitles, LocalStorageHelper, LocalStorageKeys } from "../common";
import moment from 'moment';
import "@pnp/graph/users";
import "@pnp/graph/photos";

export interface IEmployeeService {
  getEmployees(): Promise<IEmployee[] | undefined>;
}

export class EmployeeService implements IEmployeeService {
  static readonly serviceKey: ServiceKey<IEmployeeService> = ServiceKey.create<IEmployeeService>("ncifredspfx:IEmployeeService", EmployeeService);
  siteUrl: string = '';
  sp!: SPFI;
  //graph: GraphFI;
  serviceScope!: ServiceScope;
  travelRoles!: IEmployee[];
  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this.serviceScope = serviceScope;
      const pageContext = serviceScope.consume(PageContext.serviceKey);
      this.siteUrl = pageContext.web.absoluteUrl;
      //const aadTokenProviderFactory = serviceScope.consume(AadTokenProviderFactory.serviceKey);

      //SharePoint
      this.sp = spfi().using(spSPFx({ pageContext }));

      //Graph
      //this.graph = graphfi().using(DefaultInit());
    });
  }

  async getEmployees(): Promise<IEmployee[]> {
    let cache: ICached<IEmployee[]> | undefined = LocalStorageHelper.get<ICached<IEmployee[]> | undefined>(LocalStorageKeys.Employees, this.siteUrl);

    let needRefresh = !cache || !cache.item || cache.item.length === 0;
    if (!needRefresh) {
      needRefresh = await new ListCacheHelper().DoesListValueNeedRefresh(this.siteUrl, ListTitles.Employees, cache?.cachedOnDate);
      needRefresh = true;
    }

    if (needRefresh) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      let allResposes: any[] = [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      let items: any[] = [];
      const pageSize = 5000; // Define page size based on threshold
      let skip = 0;

      do {
        items = await this.sp.web.lists
          .getByTitle(ListTitles.Employees)
          .select("Id", "Title", "field_4", "field_5", "field_6", "field_8", "field_10", "field_11", "field_13", "field_14", "field_15", "DirectorateName", "First_x0020_Name", "Last_x0020_Name", "City", "KeyStaff", "DirectorateHead")
          .items.top(pageSize) // Number of items per page
          .skip(skip)(); // Pagination

        allResposes = allResposes.concat(items);
        skip += pageSize; // Move to the next page
      } while (items.length === pageSize); // Continue if the number of items returned equals the page size

      const values: IEmployee[] = [];
      if (allResposes && allResposes.length > 0) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        allResposes.forEach((value: any) => {
          let supervisorName = "";
          let supervsorEmail = "";
          if (value.field_13 !== value.field_14) {
            supervsorEmail = value.field_14;
            const supervisor = allResposes.find((r) => r.field_13 === supervsorEmail);
            if (supervisor) {
              supervisorName = supervisor.Title;
            }
          }

          const item: IEmployee = {
            ItemId: value.Id,
            id: value.field_13,
            parentId: supervsorEmail,
            EmployeeName: value.Title,
            Directorate: value.field_4,
            JobTitle: value.field_5,
            Department: value.field_6,
            HireDate: value.field_8,
            SupervisorEmail: value.field_10,
            SupervisorName: supervisorName,
            EmpO365Email: value.field_13,
            SupervisorO365Email: supervsorEmail,
            EmpPhone: value.field_15,
            DirectorateName: value.DirectorateName,
            City: value.City,
            FirstName: value.First_x0020_Name,
            LastName: value.Last_x0020_Name,
            KeyStaff: value.KeyStaff,
            DirectorateHead: value.DirectorateHead,
            _expanded: false
          };
          values.push(item);
        });
      }

      const newCacheDate = new Date();
      if (!cache) {
        cache = {};
      }
      cache.item = values;
      cache.cachedOnDate = newCacheDate;
      LocalStorageHelper.set(LocalStorageKeys.Employees, this.siteUrl, cache);
      console.log(`REFRESHED LOOKUP VALUES for ${ListTitles.Employees} CACHE: ${moment(newCacheDate).toString()}`);
    }
    return cache?.item ? cache?.item : [];
  }
}