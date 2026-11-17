import React, {createContext, useContext, useState} from 'react';
import { IEmployee } from "../model";

interface SharedContextType {
    selectedTab: number;
    setSelectedTab: React.Dispatch<React.SetStateAction<number>>;
    employeesForOrgChart: IEmployee[] | undefined;
    setEmployeesForOrgChart: React.Dispatch<React.SetStateAction<IEmployee[] | undefined>>;
    selectedNodeEmployee: IEmployee | undefined;
    setSelectedNodeEmployee: React.Dispatch<React.SetStateAction<IEmployee | undefined>>;
    orgChartRootNode: IEmployee | undefined;
    setOrgChartRootNode: React.Dispatch<React.SetStateAction<IEmployee | undefined>>;
    selectedNameKey: string;
    setSelectedNameKey: React.Dispatch<React.SetStateAction<string>>;
    selectedOrgChartView: string;
    setSelectedOrgChartView: React.Dispatch<React.SetStateAction<string>>;
    selectedDirectorate: string;
    setSelectedDirectorate: React.Dispatch<React.SetStateAction<string>>; 
    selectedEmployee: IEmployee | undefined;
    setSelectedEmployee: React.Dispatch<React.SetStateAction<IEmployee | undefined>>;
    isEmployeeModalOpen: boolean;
    setEmployeeModalOpen: React.Dispatch<React.SetStateAction<boolean>>; 
  }

const SharedContext = createContext<SharedContextType | null>(null);

export const SharedProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
    const [selectedTab, setSelectedTab] = useState(0);
    const [employeesForOrgChart, setEmployeesForOrgChart] = React.useState<IEmployee[] | undefined>();
    const [selectedNodeEmployee, setSelectedNodeEmployee] = React.useState<IEmployee | undefined>();
    const [orgChartRootNode, setOrgChartRootNode] = React.useState<IEmployee | undefined>();
    const [selectedNameKey, setSelectedNameKey] = useState<string>("");
    const [selectedOrgChartView, setSelectedOrgChartView] = React.useState<string>("Button1");
    const [selectedDirectorate, setSelectedDirectorate] = React.useState<string>("");
    const [selectedEmployee, setSelectedEmployee] = useState<IEmployee | undefined>();
    const [isEmployeeModalOpen, setEmployeeModalOpen] = useState(false);

    return (
      <SharedContext.Provider value={{ selectedTab, setSelectedTab, employeesForOrgChart, setEmployeesForOrgChart, 
                                      selectedNodeEmployee, setSelectedNodeEmployee, orgChartRootNode, setOrgChartRootNode,
                                      selectedNameKey, setSelectedNameKey,
                                      selectedOrgChartView, setSelectedOrgChartView, 
                                      selectedDirectorate, setSelectedDirectorate,
                                      selectedEmployee, setSelectedEmployee,
                                      isEmployeeModalOpen, setEmployeeModalOpen }}>
        {children}
      </SharedContext.Provider>
    );
  };
  
  export const useSharedContext = (): SharedContextType => {
    const context = useContext(SharedContext);
    if (!context) {
      throw new Error("useSharedContext must be used within a SharedProvider");
    }
    return context;
  };

