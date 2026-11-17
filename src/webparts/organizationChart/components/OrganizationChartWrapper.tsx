import { FC } from "react";
import * as React from "react";
import { ServiceScope } from "@microsoft/sp-core-library";
import { Stack, DefaultButton, Text} from "@fluentui/react";
import { OrganizationChart } from "./OrganizationChart";
import { IEmployee } from "../../../model";
import { EmployeeService, IEmployeeService } from "../../../service";
//import {sampleEmployees} from "../../../service/EmployeeSampleData2";
import {DirectorateOrgChart} from "./DirectorateOrgChart";
import { useSharedContext } from "../../../contexts/SharedContext";

export interface IOrganizationChartWrapperProps {
    serviceScope: ServiceScope;
    Directorate: string | undefined;
}

export const OrganizationChartWrapper: FC<IOrganizationChartWrapperProps> = (props) => {
    const [employees, setEmployees] = React.useState<IEmployee[]>([]);
    const [loading, setLoading] = React.useState(true);
    const [error, setError] = React.useState<string | null>(null);
    const {selectedOrgChartView, setSelectedOrgChartView, setSelectedNodeEmployee, setSelectedDirectorate, 
        setEmployeeModalOpen} = useSharedContext();

    const getEmployees = async (): Promise<void> => {
        try {
            const employeeService: IEmployeeService = props.serviceScope.consume(EmployeeService.serviceKey);
            const empsResponse = await employeeService.getEmployees();
            if (empsResponse && empsResponse.length > 0) {
                setEmployees(empsResponse);
            } else {
                setError("No employees found.");
            }
        } catch (error) {
            console.error("Error getting employees: ", error);
            setError("Error loading employee data. Please try again later.");
        } finally {
            setLoading(false); // Ensure loading is set to false after the async operation
        }
    };

    React.useEffect(() => {
      
        getEmployees().catch((error) => {
            console.error("Unhandled error in getEmployees:", error);
            setError("Error loading employee data. Please try again later.");
        });
    }, []);

    if (loading) {
        return <div>Loading...</div>;
    }

    if (error) {
        return <div>{error}</div>;
    }

    return (
        <div>
                <Stack 
                    horizontal 
                    verticalAlign="center" 
                    horizontalAlign="center" // ✅ Centers the buttons & separator
                    tokens={{ childrenGap: 10 }} // Adjust gap between buttons & separator
                    style={{ width: "100%", padding: "20px 0" }} // Ensures full width and spacing
                    >

                    {/* First Labeled Button */}
                    <DefaultButton
                        text="Leadership View"
                        onClick={() => {
                                        setSelectedOrgChartView("Button1");
                                        setSelectedNodeEmployee(undefined);
                                        setSelectedDirectorate("");
                                        setEmployeeModalOpen(false);
                                }}
                        styles={{
                        root: { 
                            backgroundColor: selectedOrgChartView === "Button1" ? "#2A7386" : "transparent",
                            color: selectedOrgChartView === "Button1" ? "white" : "#2A7386",
                            border: "1px solid #2A7386",
                            borderRadius: "5px",
                            padding: "8px 16px",
                            fontSize: "14px",
                        },
                    }}
                    />

                    {/* Pipe Separator */}
                    <Text variant="medium" style={{ fontWeight: "bold", color: "#666" }}>
                        |
                    </Text>

                    {/* Second Labeled Button */}
                    <DefaultButton
                        text="Directorate View"
                        onClick={() => {
                                        setSelectedOrgChartView("Button2")
                                        setSelectedNodeEmployee(undefined);
                                        setSelectedDirectorate("");
                                }}
                        styles={{
                        root: { 
                            backgroundColor: selectedOrgChartView === "Button2" ? "#2A7386" : "transparent",
                            color: selectedOrgChartView === "Button2" ? "white" : "#2A7386",
                            border: "1px solid #2A7386",
                            borderRadius: "5px",
                            padding: "8px 16px",
                            fontSize: "14px",
                        },
                    }}
                />
                </Stack>

                {(selectedOrgChartView === "Button1" ) ? (
                    <OrganizationChart data={employees} serviceScope={props.serviceScope} />
                ) : (
                    <DirectorateOrgChart data={employees} Directorate={props.Directorate} serviceScope={props.serviceScope} />
                )}
        </div>
    );
};