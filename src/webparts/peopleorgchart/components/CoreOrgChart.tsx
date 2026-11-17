/* eslint-disable @typescript-eslint/no-explicit-any */
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as d3 from "d3";
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as flextree from "d3-flextree";
import { OrgChart } from "d3-org-chart";
import { IEmployee } from "../../../model";
import { FC, useEffect, useRef, useState } from "react";
import { ServiceScope } from "@microsoft/sp-core-library";
import { IComboBoxOption} from "@fluentui/react";
import styles from "./CoreOrgChart.module.scss";
import { IconButton, Typography, Modal, Grid, Box, Autocomplete, TextField} from '@mui/material';
import PrintIcon from '@material-ui/icons/Print';
import ClearRoundedIcon from '@material-ui/icons/ClearRounded';
import FullscreenIcon from '@material-ui/icons/Fullscreen';
import ZoomInIcon from '@material-ui/icons/ZoomIn';
import ZoomOutIcon from '@material-ui/icons/zoomOut';
//import RefreshIcon from '@material-ui/icons/Refresh';
import { CalendarLtr20Regular, Mail20Regular, Chat20Regular, BuildingSkyscraper20Regular, 
  PeopleCommunity20Regular, People20Regular, Location20Regular, Call20Regular, 
  //Column20Filled
} from '@fluentui/react-icons';
import {TransformWrapper, TransformComponent} from "react-zoom-pan-pinch";
import { useSharedContext } from "../../../contexts/SharedContext";

interface ICoreOrgChartProps {
  data: IEmployee[] | undefined;
  serviceScope: ServiceScope;
  showHighlighers: boolean;
}

export const CoreOrgChart: FC<ICoreOrgChartProps> = (props) => {

  const nodeRef =  useRef<HTMLDivElement | null>(null);
  const chartRef = useRef<OrgChart<IEmployee> | null>(null);
  const [nameOptions, setNameOptions] = useState<IComboBoxOption[]>([]);
  const [dirOptions, setDirOptions] = useState<IComboBoxOption[]>([]);
  const [selectedNameOption, setSelectedNameOption] = useState<IComboBoxOption | null>(null);
  const [selectedDirKey, setSelectedDirKey] = useState<string>("");
  const [selectedEmployee, setSelectedEmployee] = useState<IEmployee | undefined>();
  const [isEmployeeModalOpen, setEmployeeModalOpen] = useState(false); 
 // const [showZoomControls, setShowZoomControls] = useState(false);
  const {employeesForOrgChart, selectedNodeEmployee, setSelectedNameKey, setEmployeesForOrgChart} = useSharedContext();

  const ResetFiltersLabel = (): JSX.Element  => {
    
    return (
      <Typography
        variant="body2"
        component="span"
        onClick={() => {
          setSelectedNameKey("");
          setSelectedNameOption(null);
          setSelectedDirKey("");
          filterChartByReset();
        }}
        className={styles.clickableLabel}
      >
        Reset
        <IconButton edge="end" size="small" sx={{color:'#2A7386', marginLeft: "5px"}} onClick={() => {
          setSelectedNameKey("");
          setSelectedNameOption(null);
          setSelectedDirKey("");
          filterChartByReset();
        }}>
          <ClearRoundedIcon style={{fontSize:'8'}} />
        </IconButton>
      </Typography>
    );
  };

 /* const handleAutoZoom = () => {
    if (chartRef.current) {
      chartRef.current.fit();  // ✅ Ensures the chart fills available space
      window.dispatchEvent(new Event("resize")); // ✅ Forces recalculation of size
    }
  }; */
  
  const parseName = (name: string): string => {
    // Split by the comma separating the last name and first/middle name
    const [lastName, firstName] = name.split(",").map(part => part.trim());
  
    // If firstName has middle names or initials, we can directly return it as the formatted string
    return `${firstName} ${lastName}`;
  };

  const formatPhoneNumber = (phone: string): string => {
    const cleaned = ('' + phone).replace(/\D/g, ''); // Remove all non-numeric characters
    const match = cleaned.match(/^(\d{3})(\d{3})(\d{4})$/);
    if (match) {
      return `(${match[1]}) ${match[2]}-${match[3]}`;
    }
    return phone; // Return unformatted phone if it doesn't match the pattern
  };
  
  const DateFormatter = (isoDate: Date): string => {
    const date = new Date(isoDate);
  
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are 0-based
    const day = String(date.getDate()).padStart(2, '0');
    const year = date.getFullYear();
  
    return `${month}-${day}-${year}`;
  };
  
  
  const handleTeamsLinkClick = (employeeEmail: string | undefined) : void => {
  
    const email = employeeEmail? encodeURIComponent(employeeEmail): '';
    const teamsWebUrl = `https://teams.microsoft.com/l/chat/0/0?users=${email}`;
    window.open(teamsWebUrl, '_blank');
  }

  const handleEmployeeNameCloseModal = (): void => {
    setSelectedEmployee(undefined);
    setEmployeeModalOpen(false);
  };
  
  function updateNodeStrokes(): void {
    // Select all .node elements inside the org chart container
    const nodeElements = document.querySelectorAll('#orgChartContainer .node');
  
    nodeElements.forEach((nodeElement) => {
      // Find the .node-content element within the .node
      const nodeContent = nodeElement.querySelector('.node-content') as HTMLElement;

      if (nodeContent) {
        // Access the highlighted status from the data-highlighted attribute of .node-content
        const isHighlighted = nodeContent.getAttribute('data-highlighted') === 'true';
  
        // Find the .node-rect element within the node
        const nodeRect = nodeElement.querySelector('.node-rect');
  
        if (nodeRect) {
          // If the node is highlighted, update the stroke
          if (isHighlighted) {
            nodeRect.setAttribute('stroke', '#2A7386');
            nodeRect.setAttribute('stroke-width', '10'); // Adjust stroke width as necessary
          } else {
            // If not highlighted, reset the stroke
            nodeRect.setAttribute('stroke', 'none');
          }
        }
      }
    });
  }

  // Function to recursively mark nodes as expanded or collapsed
  const markExpandedNodes = (node: IEmployee, employeesForOrgChart: IEmployee[] | undefined, rootNode?: IEmployee): void => {
    
     // If employeesForOrgChart is defined, expand only those nodes in the array and collapse all others
    if (employeesForOrgChart && employeesForOrgChart.some(emp => emp.id === node.id)) {
      node._expanded = true; // Expand the node if it exists in employeesForOrgChart
      if(node.EmpO365Email === selectedNodeEmployee?.EmpO365Email){
        node._highlighted = true;
      }
    } else if (!employeesForOrgChart && rootNode && node.SupervisorO365Email === rootNode.EmpO365Email) {
      // If employeesForOrgChart is undefined, expand root node and its direct reports
      node._expanded = true;
    }
    else
    {
      node._expanded = false;
    }
  };

  // Preprocess employee list to set expanded/collapsed states
  const preprocessEmployeeList = (employeeList: IEmployee[], employeesForOrgChart?: IEmployee[]): IEmployee[] => {

    let rootNode: IEmployee | undefined;

    // Create a deep copy of the employeeList to avoid mutation
    const processedData = JSON.parse(JSON.stringify(employeeList)) as IEmployee[];
    // Find the root node (node with no SupervisorO365Email or null supervisor)
    if (!employeesForOrgChart) {
      rootNode = processedData.find(node => !node.parentId) || processedData[0]; // If no SupervisorO365Email, it's the root node
    }
    // Recursively mark nodes as expanded or collapsed based on employeesForOrgChart
    processedData.forEach(node => {
      // Ensure _expanded is false for all nodes before applying any specific logic
      node._expanded = false;
      markExpandedNodes(node, employeesForOrgChart, rootNode)
   });

    return processedData;
  };

  const renderOrgChart= (orgChartData: IEmployee[]): void =>{
     // Initialize the chart if it hasn't been created yet
     if (!chartRef.current) {
      chartRef.current = new OrgChart<IEmployee>();
    }
    else{
      d3.select("#orgChartContainer").selectAll("*").remove();
      chartRef.current = new OrgChart<IEmployee>();
    }

    // Check if we have a valid node and data
    if (nodeRef.current && orgChartData) {

      const containerWidth = nodeRef.current.clientWidth || window.innerWidth;
      const containerHeight = nodeRef.current.clientHeight || window.innerHeight;

      chartRef.current
        .nodeHeight((d) => 105 + 25)
        .nodeWidth((d) => 220 + 2)
        .childrenMargin((d) => 50)
        .compactMarginBetween((d) => 35)
        .compactMarginPair((d) => 30)
        .neighbourMargin((a, b) => 20)
        .nodeContent(function (d, i, arr, state) {
      
          const imageDiffVert = 25 + 2;
          const backgroundColor = '#f3f9fa';
          const eNameFormatted = (d.data.FirstName? d.data.FirstName: '') + ' ' + (d.data.LastName? d.data.LastName: '');
          const isHighlighted = d.data._highlighted;

          return `
        <div class="node-content" data-email="${d.data.EmpO365Email?.toLowerCase()}" data-highlighted="${isHighlighted}" style='width:${222}px;height:${d.height}px;padding-top:${imageDiffVert - 2}px;padding-left:1px;padding-right:1px;border-width:0.6;border-color:#0F8C79;'>
        <div style="font-family: 'Arial', 'Helvetica', sans-serif; background-color:#bfd8db; margin-left:-1px;width:${220 - 2}px; height:${d.height - imageDiffVert}px; border-radius:10px;border: 1px solid #E4E2E9">
            <div style="display:flex;justify-content:flex-end;margin-top:5px;margin-right:8px">${d.data.Directorate}</div>
            <div style="background-color:${backgroundColor};margin-top:${-imageDiffVert - 20}px;margin-left:${15}px;border-radius:100px;width:50px;height:50px;" ></div>
            <div style="margin-top:${-imageDiffVert - 20}px;">
                <img src="https://nih.sharepoint.com/_layouts/15/userphoto.aspx?size=L&username=${d.data.EmpO365Email}" style="margin-left:${20}px;border-radius:100px;width:40px;height:40px;" />
            </div>
           
            <div style="font-size:15px;font-weight:bold;color:#08011E;margin-left:20px;margin-top:10px">  ${eNameFormatted} </div>
            <div style="color:#716E7B;margin-left:20px;margin-top:5px;font-size:13px;"> ${d.data.JobTitle} </div>
            <div><a style="color:#716E7B;margin-left:20px;margin-top:10px;font-size:13px;" href="mailto:${d.data.EmpO365Email}"> ${d.data.EmpO365Email} </a> </div>
            
            <!-- Placeholder for NavigateNext Icon -->
        
            <div id="${d.data.ItemId}" class="navigate-next-container" style="position: absolute; right: 10px; top: 60%; 
                        transform: translateY(-60%); cursor: pointer;">
            </div>
             `;
        })
        .container("#orgChartContainer")
        .data(orgChartData)
        .svgWidth(containerWidth) 
        .svgHeight(containerHeight)
        //.initialZoom(0.8)
        .onNodeClick((d) => {
          setSelectedEmployee(d.data);
          setEmployeeModalOpen(true);})
        .render();

        setTimeout(() => {
          chartRef.current?.fit(); // Ensures the OrgChart fits the container
          window.dispatchEvent(new Event("resize")); // ✅ Forces zoom reinitialization
        }, 500);

      if (nodeRef.current) {
        updateNodeStrokes();
      }
    }
  }
 
  useEffect(() => {
   
    if(props.data)
    {

      props.data
      // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
      .sort((a, b) => (a.EmployeeName! > b.EmployeeName! ? 1 : 0))
      .forEach((e) => {
        // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
        nameOptions.push({ key: e.EmpO365Email!, text: e.EmployeeName! });
      });

      if(nameOptions)
      {
        setNameOptions(nameOptions.sort((a,b) => a.text.localeCompare(b.text)));
      }

      if(selectedNodeEmployee){
        setSelectedNameKey(selectedNodeEmployee.EmpO365Email? selectedNodeEmployee?.EmpO365Email: '');
        const selected = nameOptions.find(option => option.key === (selectedNodeEmployee.EmpO365Email? selectedNodeEmployee?.EmpO365Email: '')) || null;
        setSelectedNameOption(selected);
      }

      // Use Set to find unique Directorates
      let uniqueDirectorates = Array.from(new Set(props.data.map((emp) => emp.Directorate).filter(Boolean)));
      uniqueDirectorates = uniqueDirectorates.filter((d) => d && d.length > 0);

      const directoratesToPopulate: IComboBoxOption[] = [];

      uniqueDirectorates
        .sort((a, b) => ((a? a : '') > (b? b : '') ? 1 : 0))
        .forEach((e) => {
          // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
          directoratesToPopulate.push({ key: e!, text: e! });
        });

        if(directoratesToPopulate)
        {
          setDirOptions(directoratesToPopulate.sort((a,b) => a.text.localeCompare(b.text)));
        }
    }

    if(props.data)
    {
          // Preprocess employee list to set expanded/collapsed states
        const processedEmpList = preprocessEmployeeList(props.data, employeesForOrgChart);
        renderOrgChart(processedEmpList);
    }
   
  }, [props.data]);

  useEffect(() => {
    const handleResize = (): void => {
        if (chartRef.current && nodeRef.current) {
          const newWidth = nodeRef.current.clientWidth || window.innerWidth;
          const newHeight = nodeRef.current.clientHeight || window.innerHeight;
    
          chartRef.current.svgWidth(newWidth);
          chartRef.current.svgHeight(newHeight);
          chartRef.current.fit(); // ✅ Ensures auto-resize
        }
      };
    
      window.addEventListener("resize", handleResize);
      return () => window.removeEventListener("resize", handleResize);
  }, []);

  useEffect(() => {
    const observer = new MutationObserver(() => {
      updateNodeStrokes(); //Customize node highlight color

     /* mutations.forEach((mutation) => {
        mutation.addedNodes.forEach((node) => {
          // Check if the added node is an element with the expected class name
          if (
            node.nodeType === 1 &&
            node.classList &&
            node.classList.contains("svg-chart-container")
          ) {
            // Override the inline width and height attributes
            node.setAttribute("width", "100%");
            node.setAttribute("height", "100%");
          }
        });
      });*/
      /*const svgChart = document.querySelector(".svg-chart-container");
      if (svgChart) {
        // Option 1: Use 100% to fill the container responsively
        svgChart.setAttribute("width", "100%");
        svgChart.setAttribute("height", "100%");
        
        // Option 2: Use a fixed width (e.g., 1200px) if desired:
        // svgChart.setAttribute("width", "1200px");
        // svgChart.setAttribute("height", "100%");
      } */
 
    });
  
    if (nodeRef.current) {
      observer.observe(nodeRef.current, {
        childList: true, // Watch for changes to child nodes
        subtree: true,   // Watch for changes within any descendant node
      });
    }

    // Clean up observer on unmount
    return () => {
      observer.disconnect();
    };
  }, []);

  function filterChartByName(e: any): void {
    if (chartRef.current) {
      // Get input value
      const value = e.text;
      // Clear previous higlighting
      chartRef.current.clearHighlighting();
      // Get chart nodes
      const data = chartRef.current.data();
      console.log(selectedDirKey);

      if (data) {
        // Mark all previously expanded nodes for collapse
        data.forEach((d) => (d._expanded = false));

        // Loop over data and check if input value matches any name
        data.forEach((d) => {
          // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
          if (value !== "" && d.EmployeeName!.toLowerCase() === value.toLowerCase()) {
            // If matches, mark node as highlighted
            d._highlighted = true;
            d._expanded = true;
          }
        });

        renderOrgChart(data);
      }
    }
  }
  
  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type, @typescript-eslint/no-unused-vars
 function filterChartByDir(e: any): void {
    if (chartRef.current) {

      // Get input value
      const value = e.text;
      // Clear previous higlighting
      chartRef.current.clearHighlighting();
      // Get chart nodes
      const data = chartRef.current.data();

      if (data) {
        // Mark all previously expanded nodes for collapse
        data.forEach((d) => (d._expanded = false));

        // Loop over data and check if input value matches any name
        data.forEach((d) => {
          // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
          if (value !== "" && d.Directorate!.toLowerCase() === value.toLowerCase()) {
            // If matches, mark node as highlighted
            d._highlighted = true;
            d._expanded = true;
          }
        });

        // Update data and rerender graph
        renderOrgChart(data);
      }
    }
  }

 function filterChartByReset(): void {
     // Get input value
     if (chartRef.current) {
        // Clear previous higlighting
        chartRef.current.clearHighlighting();
        setEmployeesForOrgChart(undefined);
        if(props.data)
        {
            // Preprocess employee list to set expanded/collapsed states
            const processedEmpList = preprocessEmployeeList(props.data, undefined);
            renderOrgChart(processedEmpList);
        }
     }
  }

  return (
    <div className={styles.coreOrgChart}>
     
      {props.showHighlighers === true && (
        <Box
          sx={{
            backgroundColor: "#f0f0f0", // Gray background
            padding: "15px",
            borderRadius: "8px",
            boxShadow: "0 4px 6px rgba(0, 0, 0, 0.1)", // Subtle shadow
            display: "flex",
            justifyContent: "space-between", // Aligns children horizontally
            alignItems: "flex-end", // Align all controls to bottom
            gap: "50px", // Space between elements
            height: "100%"
          }}
        >
          <Grid container spacing={3} alignItems="flex-end">
              {/* Employee ComboBox */}
              <Grid item sx={{ minWidth: 260 }}>
                <Typography variant="body1" sx={{ fontFamily: "Apparat", fontSize: "1rem !important" }}>
                  Highlight Employee:
                </Typography>
                <Autocomplete
                  sx={{
                    width: "100%",
                    maxWidth: 260, // 🔹 Increased width to 260px
                    "& .MuiInputBase-root": {
                      backgroundColor: "white",
                      borderRadius: "4px",
                      height: "32px !important", // 🔹 Reduced height
                      padding: "0px",
                      fontFamily: "Apparat",
                      fontSize: "0.75rem !important",
                      display: "flex",
                      alignItems: "center",
                    },
                    "& .MuiOutlinedInput-root": {
                      height: "32px !important", // 🔹 Ensure both combo boxes have the same height
                      display: "flex",
                      alignItems: "center",
                      "& fieldset": {
                        borderColor: "#2A7386",
                      },
                      "&:hover fieldset": {
                        borderColor: "#1E5A6F",
                      },
                      "&.Mui-focused fieldset": {
                        borderColor: "#2A7386",
                        borderWidth: "2px",
                      },
                      "& input": {
                        height: "18px !important", // 🔹 Reduced input field height
                        padding: "4px 8px !important", // 🔹 Adjusted padding for better text alignment
                        fontFamily: "Apparat",
                        fontSize: "0.75rem !important",
                      },
                    },
                    // 🔹 Fix for dropdown list font size
                    "& .MuiAutocomplete-popper": {
                      fontSize: "0.75rem !important",
                    },
                    "& .MuiAutocomplete-listbox": {
                      fontSize: "0.75rem !important",
                      padding: "4px 0",
                    },
                    "& .MuiAutocomplete-option": {
                      minHeight: "28px", // 🔹 Reduce height of dropdown options
                      fontSize: "0.75rem !important",
                    },
                  }}
                  options={nameOptions}
                  getOptionLabel={(option) => option.text}
                  value={selectedNameOption}
                  filterOptions={(options, state) =>
                    options.filter((option) => option.text.toLowerCase().includes(state.inputValue.toLowerCase()))
                  }
                  onChange={(e, value) => {
                    setSelectedNameKey(value?.key as string || "");
                    setSelectedNameOption(value);
                    setSelectedDirKey("");
                    filterChartByName(value);
                  }}
                  renderInput={(params) => (
                    <TextField
                      {...params}
                      variant="outlined"
                      InputLabelProps={{ shrink: true }}
                      sx={{
                        backgroundColor: "white",
                        borderRadius: "4px",
                        "& .MuiOutlinedInput-root": {
                          height: "32px !important",
                          padding: "0px",
                          "& input": {
                            height: "18px !important",
                            padding: "4px 8px !important",
                            fontFamily: "Apparat",
                            fontSize: "0.75rem !important",
                          },
                        },
                      }}
                    />
                  )}
                />
              </Grid>

              {/* Directorate ComboBox */}
              <Grid item sx={{ minWidth: 260 }}>
                <Typography variant="body1" sx={{ fontFamily: "Apparat", fontSize: "1rem !important" }}>
                  Highlight Directorate:
                </Typography>
                <Autocomplete
                  sx={{
                    width: "100%",
                    maxWidth: 260, // 🔹 Increased width to 260px
                    "& .MuiInputBase-root": {
                      backgroundColor: "white",
                      borderRadius: "4px",
                      height: "32px !important", // 🔹 Match height with Employee ComboBox
                      padding: "0px",
                      fontFamily: "Apparat",
                      fontSize: "0.75rem !important",
                      display: "flex",
                      alignItems: "center",
                    },
                    "& .MuiOutlinedInput-root": {
                      height: "32px !important", // 🔹 Ensure both combo boxes have the same height
                      display: "flex",
                      alignItems: "center",
                      "& fieldset": {
                        borderColor: "#2A7386",
                      },
                      "&:hover fieldset": {
                        borderColor: "#1E5A6F",
                      },
                      "&.Mui-focused fieldset": {
                        borderColor: "#2A7386",
                        borderWidth: "2px",
                      },
                      "& input": {
                        height: "18px !important",
                        padding: "4px 8px !important",
                        fontFamily: "Apparat",
                        fontSize: "0.75rem !important",
                      },
                    },
                    // 🔹 Fix for dropdown list font size
                    "& .MuiAutocomplete-popper": {
                      fontSize: "0.75rem !important",
                    },
                    "& .MuiAutocomplete-listbox": {
                      fontSize: "0.75rem !important",
                      padding: "4px 0",
                    },
                    "& .MuiAutocomplete-option": {
                      minHeight: "28px", // 🔹 Reduce height of dropdown options
                      fontSize: "0.75rem !important",
                    },
                  }}
                  options={dirOptions}
                  getOptionLabel={(option) => option.text}
                  filterOptions={(options, state) =>
                    options.filter((option) => option.text.toLowerCase().includes(state.inputValue.toLowerCase()))
                  }
                  value={selectedDirKey ? dirOptions.find((d) => d.key === selectedDirKey) : null}
                  onChange={(e, value) => {
                    setSelectedDirKey(value?.key as string || "");
                    setSelectedNameOption(null);
                    setSelectedNameKey("");
                    filterChartByDir(value);
                  }}
                  renderInput={(params) => (
                    <TextField
                      {...params}
                      variant="outlined"
                      InputLabelProps={{ shrink: true }}
                      sx={{
                        backgroundColor: "white",
                        borderRadius: "4px",
                        "& .MuiOutlinedInput-root": {
                          height: "32px !important",
                          padding: "0px",
                          "& input": {
                            height: "18px !important",
                            padding: "4px 8px !important",
                            fontFamily: "Apparat",
                            fontSize: "0.75rem !important",
                          },
                        },
                      }}
                    />
                  )}
                />
              </Grid>

              {/* Reset Filters Label */}
              <Grid item>
                <Box display="flex" alignItems="center">
                  <ResetFiltersLabel />
                </Box>
              </Grid>
          </Grid>

          {/* Right Side - Zoom & Controls */}
          <Box display="flex" alignItems="center" gap={3}> {/* 🔹 Increased gap for better spacing */}
            
            {/* Fullscreen & Print Buttons */}
            <Box display="flex" alignItems="center" gap={1}>
              {/* <IconButton size="small" sx={{ color: "#2A7386" }}>
                <FullscreenIcon />
              </IconButton>*/}
              <IconButton size="small" sx={{ color: "#2A7386" }}>
                <PrintIcon />
              </IconButton>
            </Box>
          </Box>

        </Box>
      )}
      <br />


    {/* 🔹 OrgChart Container (Zoomable Area) */}
    <Box
      sx={{
        width: "100%",
        height: "calc(100vh - 200px)", // Adjust height dynamically
        position: "relative",
        overflow: "hidden",
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        border: "1px solid #ddd", // Optional border for clarity
      }}
    >
      <TransformWrapper minScale={0.3} maxScale={3.0}>
        {({ zoomIn, zoomOut, resetTransform }) => (
          <Box
            sx={{
              width: "100%",
              height: "100%",
              position: "relative", // ✅ Ensures zoom controls stay fixed
            }}
           // onDoubleClick={handleDoubleClick} // ✅ Double-click to show zoom controls
          >
          {/*}  {showZoomControls ? ( */}
                    
                <Box
                  sx={{
                    position: "absolute",
                    top: "10px", // ✅ Fixed inside org chart, doesn't move with zoom
                    right: "10px",
                    display: "flex",
                    flexDirection: "column", // Stack buttons vertically
                    gap: "6px",
                    backgroundColor: "rgba(255, 255, 255, 0.8)",
                    padding: "6px",
                    borderRadius: "12px",
                    boxShadow: "0 4px 8px rgba(0,0,0,0.15)",
                    zIndex: 10,
                    pointerEvents: "auto"
                  }}
                >
                    <IconButton size="small" sx={{ color: "#2A7386" }} onClick={() => zoomIn()}>
                      <ZoomInIcon />
                    </IconButton>
                    <IconButton size="small" sx={{ color: "#2A7386" }} onClick={() => zoomOut()}>
                      <ZoomOutIcon />
                    </IconButton>
                    <IconButton size="small" sx={{ color: "#2A7386" }} onClick={() => resetTransform()}>
                      <FullscreenIcon />
                    </IconButton>
                </Box>
            {/*   ): null } */}

                {/* 🔹 OrgChart Wrapper - Ensures correct positioning & zoom functionality */}
                <TransformComponent>
                  <div
                    id="orgChartContainer"
                    ref={nodeRef} // ✅ OrgChart.js will insert SVG here
                    style={{
                      width: "100%",
                      height: "100%",
                      display: "flex",
                      justifyContent: "center",
                      alignItems: "center",
                      overflow: "hidden",
                      position: "relative", // Required for child elements
                    }}
                  />
                </TransformComponent>
           </Box>
        )}
      </TransformWrapper>

    </Box>

          {/* Modal for Employee Details on click of Org chart employee node*/}
          <Modal open={isEmployeeModalOpen} onClose={handleEmployeeNameCloseModal}>
                <Box
                  sx={{
                    position: 'absolute',
                    top: '50%',
                    left: '50%',
                    transform: 'translate(-50%, -50%)',
                    width: 500,
                    bgcolor: 'background.paper',
                    boxShadow: 24,
                    p: 2,
                    borderRadius: '8px',
                    display: 'flex',
                    flexDirection: 'column',
                    paddingTop: '40px'
                  }}
                >
                  {/* Employee Picture - Positioned on Top Edge */}
                  {selectedEmployee && (
                    <img
                      src={`https://nih.sharepoint.com/_layouts/15/userphoto.aspx?size=L&username=${selectedEmployee.EmpO365Email}`}
                      alt={`${selectedEmployee.EmployeeName} Photo`}
                      style={{
                        position: 'absolute',
                        top: '-50px', // Slightly less outside
                        left: '20px', // Align to the left side of the modal
                        borderRadius: '50%',
                        width: '70px', // Slightly reduced size
                        height: '70px',
                        border: '3px solid #2A7386', // Optional border for styling
                        zIndex: 1, // Ensure image doesn't overlap with other content
                      }}
                    />
                  )}

                  {/* Close Button */}
                  <IconButton
                    onClick={handleEmployeeNameCloseModal}
                    sx={{
                      position: 'absolute',
                      top: 8,
                      right: 8,
                      color: 'gray',
                      zIndex: 2, // Ensure it is above the image
                    }}
                  >
                    <ClearRoundedIcon style={{ fontSize: 18 }} />
                  </IconButton>

                  {selectedEmployee && (
                    <>
                      <Typography
                        variant="h6"
                        sx={{
                          mb: 0.5, // Reduced margin
                          display: 'flex',
                          alignItems: 'center',
                          gap: 2,
                          color: '#2A7386',
                          fontSize: '1rem !important',
                          fontWeight: 'bold'
                        }}
                      >
                       {selectedEmployee.EmployeeName? parseName(selectedEmployee.EmployeeName): ''}
                      </Typography>

                      <Typography
                        sx={{
                          mb: 2, // Reduced margin
                          display: 'flex',
                          alignItems: 'center',
                          color: '#2A7386',
                          fontSize: '0.9rem !important'
                        }}
                      >
                        {selectedEmployee.JobTitle}
                      </Typography>

                         {/* Grid Container with Reduced Spacing */}
                        <Grid container spacing={1.5}>
                          {/* Phone */}
                          <Grid item xs={6} sx={{ display: 'flex', alignItems: 'center', fontSize: '0.875rem !important' }}>
                            <Call20Regular color="#2A7386" />
                            <Typography sx={{ ml: 1 }}>
                              <strong>Phone:</strong> <br />
                              <span style={{ color: '#6264A7' }}>
                                {selectedEmployee.EmpPhone ? formatPhoneNumber(selectedEmployee.EmpPhone) : 'Not Available'}
                              </span>
                            </Typography>
                          </Grid>

                          {/* Email */}
                          <Grid item xs={6} sx={{ display: 'flex', alignItems: 'center', fontSize: '0.875rem !important' }}>
                            <Mail20Regular color="#2A7386" />
                            <Typography sx={{ ml: 1 }}>
                              <strong>Email:</strong> <br />
                              <a href={`mailto:${selectedEmployee.EmpO365Email}`} style={{ color: '#6264A7', textDecoration: 'none' }}>
                                {selectedEmployee.EmpO365Email}
                              </a>
                            </Typography>
                          </Grid>

                          {/* Microsoft Teams */}
                        <Grid item xs={6} sx={{ display: 'flex', alignItems: 'center', fontSize: '0.875rem !important' }}>
                          <Chat20Regular color="#2A7386" />
                          <Typography sx={{ ml: 1 }}>
                            <strong>Microsoft Teams:</strong> <br />
                            <a
                              href="#"
                              onClick={(e) => {
                                e.preventDefault();
                                handleTeamsLinkClick(selectedEmployee.EmpO365Email);
                              }}
                              style={{ color: '#6264A7', textDecoration: 'none' }}
                            >
                              {selectedEmployee.EmployeeName}
                            </a>
                          </Typography>
                        </Grid>

                           {/* Hire Date */}
                          <Grid item xs={6} sx={{ display: 'flex', alignItems: 'center', fontSize: '0.875rem !important' }}>
                            <CalendarLtr20Regular color="#2A7386" />
                            <Typography sx={{ ml: 1 }}>
                              <strong>Hire Date:</strong> <br />
                              {selectedEmployee.HireDate ? DateFormatter(selectedEmployee.HireDate) : 'Not Available'}
                            </Typography>
                          </Grid>

                          {/* Directorate */}
                          <Grid item xs={6} sx={{ display: 'flex', alignItems: 'center', fontSize: '0.875rem !important' }}>
                            <BuildingSkyscraper20Regular color="#2A7386" />
                            <Typography sx={{ ml: 1 }}>
                              <strong>Directorate:</strong> <br />
                              {selectedEmployee.DirectorateName || 'Not Available'}
                            </Typography>
                          </Grid>

                          {/* Sub Group */}
                          <Grid item xs={6} sx={{ display: 'flex', alignItems: 'center', fontSize: '0.875rem !important' }}>
                            <PeopleCommunity20Regular color="#2A7386" />
                            <Typography sx={{ ml: 1 }}>
                              <strong>Sub Group:</strong> <br />
                              {selectedEmployee.Department || 'Not Available'}
                            </Typography>
                          </Grid>

                          {/* Supervisor */}
                          <Grid item xs={6} sx={{ display: 'flex', alignItems: 'center', fontSize: '0.875rem !important' }}>
                            <People20Regular color="#2A7386" />
                            <Typography sx={{ ml: 1 }}>
                              <strong>Supervisor:</strong> <br />
                              {selectedEmployee.SupervisorName || 'Not Available'}
                            </Typography>
                          </Grid>

                          {/* Supervisor Email */}
                          <Grid item xs={6} sx={{ display: 'flex', alignItems: 'center', fontSize: '0.875rem !important' }}>
                            <Mail20Regular color="#2A7386" />
                            <Typography sx={{ ml: 1 }}>
                              <strong>Supervisor Email:</strong> <br />
                              <a href={`mailto:${selectedEmployee.SupervisorO365Email}`} style={{ color: '#6264A7', textDecoration: 'none' }}>
                                {selectedEmployee.SupervisorO365Email}
                              </a>
                            </Typography>
                          </Grid>

                          {/* Location */}
                          <Grid item xs={6} sx={{ display: 'flex', alignItems: 'center', fontSize: '0.875rem !important' }}>
                            <Location20Regular color="#2A7386" />
                            <Typography sx={{ ml: 1 }}>
                              <strong>Location:</strong> <br />
                              {selectedEmployee.City || 'Not Available'}
                            </Typography>
                          </Grid>
                        </Grid>
                    </>
                  )}
                </Box>
          </Modal>
   
      
    </div>
  );
};
