/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
import { FC, useEffect, useRef, useState } from "react";
require("./DirectorateOrgChart.css"); // Import the CSS
import $ from "jquery";
// eslint-disable-next-line @typescript-eslint/no-var-requires
const orgchart = require("orgchart/src/js/jquery.orgchart.js");
import { IEmployee } from "../../../model";
import { IconButton, Typography, Grid, Box, Autocomplete, TextField, Tooltip, Snackbar} from '@mui/material';
import { IComboBoxOption} from "@fluentui/react";
import PrintIcon from '@material-ui/icons/Print';
import ClearRoundedIcon from '@material-ui/icons/ClearRounded';
import ZoomInIcon from '@material-ui/icons/ZoomIn';
import ZoomOutIcon from '@material-ui/icons/zoomOut';
import RefreshIcon  from '@material-ui/icons/Refresh';
import { CenterFocusStrong } from "@material-ui/icons";
import html2canvas from "html2canvas";
import throttle from "lodash/throttle";
import { ServiceScope } from "@microsoft/sp-core-library";
import { useSharedContext } from "../../../contexts/SharedContext";
import { EmployeeDetailsPopUp } from "./EmployeeDetailsPopUp";

interface IOrganizationChartProps {
  data: IEmployee[] | undefined;
  serviceScope: ServiceScope;
}

declare global {
  interface JQuery {
    orgchart(method: string, ...args: any[]): any;
    orgchart(options: any): JQuery;
  }
}

export const OrganizationChart: FC<IOrganizationChartProps> = (props) => {

  let $orgChartElement: any;
  const orgChartElement: any = useRef<any>(undefined);
  const [nameOptions, setNameOptions] = useState<IComboBoxOption[]>([]);
  const [dirOptions, setDirOptions] = useState<IComboBoxOption[]>([]);
  const [selectedNameOption, setSelectedNameOption] = useState<IComboBoxOption | null>(null);
  const [selectedDirKey, setSelectedDirKey] = useState<string>("");
  const [orgChartData, setOrgChartData] = useState<IEmployee | null>(null);
  const orgChartWrapperRef = useRef<HTMLDivElement | null>(null);
  const {employeesForOrgChart, selectedNodeEmployee, setSelectedNodeEmployee, setSelectedNameKey, setEmployeesForOrgChart, 
     setSelectedDirectorate, setSelectedEmployee, setEmployeeModalOpen} = useSharedContext();
 const initialZoomRef = useRef<number>(0.65);
 const zoomInitializedRef = useRef<boolean>(false);
 const zoomLevelRef = useRef<number>(initialZoomRef.current);
 const [zoomLevel, setZoomLevel] = useState<number>(initialZoomRef.current);
 const [showZoomHint, setShowZoomHint] = useState(false);

 const MIN_ZOOM = 0.3;
 const MAX_ZOOM = 2;



  const resetZoomState = () : void => {
    zoomInitializedRef.current = false;
    applyInitialZoomIfNeeded();
  };

  // Function to find an employee by ID
  const getEmployeeDataById = (nodeId: string): IEmployee | undefined => {
  return props.data?.find(emp => emp.EmpO365Email === nodeId);
  };
  
  function afterDOMUpdate(callback: () => void): void {
    requestAnimationFrame(callback);
  }

  const resetZoomAndScroll = (): void => {
    zoomInitializedRef.current = false;
    const zoom = initialZoomRef.current;
    setZoomLevel(zoom);
  
    requestAnimationFrame(() => {
      const container = document.getElementById("chart-container");
      const wrapper = document.getElementById("orgChartWrapper");
  
      if (container) {
        container.style.transform = `scale(${zoom})`;
        container.style.transformOrigin = "top left";
      }
  
      if (wrapper) {
        wrapper.scrollTo({ top: 0, left: 0, behavior: "auto" });
      }
  
      requestAnimationFrame(() => {
        resizeChartContainer();
  
        requestAnimationFrame(() => {
          applyInitialZoomIfNeeded();
        });
      });
    });
  };

  function focusOnNode(nodeElement: HTMLElement): void {
    const chartWrapper = document.getElementById("orgChartWrapper");
    const chartContainer = document.getElementById("chart-inner");
  
    if (!chartWrapper || !chartContainer || !nodeElement) return;
  
    // Get positions
    const wrapperRect = chartWrapper.getBoundingClientRect();
    const nodeRect = nodeElement.getBoundingClientRect();
  
    // Calculate new scroll positions
    const TOP_BUFFER = 150;
    const scrollX = Math.max(0, chartWrapper.scrollLeft + (nodeRect.left - wrapperRect.left) - (wrapperRect.width / 2) + (nodeRect.width / 2));
    const scrollY = Math.max(0, chartWrapper.scrollTop + (nodeRect.top - wrapperRect.top) - (wrapperRect.height / 2) + (nodeRect.height / 2)  - TOP_BUFFER);
  
    // Add horizontal padding to avoid edge clipping
    //const horizontalPadding = 200;
    // Ensure chart expands when scrolling near edges
    chartContainer.style.minWidth = `${Math.max(chartContainer.scrollWidth, scrollX + wrapperRect.width)}px`;
    chartContainer.style.minHeight = `${Math.max(chartContainer.scrollHeight, scrollY + wrapperRect.height)}px`;

    // Expand container dynamically if needed (IMPORTANT)
    chartContainer.style.width = `${Math.max(scrollX + wrapperRect.width + 200, chartContainer.scrollWidth)}px`;
    chartContainer.style.height = `${Math.max(scrollY + wrapperRect.height, chartContainer.scrollHeight)}px`;
  
    // Smooth scroll to node
    chartWrapper.scrollTo({ left: scrollX, top: scrollY, behavior: "smooth" });
  }

  function resizeChartContainer(): void {
    const container = document.getElementById("chart-container");
    const chartInner = document.getElementById("chart-inner");
    const wrapper = document.getElementById("orgChartWrapper");
    const orgChart = chartInner?.querySelector(".orgchart") as HTMLElement;
  
    if (!container || !chartInner || !wrapper || !orgChart) return;
  
    const baseWidth = orgChart.scrollWidth;
    const baseHeight = orgChart.scrollHeight;
  
    // Proportional buffer based on content width, fallback to 300
    const horizontalBuffer = Math.min(300, baseWidth * 0.1);
    const verticalBuffer = Math.min(100, baseHeight * 0.1); // add top/bottom cushion
  
    const widthNeeded = Math.max(baseWidth, wrapper.clientWidth);
    const heightNeeded = Math.max(baseHeight, wrapper.clientHeight);
  
    const finalWidth = (widthNeeded + horizontalBuffer) / zoomLevelRef.current;
    const finalHeight = (heightNeeded + verticalBuffer) / zoomLevelRef.current;

    chartInner.style.width = `${finalWidth}px`;
    chartInner.style.minWidth = `${finalWidth}px`;
    chartInner.style.height = `${finalHeight}px`;
    chartInner.style.minHeight = `${finalHeight}px`;
  
    container.style.width = `${finalWidth}px`;
    container.style.minWidth = `${finalWidth}px`;
    container.style.height = `${finalHeight}px`;
    container.style.minHeight = `${finalHeight}px`;
  }
  
  function centerOrgChart1(zoom = zoomLevel): void {
    const wrapper = document.getElementById("orgChartWrapper");
    const chartInner = document.getElementById("chart-inner");
    const orgChart = chartInner?.querySelector(".orgchart") as HTMLElement;
  
    if (!wrapper || !chartInner || !orgChart) return;
  
    const chartRect = orgChart.getBoundingClientRect();
    const wrapperRect = wrapper.getBoundingClientRect();
  
    const scrollX = chartRect.left - wrapperRect.left + (chartRect.width / 2) - (wrapperRect.width / 2);
    const scrollY = orgChart.offsetTop + (chartRect.height / 2) - (wrapperRect.height / 2);
  
    wrapper.scrollTo({
      left: Math.max(scrollX, 0),
      top: Math.max(scrollY, 0),
      behavior: "auto" // no sweep
    });
  }

  function waitAndFocusNode(emailId: string, retries = 30): void {
    setTimeout(() => {
      const target = document.querySelector(`.nodeTemplate[data-id="${emailId}"]`)?.closest('.node') as HTMLElement;
      
      if (target) {
        requestAnimationFrame(() => {
          resizeChartContainer();
          requestAnimationFrame(() => focusOnNode(target));
        });
      } else if (retries > 0) {
        setTimeout(() => waitAndFocusNode(emailId, retries - 1), 100);
      }
    }, 100); // delay to ensure node is mounted
  }
    
  function teardownOrgChart(): void {
    const $element = $(orgChartElement.current);
    const chartInner = document.getElementById("chart-inner");
    const container = document.getElementById("chart-container");

    // Remove org chart data and event handlers
    $element.off();         
    $element.removeData();  
    $element.empty();       

    // Clean accumulated size growth
    chartInner?.removeAttribute("style");

    // Also reset zoom & transform styles
    if (container) {
      container.style.transform = "";
      container.style.transformOrigin = "";
      container.style.transition = "";
      container.removeAttribute("style"); // Resets all sizing applied
    }
  }

  const applyZoomChange = (value: number, isAbsolute = false): void => {
    const wrapper = document.getElementById("orgChartWrapper");
    const container = document.getElementById("chart-container");
  
    if (!wrapper || !container) return;
  
    const oldZoom = zoomLevelRef.current;
    const newZoom = isAbsolute
    ? parseFloat(Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, value)).toFixed(2))
    : parseFloat(Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, oldZoom + value)).toFixed(2));

    if (newZoom === oldZoom) return;
  
    // Calculate current center in scroll coordinates
    const centerX = wrapper.scrollLeft + wrapper.clientWidth / 2;
    const centerY = wrapper.scrollTop + wrapper.clientHeight / 2;
  
    // Calculate percentage positions before zoom
    const relX = centerX / (container.offsetWidth * oldZoom);
    const relY = centerY / (container.offsetHeight * oldZoom);
  
    setZoomLevel(newZoom);
    zoomLevelRef.current = newZoom;
  
    // Apply zoom with transform
    requestAnimationFrame(() => {
    container.style.transform = `scale(${newZoom})`;
    container.style.transformOrigin = "top left";
    container.style.transition = "transform 0.2s ease";

    resizeChartContainer(); // always after transform is applied

    requestAnimationFrame(() => {
      const newScrollLeft = (container.offsetWidth * newZoom * relX) - (wrapper.clientWidth / 2);
      const newScrollTop = (container.offsetHeight * newZoom * relY) - (wrapper.clientHeight / 2);

      wrapper.scrollTo({ left: newScrollLeft, top: newScrollTop, behavior: "auto" });
    });
  });
  };

  function applyInitialZoomIfNeeded(): void {
    if (zoomInitializedRef.current) return;
  
    zoomInitializedRef.current = true;
    const zoom = initialZoomRef.current;
    setZoomLevel(zoom);
  
    requestAnimationFrame(() => {
      const container = document.getElementById("chart-container");
      if (!container) return;
  
      container.style.transform = `scale(${zoom})`;
      container.style.transformOrigin = "top left";
  
      requestAnimationFrame(() => {
        resizeChartContainer();
  
        requestAnimationFrame(() => {
          centerOrgChart1(zoom); // Only after resize is visually applied
        });
      });
      console.log("applyInitialZoomIfNeeded called. Zoom:", zoomInitializedRef.current ? "already applied" : "applying...");

    });
  }
  
  function updateNodeStrokes(): void {
    const nodeElements = document.querySelectorAll('#chart-inner .node');
    let highlightedNode: HTMLElement | null = null;
  
    nodeElements.forEach((nodeElement) => {
      const nodeTemplate = nodeElement.querySelector('.nodeTemplate') as HTMLElement;
  
      if (nodeTemplate) {
        const isHighlighted = nodeTemplate.getAttribute('data-highlighted') === 'true';
        const isExpanded = nodeTemplate.getAttribute('data-expanded') === 'true';
  
        if (isHighlighted) {
          nodeTemplate.classList.add('highlight-node');
          highlightedNode = nodeElement as HTMLElement;
        } else {
          nodeTemplate.classList.remove('highlight-node');
        }
  
        if (isExpanded) {
          nodeTemplate.classList.add('expand-node');
          nodeTemplate.setAttribute("data-expanded", "true");
        } else {
          nodeTemplate.classList.remove('expand-node');
          nodeTemplate.setAttribute("data-expanded", "false");
        }
      }
    });
  
    if (highlightedNode) {
      // 1. Resize first, wait for layout to stabilize
      requestAnimationFrame(() => {
        resizeChartContainer();
    
        // 2. Wait one more frame to ensure resize takes effect
        requestAnimationFrame(() => {
          if (highlightedNode) {
          focusOnNode(highlightedNode);
          }
        });
      });
    }
  }

  const updateEmployeeHighlighting = (node: IEmployee, emailId: string): { updatedNode: IEmployee; matchedNode?: IEmployee } => {
    // Check if the current node matches the email ID
    const isMatch = node.EmpO365Email?.toLowerCase() === emailId;

    // Recursively update children first
    const updatedChildrenResults = node.children?.map((child) => updateEmployeeHighlighting(child, emailId)) || [];

    // Extract updated children and find the matched node if any
    const updatedChildren = updatedChildrenResults.map(result => result.updatedNode);
    const matchedNode = isMatch ? { ...node, children: [] } : updatedChildrenResults.find(result => result.matchedNode)?.matchedNode;

    // Check if any child is highlighted (to ensure parent expansion)
    const isChildHighlighted = updatedChildrenResults.some(result => result.matchedNode || result.updatedNode._highlighted);

    return {
      updatedNode: {
        ...node,
        _highlighted: isMatch,
        _expanded: isMatch || isChildHighlighted || node._expanded, // Expand node if matched or a child is highlighted
        children: updatedChildren, // Update children recursively
      },
      matchedNode,
    };
  };
  
  // Function to recursively mark nodes as expanded or collapsed
  const markExpandedNodes = (node: IEmployee, employeesForOrgChart: IEmployee[] | undefined, rootNode?: IEmployee): void => {
    // If employeesForOrgChart is defined, expand only those nodes in the array and collapse all others
    if (employeesForOrgChart && employeesForOrgChart.some(emp => emp.id === node.id)) {
      node._expanded = true;
    } else if (!employeesForOrgChart && rootNode && node.EmpO365Email === rootNode.EmpO365Email) {
      // If employeesForOrgChart is undefined, expand root node and its direct reports
      node._expanded = true;
    }
    else
    {
      node._expanded = false;
    }
  };

  /**
 * Preprocesses the given list of employees, 
 *      1. Filtering nodes to be displayed on KeyStaff and Directorate Head Flags
 *      2. Find and mark root node and assign Tier 1
 *      3. Preprocess employee list to set expanded/collapsed states
 * @param employeeList - Base Employee list to be processed for Organization Chart display.
 * @param employeesForOrgChart - Employee list selected from Employee Search Tab to be selected
 * 
 * @returns A new list of employees after preprocessing. 
 */
  const preprocessEmployeeList = (employeeList: IEmployee[], employeesForOrgChart?: IEmployee[]): IEmployee[] => {

    //let rootNode: IEmployee | undefined;

    // Create a deep copy of the employeeList to avoid mutation
    let processedData = JSON.parse(JSON.stringify(employeeList)) as IEmployee[];

    //Filtered Employee List for Org Chart based on DirectorateHead and KeyStaff flags
    if(processedData && processedData.length > 0){
      processedData = processedData.filter(node => (node.KeyStaff && (node.KeyStaff?.trim().toLowerCase() === "yes")) ||
                                                    (node.DirectorateHead && (node.DirectorateHead?.trim().toLowerCase() === "yes")));
    }

    //Find the root node and assign Tier
    const rootNode = processedData.find(node => !node.parentId) || processedData[0];
    if(rootNode)
    {
      const matchedRootNode = processedData.find(node => node.EmpO365Email?.trim().toLowerCase() === rootNode?.EmpO365Email?.trim().toLowerCase());

      if(matchedRootNode)
        matchedRootNode.Tier = 1;
    }

    // Recursively mark nodes as expanded or collapsed based on employeesForOrgChart
    processedData.forEach(node => {
      // Ensure _expanded is false for all nodes before applying any specific logic
      node._expanded = false;
      markExpandedNodes(node, employeesForOrgChart, rootNode)
   });

    return processedData;
  };

  /**
* Function to build hierarchy from flat employee list
* @param employeeList - Processed employee list to be displayed as Organization Chart 
* 
* @returns The root node with the full hierarchy of employees nested within it. 
*          If no valid hierarchy can be built, the function may return `undefined`.
*/
 const buildHierarchy = (employees: IEmployee[]): IEmployee | undefined => {
  if (!employees || employees.length === 0) return undefined;

  const employeeMap = new Map<string, IEmployee>();

  // Unique key = employee record + directorate
  const getNodeKey = (emp: IEmployee): string => {
    const directorate =
      emp.Directorate?.trim().toLowerCase() || "no-directorate";

    return `${emp.id}__${directorate}`;
  };

  // Create separate nodes for every employee/directorate assignment
  employees.forEach(emp => {
    const nodeKey = getNodeKey(emp);

    employeeMap.set(nodeKey, {
      ...emp,
      children: []
    });
  });

  let root: IEmployee | undefined;

  employees.forEach(emp => {
    const childKey = getNodeKey(emp);
    const child = employeeMap.get(childKey);

    if (!child) return;

    // Root/top leader
    if (!emp.parentId) {
      if (!root) {
        root = child;
      }
      return;
    }

    // A parent ID may now have multiple directorate assignments
    const possibleParents = employees.filter(
      parent => parent.id === emp.parentId
    );

    if (possibleParents.length === 0) {
      return;
    }

    // First try to connect the child to the parent's
    // matching directorate assignment
    const sameDirectorateParent = possibleParents.find(
      parent =>
        (parent.Directorate?.trim().toLowerCase() || "") ===
        (emp.Directorate?.trim().toLowerCase() || "")
    );

    // If there is no directorate match,
    // preserve the existing hierarchy relationship
    const parentRecord =
      sameDirectorateParent ?? possibleParents[0];

    const parentKey = getNodeKey(parentRecord);
    const parent = employeeMap.get(parentKey);

    if (parent) {
      parent.children = parent.children ?? [];
      parent.children.push(child);
    }
  });

  return root;
};

   /**
  * Second Pass: Assign tiers to each node based on their roles and parent-child relationships
  * @param employeeList - Sort tiered hierarchical employee node
  * 
  * @returns - Processed sorted tiered employee node as per KeyStaff, DirectorateHead and Employee Name
  */
   const sortedAssignedTiers = (tieredEmployee: IEmployee): IEmployee => {
    if (tieredEmployee?.children?.length) {
      const keyStaff = tieredEmployee.children
        .filter(emp => emp.KeyStaff?.trim().toLowerCase() === 'yes')
        .sort((a, b) => (a.FirstName ?? '').localeCompare(b.FirstName ?? ''));
  
      const directorateHeads = tieredEmployee.children
        .filter(emp => emp.KeyStaff?.trim().toLowerCase() !== 'yes' && emp.DirectorateHead?.trim().toLowerCase() === 'yes')
        .sort((a, b) => (a.FirstName ?? '').localeCompare(b.FirstName ?? ''));
  
      // Replace children with sorted list
      tieredEmployee.children = [...keyStaff, ...directorateHeads];
    }
    return tieredEmployee;
  };  
  
    /**
  * Second Pass: Assign tiers to each node based on their roles and parent-child relationships
  * @param employeeList - Processed hierarchical employee node
  * 
  * @returns - Processed hierarchical employee node with tiers assigned for tiered organization chart display 
  */
  const assignTiers = (node: IEmployee | undefined, parentTier: number = 0):  IEmployee | undefined => {
    if (!node) return;

    // Assign the tier for the current node based on its role or its parent tier
    //if (node.parentId === ''){
      //node.Tier = 1;
    //}
    //else 
    if (parentTier === 1 && node.KeyStaff) {
        node.Tier = 2;
    }
    else if(parentTier === 1 && node.DirectorateHead){
        node.Tier = 3;
    }
    else if(parentTier === 2 && node.DirectorateHead){
      node.Tier = 4;
    }

    // Recursively assign tiers to children
    node.children?.forEach(child => assignTiers(child, node.Tier));

    return node;
  };

  const handleRenderComplete = (): void => {
    requestAnimationFrame(() => {
      resizeChartContainer();        // Recalculate layout without transform hacks
      requestAnimationFrame(() => {
        //centerOrgChart1();
        applyInitialZoomIfNeeded(); 
      }); 
    });
  };

  function syncAllToggleIcons(): void {
    $('#chart-inner .custom-bottom-toggle').each(function () {
      const $toggle = $(this);
      const $li = $toggle.closest('li.hierarchy');
      const isCollapsed = $li.hasClass('isChildrenCollapsed');
  
      $toggle.find('.icon').text(isCollapsed ? 'add_circle' : 'remove_circle');
  
      // Optional: update expand-node class
      const $template = $li.find('.nodeTemplate');
      $template
        .attr('data-expanded', isCollapsed ? 'false' : 'true')
        .toggleClass('expand-node', !isCollapsed);
    });
  }

  /**
  * Render Organization Chart from the tiered hierarchical employee list 
  * @param employeeList - Processed tiered hierarchical employee list node
  */
  const renderOrgChart = (organizationChartData: IEmployee, matchedNode: IEmployee | undefined): void => {
    const $element = $(orgChartElement.current);

    //Full teardown before creating a new chart
    teardownOrgChart(); // Unmounts React icons
    $element.off();     // Detach all event handlers (important!)
    $element.removeData(); // Remove plugin-related jQuery data
    $element.empty();   // Clear any remaining DOM
  
    // Store the reference (optional if needed globally)
    $orgChartElement = $element;

    console.log(orgchart);
  
    const getMatchedNodeTier = (node: IEmployee | undefined): number => {
      if (!node) return 2;
      return node.Tier && node.Tier >= 3 ? (node.Tier - 1) : 2;
    };

    $orgChartElement.orgchart({
        data: organizationChartData ? organizationChartData : {},
        nodeContent: "id",
        nodeID: "id",
        zoom: false,
        pan: true,
        initCollapsed: true,
        visibleLevel: getMatchedNodeTier(matchedNode),
        toggleSiblingsResp: true,
        verticalDepth: (data: any) => (data.Tier === 2 ? 0 : 3), // Force Tier 4 to stack vertically
        siblingSeparation: (data: any) => (data.Tier === 2 ? 30 : 50), // Reduce horizontal spacing
        subtreeSeparation: (data: any) => (data.Tier === 2 ? 40 : 80), // Adjust spacing between levels
        levelSeparation: (data: any) => (data.Tier === 2 ? 40 : 70), // Closer vertical spacing for better stacking
        direction: "t2b",

        linkTemplate: function (linkData: { from: { x: number, y: number }, to: { x: number, y: number } }) {
            const { from, to } = linkData;
            const midX = (from.x + to.x) / 2; // Midpoint for a smooth curve
            const controlY = from.y + 25;

            return `
                <path d="M${from.x},${from.y} 
                         C${midX},${controlY} 
                         ${midX},${controlY} 
                         ${to.x},${to.y}" 
                      stroke="black" 
                      fill="none" 
                      stroke-width="2"/>
            `;
        },
        nodeTemplate: (data: any) => {
            const eNameFormatted = (data.FirstName ? data.FirstName : '') + ' ' + (data.LastName ? data.LastName : '');
            const nodeId = data.EmpO365Email;

            const nodeLevel = data.levelOffset && data.levelOffset > 0 ? data.level + data.levelOffset : data.level;
            const acronym = data.KeyStaff ? "KEY" : data.DirectorateHead ? "SLT" : "";

           /* const jobTitle = data.JobTitle || "";
            const isLongJobTitle = jobTitle.length > 25; // adjust threshold as needed */

            const inlineStyle = 'min-height: 130px;'; // add room only if needed*/

            let nodeStyle;

            switch (nodeLevel) {
                case 1:
                    nodeStyle = "nodeTemplate nodeTemplate1";
                    break;
                case 2:
                    nodeStyle = "nodeTemplate nodeTemplate2";
                    break;
                case 3:
                    nodeStyle = "nodeTemplate nodeTemplate3";
                    break;
                case 4:
                    nodeStyle = "nodeTemplate nodeTemplate4";
                    break;
                case 5:
                    nodeStyle = "nodeTemplate nodeTemplate5";
                    break;
                case 6:
                    nodeStyle = "nodeTemplate nodeTemplate6";
                    break;
                case 7:
                    nodeStyle = "nodeTemplate nodeTemplate7";
                    break;
                default:
                    nodeStyle = "nodeTemplate";
                    break;
            }

            return `
              <div class="${nodeStyle}" style="${inlineStyle}"
                  data-tier="${data.Tier}" data-id="${nodeId}"
                  data-highlighted="${data._highlighted}" data-expanded="${data._expanded}">
                <div class="profilePic">
                   <img src="https://nih.sharepoint.com/_layouts/15/userphoto.aspx?size=L&username=${data.EmpO365Email}" alt="Profile Picture of ${eNameFormatted}" />
                </div>
                <div class="info">
                  <div class="personName">${eNameFormatted}</div>
                  <div class="personTitle">${data.Directorate}</div>
                  <div class="personTitle">${data.JobTitle}</div>
                </div>
                ${acronym ? `<div class="tierAcronym" aria-label="${acronym}">${acronym}</div>` : ""}
              </div>
            `;
        },

        createNode: function (node: any, data: any) {
            const mutiplyBy: number = 40;
            
            if (data.Tier && (data.Tier === 1 || data.Tier === 2 || data.Tier === 3 || data.Tier === 4)) {
                node.css({
                    "margin-top": data.Tier * mutiplyBy + "px",
                    "--top": -11 - data.Tier * mutiplyBy + "px",
                    "--height": 9 + data.Tier * mutiplyBy + "px",
                    "--top-cross-point": -13 - data.Tier * mutiplyBy + "px",
                    "--height-cross-point": 11 + data.Tier * mutiplyBy + "px",
                });

                node.attr("data-isExpanded", data._expanded);
            }

            const childrenCount = data.children?.length || 0;
            if (childrenCount === 0) return;
      
            // Create the toggle element with default icon (will be updated)
            const $toggle = $(`
              <div class="custom-bottom-toggle" title="Expand/Collapse team">
                <span class="material-icons icon">add_circle</span>
                <span class="child-count">${childrenCount}</span>
              </div>
            `);
      
            const $template = node.find('.nodeTemplate');
            if ($template.length) {
              $toggle.insertAfter($template);
            }
    
            // Attach click handler
            $toggle.on('click', () => {
              const $chevron = node.find('.bottomEdge');
              const $hierarchy = node.closest('li.hierarchy');
          
              if ($chevron.length) {
                $chevron.trigger('click');
          
                afterDOMUpdate(() => {
                  updateAllDescendantIcons($hierarchy);
                
                  requestAnimationFrame(() => {
                    resizeChartContainer(); // 1st pass
                
                    requestAnimationFrame(() => {
                      resizeChartContainer(); // 2nd pass after reflow
                      focusOnNode(node.get(0)); // then scroll
                    });
                  });
                });                
              }
            });

            const $sideEdges = node.find('.leftEdge, .rightEdge');
        if ($sideEdges.length) {
          $sideEdges.on("click", function (e: JQuery.ClickEvent) {
            const $thisEdge = $(e.currentTarget);
            const isLeft = $thisEdge.hasClass("leftEdge");
          
            const $currentHierarchy = node.closest("li.hierarchy");
          
            afterDOMUpdate(() => {
              requestAnimationFrame(() => {
                console.log("Edge clicked:", isLeft ? "left" : "right");
                updateAdjacentSiblingIcons($currentHierarchy, isLeft ? "left" : "right");
              });
            });
            
          });                        
        }
           
            // Initial icon setup
            afterDOMUpdate(() => {
              const $hierarchy = node.closest('li.hierarchy');
              const isCollapsed = $hierarchy.hasClass('isChildrenCollapsed');
              $toggle.find('.icon').text(isCollapsed ? 'add_circle' : 'remove_circle');
            });

           afterDOMUpdate(() => {
              const $template = node.find('.nodeTemplate');
              const $jobTitle = $template.find('.personTitle').last();
          
              if ($template.length && $jobTitle.length) {
                const jobTitleEl = $jobTitle.get(0);
                const didWrap = jobTitleEl?.scrollHeight > jobTitleEl?.offsetHeight;
                const baseHeight = $template.outerHeight();

                if (didWrap && typeof baseHeight === "number") {
                  const paddedHeight = baseHeight + 20;
                  $template.css('height', `${paddedHeight}px`);
                } 
                /*else {
                  $template.css('height', 'auto');
                }*/
              }
            });
          
            // Update all nested icons recursively
            function updateAllDescendantIcons($root: JQuery<HTMLElement>): void {
              $root.find('.custom-bottom-toggle').each(function () {
                const $thisToggle = $(this);
                const $li = $thisToggle.closest('li.hierarchy');
                const isCollapsed = $li.hasClass('isChildrenCollapsed');
                $thisToggle.find('.icon').text(isCollapsed ? 'add_circle' : 'remove_circle');
              });
            }

            function updateAdjacentSiblingIcons($current: JQuery<HTMLElement>, direction: "left" | "right"): void {
              const $siblings = $current.siblings("li.hierarchy");
            
              $siblings.each(function () {
                const $sibling = $(this);
                const siblingOffset = $sibling.offset()?.left ?? 0;
                const currentOffset = $current.offset()?.left ?? 0;
            
                const isTarget =
                  (direction === "left" && siblingOffset < currentOffset) ||
                  (direction === "right" && siblingOffset > currentOffset);
            
                if (!isTarget) return;
    
                console.log("Direction:", direction);
                console.log("Current offset:", currentOffset, "Sibling offset:", siblingOffset);
            
                const isCollapsed = $sibling.hasClass("isChildrenCollapsed");
                const $toggle = $sibling.find(".custom-bottom-toggle");
            
                if ($toggle.length) {
                  console.log("Updating toggle icon for sibling:", $sibling);
                  $toggle.find(".icon").text(isCollapsed ? "add_circle" : "remove_circle");
                }
            
                $sibling.find(".nodeTemplate")
                  .attr("data-expanded", isCollapsed ? "false" : "true")
                  .toggleClass("expand-node", !isCollapsed);
              });
            }
        },
        subTreeConfig: {
            level_2: { // Applies only to Tier 2
                layout: orgchart.tree // Forces vertical stacking
            }
        }
    });
    
    $orgChartElement.on('click', '.leftEdge, .rightEdge, .bottomEdge', () => {
      requestAnimationFrame(() => {
        syncAllToggleIcons();
      });
    });

    // Attach click event dynamically to `.nodeTemplate` divs
    $orgChartElement.on("click", ".nodeTemplate", function (this: HTMLElement, event: JQuery.ClickEvent) {

      event.stopPropagation();
      const nodeId = $(this).attr("data-id");
      if (nodeId) {
        const selectedEmployee = getEmployeeDataById(nodeId);
        if (selectedEmployee) {
          setSelectedEmployee(selectedEmployee);
          setEmployeeModalOpen(true);
        }
      }
    });

     // Render complete + scroll focus
    requestAnimationFrame(() => {
      handleRenderComplete(); // resize & center chart

      // Focus on matched employee node (after layout settles)
      if (matchedNode?.EmpO365Email) {
        requestAnimationFrame(() => {
          waitAndFocusNode(matchedNode.EmpO365Email ?? '');
        });
      }
    });

    if (!$orgChartElement.data("touchEventSet")) {
      $orgChartElement.on("touchmove", (event: JQuery.Event) => {
          event.preventDefault();
      });
      $orgChartElement.data("touchEventSet", true);
    }
  };

  const renderDirectorateEmployeeOrgChart = (highlightNodeOverride?: IEmployee): void => {
    if (!props.data) return;
    if (!orgChartElement.current) return;
  
    teardownOrgChart();
  
    const processedEmpList = preprocessEmployeeList(props.data, employeesForOrgChart);
    if (processedEmpList.length === 0) return;
  
    const hierarchy = buildHierarchy(processedEmpList);
    if (!hierarchy) return;
  
    const tierAssignedHierarchy = assignTiers(hierarchy);
    if (!tierAssignedHierarchy) return;
  
    let chartData = tierAssignedHierarchy;
    let highlightNode = highlightNodeOverride;
  
    // If nothing passed, fallback to context
    const emailToHighlight = highlightNodeOverride?.EmpO365Email || selectedNodeEmployee?.EmpO365Email;
  
    if (emailToHighlight) {
      const { updatedNode, matchedNode } = updateEmployeeHighlighting(tierAssignedHierarchy, emailToHighlight);
      chartData = updatedNode;
      highlightNode = matchedNode;
    }
  
    setOrgChartData(chartData);
    //applyInitialZoomEarly();
    const sortedTieredHierarchy = sortedAssignedTiers(chartData);
    renderOrgChart(sortedTieredHierarchy, highlightNode);
    updateNodeStrokes();

   /* requestAnimationFrame(() => {
      if (highlightNode?.EmpO365Email) {
        waitAndFocusNode(highlightNode.EmpO365Email);
      }
    });*/
    
  };
  

/****************************************************************************************************************************** */
 
  const exportOrgChartToPDF = async (): Promise<void> => {
    const outer = document.getElementById("orgChartOuter");
    if (!outer) return;
  
    // Hide toolbar buttons
    const controls = document.querySelectorAll(".MuiIconButton-root");
    controls.forEach(btn => {(btn as HTMLElement).style.display = "none"});
  
    // Temporarily disable node shadows using a global class
    document.body.classList.add("exporting-chart");
  
    // Replace profile pictures with gray SVG placeholders to avoid black blobs
    const profileImgs = Array.from(outer.querySelectorAll(".nodeTemplate img"));
    const originalSrcs: string[] = [];
  
    profileImgs.forEach((img, i) => {
      const el = img as HTMLImageElement;
      originalSrcs[i] = el.src;
  
      // Force visibility and reset style overrides that html2canvas may respect
      el.style.visibility = "visible";
      el.style.background = "none";
      el.style.boxShadow = "none";
      el.style.filter = "none";
  
      el.src =
        "data:image/svg+xml;base64," +
        btoa(
          `<svg xmlns='http://www.w3.org/2000/svg' width='60' height='60'>
            <circle cx='30' cy='30' r='28' fill='#ccc' />
          </svg>`
        );
    });
  
    try {
      // Wait for all modified images to load
      await Promise.all(
        profileImgs.map((img) => {
          const el = img as HTMLImageElement;
          return el.complete
            ? Promise.resolve()
            : new Promise((resolve) => {
                el.onload = resolve;
                el.onerror = resolve;
              });
        })
      );      
  
      await new Promise((resolve) => {
        setTimeout(resolve, 100);
      });
  
      const canvas = await html2canvas(outer, {
        backgroundColor: "#fff",
        useCORS: true,
        scale: 2,
      });
  
      const imgData = canvas.toDataURL("image/png");
      const printWindow = window.open("", "_blank");
      if (!printWindow) return;
  
      printWindow.document.write(`
        <html>
          <head>
            <title>Leadership View - ${selectedDirKey || selectedNameOption?.text || ""}</title>
            <style>
              body { margin: 0; padding: 0; text-align: center; background: #fff; }
              img { max-width: 100%; height: auto; margin-top: 40px; }
            </style>
          </head>
          <body>
            <img src="${imgData}" />
            <script>window.onload = () => window.print();</script>
          </body>
        </html>
      `);
  
      printWindow.document.close();
    } catch (error) {
      console.error("Print error:", error);
    } finally {
      // Restore images
      profileImgs.forEach((img, i) => {
        const el = img as HTMLImageElement;
        el.src = originalSrcs[i];
        el.style.visibility = "";
        el.style.background = "";
        el.style.boxShadow = "";
        el.style.filter = "";
      });
  
      document.body.classList.remove("exporting-chart");
      controls.forEach(btn => {(btn as HTMLElement).style.display = ""});
    }
  };

  const zoomOut = (): void => {
    applyZoomChange(-0.1);
  };

  const zoomIn = (): void => {
    applyZoomChange(0.1);
  };

  /**
   * Clears highlighting for selected organization chart nodes
  */
  const clearHighlighting = (): void => {
    // Select all highlighted nodes and remove the highlight class
    document.querySelectorAll(".highlight-node").forEach((node) => {
      node.classList.remove("highlight-node");
    });
  
    // Reset `data-highlighted` attributes if used
    document.querySelectorAll(".node-content[data-highlighted='true']").forEach((node) => {
      node.setAttribute("data-highlighted", "false");
    });
  
    updateNodeStrokes(); // Ensure stroke resets
  };

  /**
  * Clears previous selection highlighting and resets Organization Chart in the initial view
  */
  function filterChartByReset(): void {
   
    if (orgChartElement.current && orgChartData) {
   
       clearHighlighting(); // Clear old highlights first
       setEmployeesForOrgChart(undefined);
       setSelectedNodeEmployee(undefined);
       setSelectedNameKey('');
       setSelectedDirectorate('');
       setSelectedDirKey('');

      if (orgChartElement.current) {
        teardownOrgChart();
        $(orgChartElement.current).empty();
      }

      resetZoomAndScroll();
      renderDirectorateEmployeeOrgChart();
    }
  }

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
        className="clickableLabel"
      >
        Reset
        <IconButton edge="end" size="small" aria-label="Reset" sx={{color:'#2A7386', marginLeft: "4px", padding: "2px"}} onClick={() => {
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

  /**
  * Selects and highlights the selected employee node
  * @param value - Selected Employee ComboBox option
  */
  const filterChartByName = (value: IComboBoxOption | null): void => {
    if (!value || value.key === undefined || value.key === null) return;

    const emailId = value.key.toString().toLowerCase();
    setSelectedDirKey('');

    const wrapper = document.getElementById("orgChartWrapper");
    if (wrapper) {
      wrapper.scrollTo({ top: 0, left: 0, behavior: "smooth" });
    }

    if(orgChartData){

      const { updatedNode: updatedOrgChartData, matchedNode } = updateEmployeeHighlighting(orgChartData, emailId);

      resetZoomState();

      const sortedupdatedOrgChartData = sortedAssignedTiers(updatedOrgChartData);
      renderOrgChart(sortedupdatedOrgChartData, matchedNode);
      updateNodeStrokes();

      // Wait and focus after DOM is updated
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          waitAndFocusNode(emailId);
        });
      })
    }
  };

   /**
  * Populate "Highlight Employee" Dropdown with employee list
  * @param employeeList - Employee list for Organization Chart 
  */
   const populateEmployeeDropdown = (employees: IEmployee[]): void => {
    const options: IComboBoxOption[] = employees
      .filter(e =>
        !!e.EmployeeName &&
        !!e.EmpO365Email &&
        ((e.KeyStaff?.trim().toLowerCase() === "yes") ||
         (e.DirectorateHead?.trim().toLowerCase() === "yes"))
      )
      .sort((a, b) => {
        const nameA = a.EmployeeName ?? "";
        const nameB = b.EmployeeName ?? "";
        return nameA.localeCompare(nameB);
      })
      .map(e => ({
        key: e.EmpO365Email ?? "",
        text: e.EmployeeName ?? ""
      }));
  
    setNameOptions(options);
  };
  
  
  /**
  * Adds highlighting for the selected directorate and clears existing highlighting
  * @param node - Hiearchical employee node with children tiers
  * @param Directorate - Selected directorate to be highlighted
  * @returns - Returns hierarchical employee node with updated highlighting
  */
  const updateDirectorateHighlighting = (node: IEmployee, directorateKey: string): { updatedNode: IEmployee; matchedNode?: IEmployee } => {
    const isMatch = node.Directorate?.trim().toLowerCase() === directorateKey.trim().toLowerCase();

    const childrenResults = node.children?.map(child => updateDirectorateHighlighting(child, directorateKey)) || [];
    const updatedChildren = childrenResults.map(result => result.updatedNode);
    const matchedNode = isMatch ? { ...node, children: [] } : childrenResults.find(result => result.matchedNode)?.matchedNode;
    const isChildHighlighted = childrenResults.some(result => result.matchedNode || result.updatedNode._highlighted);

    return {
        updatedNode: { ...node, 
                    _highlighted: isMatch, _expanded: isMatch || isChildHighlighted, 
                    children: updatedChildren },
        matchedNode,
    };
  };

  /**
  * Selects and highlights nodes with selected directorate
  * @param value - Selected Directorate ComboBox option
  */
  function filterChartByDir(value: IComboBoxOption | null): void {

    if (!value || value.key === undefined || value.key === null) return;

    const directorate = value.key.toString().toLowerCase();
    setSelectedNameOption(null);

    const wrapper = document.getElementById("orgChartWrapper");
    if (wrapper) {
      wrapper.scrollTo({ top: 0, left: 0, behavior: "auto" });
    }

    if(orgChartData){
      const { updatedNode: updatedOrgChartData, matchedNode }  = updateDirectorateHighlighting(orgChartData, directorate);

      resetZoomState();

      const sortedupdatedOrgChartData = sortedAssignedTiers(updatedOrgChartData);
      renderOrgChart(sortedupdatedOrgChartData, matchedNode);
      updateNodeStrokes();

      // Wait and focus after DOM is updated
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          waitAndFocusNode(directorate);
        });
      })
    }
  }

  /**
  * Populate "Highlight Directorate" Dropdown distict directorates from employee list
  * @param employeeList - Employee list for Organization Chart 
  */
  const populateDirectorateDropDown = (employees: IEmployee[]): void => {
    const uniqueDirectoratesMap = new Map<string, string>();

    employees.forEach(emp => {
      if (emp.Directorate && emp.DirectorateName && emp.DirectorateName.length > 0) {
        uniqueDirectoratesMap.set(emp.Directorate, emp.DirectorateName);
      }
    });
    
    const uniqueDirectorates = Array.from(uniqueDirectoratesMap.entries()).map(
      ([key, text]) => ({ key, text })
    );
    
    const excludedKey = "stg";
    setDirOptions(uniqueDirectorates.filter(dir => dir.key.toLocaleLowerCase() !== excludedKey).sort((a, b) => a.text.localeCompare(b.text)));
   // setDirOptions(uniqueDirectorates.sort((a, b) => a.text.localeCompare(b.text)));
  };

  const reloadOrgChart = (): void => {
    if (!orgChartElement.current) return;

    teardownOrgChart();
    zoomInitializedRef.current = false;
    const zoom = initialZoomRef.current;
    setZoomLevel(zoom);

    requestAnimationFrame(() => {
      const container = document.getElementById("chart-container");
      if (container) {
        container.style.transform = `scale(${zoom})`;
        container.style.transformOrigin = "top left";
      }

      requestAnimationFrame(() => {
        resizeChartContainer();

        requestAnimationFrame(() => {
          if (selectedDirKey?.length) {
            const selDir = dirOptions.find(opt => opt.key.toString().trim().toLowerCase() === selectedDirKey.trim().toLowerCase()) ?? null;
            filterChartByDir(selDir);
          } else if (selectedNameOption) {
            filterChartByName(selectedNameOption);
          } else {
            renderDirectorateEmployeeOrgChart();
          }

          applyInitialZoomIfNeeded(); // Delayed zoom reset
        });
      });
    });
  };

 /* const resetZoom = (): void => {
    if (!orgChartElement.current) return;
  
    // Reset zoom state
    zoomInitializedRef.current = false;
    const zoom = initialZoomRef.current;
    setZoomLevel(zoom);
    zoomLevelRef.current = zoom;
  
    requestAnimationFrame(() => {
      const container = document.getElementById("chart-container");
      if (container) {
        container.style.transform = `scale(${zoom})`;
        container.style.transformOrigin = "top left";
      }
  
      requestAnimationFrame(() => {
        resizeChartContainer();
  
        requestAnimationFrame(() => {
          // Prefer selected directorate
          if (selectedDirKey?.length) {
            const selDir = dirOptions.find(opt => opt.key.toString().trim().toLowerCase() === selectedDirKey.trim().toLowerCase()) ?? null;
            filterChartByDir(selDir);
          }
          // Fallback to selected employee
          else if (selectedNameOption) {
            filterChartByName(selectedNameOption);
          }
          // Default rendering
          else {
            renderDirectorateEmployeeOrgChart();
          }
  
          // Delayed zoom re-application to align scroll
          applyInitialZoomIfNeeded();
        });
      });
    });
  }; */

  const resetZoom = (): void => {
    const wrapper = document.getElementById("orgChartWrapper");
    const chartInner = document.getElementById("chart-inner");
  
    if (!wrapper || !chartInner) return;
  
    const newZoom = initialZoomRef.current;
  
    // 1. Capture current scroll position relative to the chart dimensions
    const scrollLeftRatio = wrapper.scrollLeft / chartInner.scrollWidth;
    const scrollTopRatio = wrapper.scrollTop / chartInner.scrollHeight;
  
    // 2. Update zoom level and apply transform
    applyZoomChange(newZoom, true);
  
    // 3. Wait for the zoom to apply, then restore proportional scroll position
    requestAnimationFrame(() => {
      const newScrollLeft = chartInner.scrollWidth * scrollLeftRatio;
      const newScrollTop = chartInner.scrollHeight * scrollTopRatio;
  
      //wrapper.scrollLeft = newScrollLeft;
      //wrapper.scrollTop = newScrollTop;
      wrapper.scrollTo({ left: newScrollLeft, top: newScrollTop, behavior: "auto" });
    });
  };
  
  useEffect(() => {
    const link = document.createElement("link");
    link.href = "https://fonts.googleapis.com/icon?family=Material+Icons";
    link.rel = "stylesheet";
    document.head.appendChild(link);
  }, []);
  
  useEffect(() => {
    if (!props.data) return;
  
    // Populate employee & directorate dropdowns if not already populated
    if (nameOptions.length === 0) populateEmployeeDropdown(props.data);
    if (dirOptions.length === 0) populateDirectorateDropDown(props.data);
  }, [props.data]);

  useEffect(() => {
    if (!selectedNodeEmployee || nameOptions.length === 0) return;
  
    const matchedOption = nameOptions.find(opt => opt.key === selectedNodeEmployee.EmpO365Email) || null;
    setSelectedNameOption(matchedOption);
    setSelectedNameKey(matchedOption?.key.toString() || '');
  }, [selectedNodeEmployee, nameOptions]);

useEffect(() => {
  if (!props.data || props.data.length === 0) return;
  if (!orgChartElement.current) return;

  let cancelled = false;

  const renderTimer = window.setTimeout(() => {
    if (cancelled || !orgChartElement.current) return;

    if (selectedNodeEmployee) {
      renderDirectorateEmployeeOrgChart(selectedNodeEmployee);
    } else {
      renderDirectorateEmployeeOrgChart();
    }
  }, 100);

  return () => {
    cancelled = true;
    window.clearTimeout(renderTimer);

    // Do NOT empty the chart here.
  };

}, [props.data, selectedNodeEmployee]);

  useEffect(() => {
    zoomLevelRef.current = zoomLevel;
  }, [zoomLevel]);
  
  useEffect(() => {
    const container = document.getElementById("chart-container");
    if (!container) return;
  
    let isZooming = false;
  
    const throttledWheel = throttle((e: WheelEvent) => {
      const isZoomGesture = e.ctrlKey || e.metaKey;
  
      if (!isZoomGesture) {
        setShowZoomHint(true);
        setTimeout(() => setShowZoomHint(false), 2000);
        return;
      }
  
      e.preventDefault();
  
      if (isZooming) return;
      isZooming = true;
  
      const delta = e.deltaY < 0 ? 0.02 : -0.02;
  
      applyZoomChange(delta);
  
      // Reset after short delay
      setTimeout(() => {
        isZooming = false;
      }, 100); // 100ms debounce
    }, 50);
  
    container.addEventListener("wheel", throttledWheel, { passive: false });
  
    return () => {
      container.removeEventListener("wheel", throttledWheel);
      throttledWheel.cancel();
    };
  }, []);  

  useEffect(() => {
    const chartContainer = document.getElementById("chart-container");
    const chartInner = document.getElementById("chart-inner");
  
    const setFullViewportSize = (): void => {
      const vw = window.innerWidth;
      const vh = window.innerHeight;
  
      if (chartContainer) {
        chartContainer.style.width = `${vw}px`;
        chartContainer.style.height = `${vh}px`;
        chartContainer.style.minWidth = `${vw}px`;
        chartContainer.style.minHeight = `${vh}px`;
      }
  
      if (chartInner) {
        chartInner.style.width = `${vw}px`;
        chartInner.style.height = `${vh}px`;
        chartInner.style.minWidth = `${vw}px`;
        chartInner.style.minHeight = `${vh}px`;
      }
    };
  
    setFullViewportSize();
    window.addEventListener('resize', setFullViewportSize);
    
    return () => {
      window.removeEventListener('resize', setFullViewportSize);
    };
  }, []);
  

  useEffect(() => {
    const wrapper = document.getElementById("orgChartWrapper");
    if (!wrapper) return;
  
    let isDragging = false;
    let startX = 0;
    let startY = 0;
    let scrollLeft = 0;
    let scrollTop = 0;
  
    const handleMouseDown = (e: MouseEvent): void => {
      isDragging = true;
      wrapper.classList.add("grabbing");
      startX = e.pageX - wrapper.offsetLeft;
      startY = e.pageY - wrapper.offsetTop;
      scrollLeft = wrapper.scrollLeft;
      scrollTop = wrapper.scrollTop;
    };
  
    const handleMouseLeave = (): void => {
      isDragging = false;
      wrapper.classList.remove("grabbing");
    };
  
    const handleMouseUp = (): void => {
      isDragging = false;
      wrapper.classList.remove("grabbing");
    };
  
    const handleMouseMove = (e: MouseEvent): void => {
      if (!isDragging) return;
      e.preventDefault();
      const x = e.pageX - wrapper.offsetLeft;
      const y = e.pageY - wrapper.offsetTop;
      const walkX = x - startX;
      const walkY = y - startY;
      wrapper.scrollLeft = scrollLeft - walkX;
      wrapper.scrollTop = scrollTop - walkY;
    };
  
    wrapper.addEventListener("mousedown", handleMouseDown);
    wrapper.addEventListener("mouseleave", handleMouseLeave);
    wrapper.addEventListener("mouseup", handleMouseUp);
    wrapper.addEventListener("mousemove", handleMouseMove);
  
    return () => {
      wrapper.removeEventListener("mousedown", handleMouseDown);
      wrapper.removeEventListener("mouseleave", handleMouseLeave);
      wrapper.removeEventListener("mouseup", handleMouseUp);
      wrapper.removeEventListener("mousemove", handleMouseMove);
    };
  }, []);  

  return (
    <div>
   
      <Box
      sx={{
        backgroundColor: "#f0f0f0",
        padding: "15px",
        borderRadius: "8px",
        boxShadow: "0 4px 6px rgba(0, 0, 0, 0.1)",
        display: "flex",
        justifyContent: "space-between",
        alignItems: "flex-end",
        gap: "50px",
        height: "100%"
      }}
    >

      <Grid container spacing={3} alignItems="flex-end">
        {/* Directorate ComboBox */}
        <Grid item sx={{ minWidth: 450 }}>
          <Typography variant="body1" sx={{ fontFamily: "Apparat", fontSize: "0.875rem !important" }}>
            Highlight Directorate:
          </Typography>
          <Autocomplete
            sx={{
              width: "100%",
              maxWidth: 450,
              "& .MuiInputBase-root": {
                backgroundColor: "white",
                borderRadius: "4px",
                height: "28px !important",
                padding: "0px",
                fontFamily: "Apparat",
                fontSize: "0.7rem !important",
                display: "flex",
                alignItems: "center",
              },
              "& .MuiOutlinedInput-root": {
                height: "28px !important",
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
                  height: "16px !important",
                  padding: "2px 6px !important",
                  fontFamily: "Apparat",
                  fontSize: "0.7rem !important",
                },
              },
              "& .MuiAutocomplete-popper": {
                fontSize: "0.7rem !important",
              },
              "& .MuiAutocomplete-listbox": {
                fontSize: "0.7rem !important",
                padding: "2px 0",
              },
              "& .MuiAutocomplete-option": {
                minHeight: "26px",
                fontSize: "0.7rem !important",
              },
            }}
            options={dirOptions}
            getOptionLabel={(option) => option.text}
            filterOptions={(options, state) =>
              options.filter((option) => option.text.toLowerCase().includes(state.inputValue.toLowerCase()))
            }
            value={selectedDirKey ? dirOptions.find((d) => d.key === selectedDirKey) : null}
            onChange={(event, value, reason) => {
              if (reason === "clear") { 
                setSelectedDirKey("");
                setSelectedNameKey("");
                setSelectedNameOption(null);
                filterChartByReset();
              } else {
                setSelectedDirKey(value?.key as string || "");
                setSelectedNameOption(null);
                setSelectedNameKey("");
                filterChartByDir(value);
              }
            }}
            renderInput={(params) => (
              <TextField
                {...params}
                placeholder="Please select a directorate"
                aria-label="Directorate selection"
                variant="outlined"
                InputLabelProps={{ shrink: true }}
                sx={{
                  backgroundColor: "white",
                  borderRadius: "4px",
                  "& .MuiOutlinedInput-root": {
                    height: "28px !important",
                    padding: "0px",
                    "& input": {
                      height: "16px !important",
                      padding: "2px 6px !important",
                      fontFamily: "Apparat",
                      fontSize: "0.7rem !important",
                    },
                  },
                }}
              />
            )}
          />
        </Grid>

        {/* Employee ComboBox */}
        <Grid item sx={{ minWidth: 320 }}>
            <Typography variant="body1" sx={{ fontFamily: "Apparat", fontSize: "0.875rem !important" }}>
              Highlight Employee:
            </Typography>
            <Autocomplete
              sx={{
                width: "100%",
                maxWidth: 320,
                "& .MuiInputBase-root": {
                  backgroundColor: "white",
                  borderRadius: "4px",
                  height: "28px !important",
                  padding: "0px",
                  fontFamily: "Apparat",
                  fontSize: "0.7rem !important",
                  display: "flex",
                  alignItems: "center",
                },
                "& .MuiOutlinedInput-root": {
                  height: "28px !important",
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
                    height: "16px !important",
                    padding: "2px 6px !important",
                    fontFamily: "Apparat",
                    fontSize: "0.7rem !important",
                  },
                },
                "& .MuiAutocomplete-popper": {
                  fontSize: "0.7rem !important",
                },
                "& .MuiAutocomplete-listbox": {
                  fontSize: "0.7rem !important",
                  padding: "2px 0",
                },
                "& .MuiAutocomplete-option": {
                  minHeight: "26px",
                  fontSize: "0.7rem !important",
                },
              }}
              options={nameOptions}
              getOptionLabel={(option) => option.text}
              value={selectedNameOption}
              filterOptions={(options, state) =>
                options.filter((option) => option.text.toLowerCase().includes(state.inputValue.toLowerCase()))
              }
              onChange={(event, value, reason) => {
                if (reason === "clear") { 
                  setSelectedNameKey("");
                  setSelectedNameOption(null);
                  setSelectedDirKey("");
                  filterChartByReset(); // ✅ Call reset function when clearing
                } else {
                  setSelectedNameKey(value?.key as string || "");
                  setSelectedNameOption(value);
                  setSelectedDirKey("");
                  filterChartByName(value);
                }
              }}
              renderInput={(params) => (
                <TextField
                  {...params}
                  placeholder="Please select an employee"
                  aria-label="Employee selection"
                  variant="outlined"
                  InputLabelProps={{ shrink: true }}
                  sx={{
                    backgroundColor: "white",
                    borderRadius: "4px",
                    "& .MuiOutlinedInput-root": {
                      height: "28px !important",
                      padding: "0px",
                      "& input": {
                        height: "16px !important",
                        padding: "2px 6px !important",
                        fontFamily: "Apparat",
                        fontSize: "0.7rem !important",
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
      <Box display="flex" alignItems="center" gap={2}>
        <Box display="flex" alignItems="center" gap={1}>
            {/* Zoom Controls */}
            <Box
            sx={{
              display: "flex",
              alignItems: "center",
              backgroundColor: "rgba(255,255,255,0.9)",
              borderRadius: "10px",
              padding: "6px 10px",
              boxShadow: "0 4px 12px rgba(0,0,0,0.1)"
            }}
          >
            <Tooltip title="Zoom In">
              <IconButton
                size="small"
                sx={{ color: "#2A7386", padding: "6px" }}
                onClick={zoomIn}
                disabled={zoomLevel >= MAX_ZOOM}
                aria-label="Zoom In"
              >
                <ZoomInIcon style={{ fontSize: "20px" }} />
              </IconButton>
            </Tooltip>
            <Tooltip title="Zoom Out">
              <IconButton
              size="small"
              sx={{ color: "#2A7386", padding: "6px" }}
              onClick={zoomOut}
              disabled={zoomLevel <= MIN_ZOOM}
              aria-label="Zoom Out"
              >
              <ZoomOutIcon style={{ fontSize: "20px" }} />
            </IconButton>
            </Tooltip>
            <Tooltip title="Reset Zoom">
              <IconButton
                size="small"
                sx={{ color: "#2A7386", padding: "6px" }}
                onClick={() => {
                  resetZoom();
                }}
                aria-label="Reset Zoom"
              >
                <CenterFocusStrong style={{ fontSize: "20px" }} />
              </IconButton>
            </Tooltip>
            <Tooltip title="Current Zoom">
              <Typography
                sx={{
                  marginLeft: "8px",
                  marginRight: "8px",
                  backgroundColor: "#e5f6f6",
                  padding: "2px 8px",
                  borderRadius: "8px",
                  color: "#2A7386",
                  fontSize: "0.75rem",
                  fontWeight: 600,
                  minWidth: "40px",
                  textAlign: "center"
                }}
              >
                {(zoomLevel * 100).toFixed(0)}%
              </Typography>
            </Tooltip>
            <Tooltip title="Reload">
              <IconButton
                size="small"
                sx={{ color: "#2A7386", padding: "6px" }}
                aria-label="Reload"
                onClick={() => reloadOrgChart()}
              >
                <RefreshIcon style={{ fontSize: "20px" }} />
              </IconButton>
            </Tooltip>
            <Tooltip title="Print Chart">
              <IconButton
                size="small"
                sx={{ color: "#2A7386", padding: "8px" }}
                aria-label="Print Chart"
                onClick={async () => {
                  setTimeout(async () => {
                    try {
                      await exportOrgChartToPDF();
                    } catch (err) {
                      console.error("Print failed:", err);
                    }
                  }, 300);
                }}
              >
                <PrintIcon style={{ fontSize: "22px" }} />
              </IconButton>
            </Tooltip>
          </Box>
        </Box>
      </Box>

      </Box>

      <Box
      id="orgChartOuter"
      sx={{
        position: "relative", // critical!
        width: "100%",
        height: "calc(100vh - 100px)", // or 100vh depending on your layout
        overflow: "hidden", // optional
      }}
    >
      <Box
        id="orgChartWrapper"
        ref={orgChartWrapperRef}
        sx={{
          position: "absolute", // so it scrolls over the full outer box
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          overflow: "auto",
          cursor: "grab",
          backgroundColor: "#fbfbfc",
          }}
      >
        {/* CHART CONTAINER */}
        <Box
          id="chart-container"
          className="chartContainer"
          sx={{
            width: 'auto', // REMOVE fixed 100vw or 2560px
            minWidth: '100%',
            height: 'auto',
            minHeight: '100%',
            backgroundColor: "#fbfbfc",
            opacity: orgChartData ? 1 : 0,
            position: 'relative',
            padding: '0',
            overflow: "unset",
            imageRendering: "optimizeQuality",
            backfaceVisibility: "hidden"
           /* "&.chartContainer": {
              overflow: "visible !important",
            },*/
          }}
        >
          <div id="chart-inner"
            style={{
              /*width: '100vw',
              height: '100vh',
              minWidth: '100vw',
              minHeight: '100vh',*/
              width: '100%',
              height: '100%',
              minWidth: '100%',
              minHeight: '100%',
              display: "inline-block",
              cursor: "grab",
              position: "relative",
            }}>
            <div
              ref={orgChartElement}
              style={{
                minHeight: "100%",
                minWidth: "100%",
                cursor: "grab",
              }}
            />
          </div>
        </Box>
      </Box>
      </Box>
    
      <EmployeeDetailsPopUp serviceScope={props.serviceScope} />

      <Snackbar
        open={showZoomHint}
        message="Hold Ctrl (or Cmd on Mac) and scroll to zoom"
        anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
        ContentProps={{
          sx: {
            backgroundColor: "#2A7386",
            color: "#fff",
            fontSize: "0.8rem",
            borderRadius: "6px"
          }
        }}
      />

    </div>
  );
};