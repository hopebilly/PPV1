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

interface IDirectorateOrgChartprops {
  data: IEmployee[] | undefined;
  Directorate: string | undefined;
  serviceScope: ServiceScope;
}

export const DirectorateOrgChart: FC<IDirectorateOrgChartprops> = (props) => {
  
  let $orgChartElement: any;
  const orgChartElement: any = useRef<any>(undefined);
  const [nameOptions, setNameOptions] = useState<IComboBoxOption[]>([]);
  const [dirOptions, setDirOptions] = useState<IComboBoxOption[]>([]);
  const [selectedNameOption, setSelectedNameOption] = useState<IComboBoxOption | null>(null);
  const [selectedDirKey, setSelectedDirKey] = useState<string>("");
  const [orgChartData, setOrgChartData] = useState<IEmployee | null>(null);
  const [isNameAutocompleteEnabled, setIsNameAutocompleteEnabled] = useState(false);
  const orgChartWrapperRef = useRef<HTMLDivElement | null>(null);
  const { selectedNodeEmployee, setSelectedNodeEmployee,
          setSelectedNameKey, setEmployeesForOrgChart, selectedDirectorate,
          setSelectedEmployee, setEmployeeModalOpen} = useSharedContext();
  const chartInstanceRef = useRef<any>(null);
  const initialZoomRef = useRef<number>(0.65);
  const zoomInitializedRef = useRef<boolean>(false);
  const zoomLevelRef = useRef<number>(initialZoomRef.current);
  const [zoomLevel, setZoomLevel] = useState<number>(initialZoomRef.current);
  const [showZoomHint, setShowZoomHint] = useState(false);

  const MIN_ZOOM = 0.3;
  const MAX_ZOOM = 2;
  const ZOOM_STEP = 0.05;

  const snapZoom = (value: number): number => Math.round(value * 100) / 100;
  
  // Function to apply zoom with accurate scroll correction
  function applyZoomChange(value: number, isAbsolute = false): void {
    const wrapper = document.getElementById("orgChartWrapper");
    const container = document.getElementById("chart-container");
    const chartInner = document.getElementById("chart-inner");
  
    if (!wrapper || !container || !chartInner) return;
  
    const oldZoom = zoomLevelRef.current;
    const rawZoom = isAbsolute ? value : oldZoom + value;
    const newZoom = snapZoom(Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, rawZoom)));
    if (newZoom === oldZoom) return;
  
    // Current horizontal center before zoom
    const centerX = wrapper.scrollLeft + wrapper.clientWidth / 2;
    const relX = centerX / (chartInner.scrollWidth * oldZoom);
  
    // Update zoom level
    zoomLevelRef.current = newZoom;
    setZoomLevel(newZoom);
    container.style.transform = `scale(${newZoom})`;
    container.style.transformOrigin = "top left";
    container.style.transition = "transform 0.2s ease";
  
    // Always scroll to top to keep Tier 1 visible and center horizontally
    requestAnimationFrame(() => {
      const scaledWidth = chartInner.scrollWidth * newZoom;
      const scrollLeft = Math.max(0, scaledWidth * relX - wrapper.clientWidth / 2);
  
      wrapper.scrollTo({
        left: scrollLeft,
        top: 0, // Force vertical scroll to top
        behavior: "auto"
      });
    });
  }
  
  const zoomIn = (): void => applyZoomChange(ZOOM_STEP);
  const zoomOut = (): void => applyZoomChange(-ZOOM_STEP);
  
  /*const zoomIn = (): void => {
    applyZoomChange(0.1);
  };
  
  const zoomOut = (): void => {
    applyZoomChange(-0.1);
  };*/
  
 /* const resetZoom = (): void => {
    const wrapper = document.getElementById("orgChartWrapper");
    const container = document.getElementById("chart-container");
    const chartInner = document.getElementById("chart-inner");
    const orgChart = chartInner?.querySelector(".orgchart") as HTMLElement;
  
    if (!wrapper || !container || !chartInner || !orgChart) return;
  
    zoomInitializedRef.current = false;
    const zoom = initialZoomRef.current;
    setZoomLevel(zoom);
    zoomLevelRef.current = zoom;
  
    container.style.transform = `scale(${zoom})`;
    container.style.transformOrigin = "top left";
    container.style.transition = "transform 0.2s ease";
  
    // Matching reload logic
    requestAnimationFrame(() => {
      resizeChartContainer();
  
      requestAnimationFrame(() => {
        if (selectedNameOption?.key) {
          const emailId = selectedNameOption.key.toString().toLowerCase();
          waitAndFocusNode(emailId);
        } else {
          centerOrgChart1(zoom);
        }
      });
    });
  };*/

  /*function resetZoom(): void {
    const wrapper = document.getElementById("orgChartWrapper");
    const container = document.getElementById("chart-container");
    const chartInner = document.getElementById("chart-inner");

    if (!wrapper || !container || !chartInner) return;

    const oldZoom = zoomLevelRef.current;
    const newZoom = snapZoom(initialZoomRef.current);

    // Capture ratios BEFORE zoom changes
    const scrollRatioX = wrapper.scrollLeft / (chartInner.scrollWidth * oldZoom);
    const scrollRatioY = wrapper.scrollTop / (chartInner.scrollHeight * oldZoom);

    // Clear transform to prep for clean reset
    container.style.transition = "none";
    container.style.transform = "none";

    // Set zoom state refs
    zoomLevelRef.current = newZoom;
    setZoomLevel(newZoom);

    // Resize using new zoom value
    requestAnimationFrame(() => {
      resizeChartContainer();

      requestAnimationFrame(() => {
        // Apply transform now that layout is updated
        container.style.transition = "transform 0.2s ease";
        container.style.transform = `scale(${newZoom})`;
        container.style.transformOrigin = "top left";

        // Scroll using new chart size and old ratios
        requestAnimationFrame(() => {
          const scaledWidth = chartInner.scrollWidth * newZoom;
          const scaledHeight = chartInner.scrollHeight * newZoom;

          wrapper.scrollTo({
            left: Math.max(0, scaledWidth * scrollRatioX),
            top: Math.max(0, scaledHeight * scrollRatioY),
            behavior: "auto"
          });
        });
      });
    });
  }*/
   
    function resetZoom(): void {
      const wrapper = document.getElementById("orgChartWrapper");
      const container = document.getElementById("chart-container");
      const chartInner = document.getElementById("chart-inner");
    
      if (!wrapper || !container || !chartInner) return;
    
      const oldZoom = zoomLevelRef.current;
      const newZoom = snapZoom(initialZoomRef.current);
    
      // Capture current scroll ratio (safe fallback)
      const scrollRatioX = wrapper.scrollLeft / (chartInner.scrollWidth * oldZoom);
      const scrollRatioY = wrapper.scrollTop / (chartInner.scrollHeight * oldZoom);
    
      // Step 1: Clear existing transform immediately
      container.style.transition = "none";
      container.style.transform = "";
    
      // Step 2: Update zoom level in refs/state
      zoomLevelRef.current = newZoom;
      setZoomLevel(newZoom);
    
      // Step 3: Resize container based on new zoom
      requestAnimationFrame(() => {
        resizeChartContainer();
    
        requestAnimationFrame(() => {
          container.style.transition = "transform 0.2s ease";
          container.style.transform = `scale(${newZoom})`;
          container.style.transformOrigin = "top left";
    
          // Step 4: Adjust scroll after layout settles
          requestAnimationFrame(() => {
            const scaledWidth = chartInner.scrollWidth * newZoom;
            const scaledHeight = chartInner.scrollHeight * newZoom;
    
            wrapper.scrollTo({
              left: scaledWidth * scrollRatioX,
              top: scaledHeight * scrollRatioY,
              behavior: "auto"
            });
    
            // Optional: one more resize pass (edge case safety)
            requestAnimationFrame(() => {
              resizeChartContainer();
            });
          });
        });
      });
    }

  function applyInitialZoomIfNeeded(): void {
    if (zoomInitializedRef.current) return;
    const zoom = initialZoomRef.current;
    setZoomLevel(zoom);
    zoomInitializedRef.current = true;
  
    requestAnimationFrame(() => {
      const container = document.getElementById("chart-container");
      if (container) {
        container.style.transform = `scale(${zoom})`;
        container.style.transformOrigin = "top left";
      }
  
      requestAnimationFrame(() => {
        resizeChartContainer();
        centerOrgChart1(zoom);
      });
    });
  } 

  function focusOnNode(nodeElement: HTMLElement): void {
    const chartWrapper = document.getElementById("orgChartWrapper");
    const chartContainer = document.getElementById("chart-inner");
  
    if (!chartWrapper || !chartContainer || !nodeElement) return;
  
    const wrapperRect = chartWrapper.getBoundingClientRect();
    const nodeRect = nodeElement.getBoundingClientRect();
  
    // Calculate scroll positions to center the node
    const scrollX = Math.max(
      0,
      chartWrapper.scrollLeft + (nodeRect.left - wrapperRect.left) - (wrapperRect.width / 2) + (nodeRect.width / 2)
    );
    const scrollY = Math.max(
      0,
      chartWrapper.scrollTop + (nodeRect.top - wrapperRect.top) - (wrapperRect.height / 2) + (nodeRect.height / 2)
    );
  
    // Optional: Add horizontal buffer to avoid clipping on left/right
    const horizontalBuffer = 200; // pixels of safe space on both sides
  
    const paddedWidth = Math.max(chartContainer.scrollWidth, scrollX + wrapperRect.width + horizontalBuffer);
  
    chartContainer.style.width = `${paddedWidth}px`;
    chartContainer.style.minWidth = `${paddedWidth}px`;
    // Keep height as-is
    chartContainer.style.height = `${chartContainer.scrollHeight}px`;
    chartContainer.style.minHeight = `${chartContainer.scrollHeight}px`;
  
    // Smooth scroll to node
    chartWrapper.scrollTo({ left: scrollX, top: scrollY, behavior: "smooth" });
  }
  
  
  function waitAndFocusNode(emailId: string, retries = 10): void {
    const target = document.querySelector(`.nodeTemplate[data-id="${emailId}"]`)?.closest('.node') as HTMLElement;
    
    if (target) {
      setTimeout(() => focusOnNode(target), 50);  // Only focus
    } else if (retries > 0) {
      setTimeout(() => waitAndFocusNode(emailId, retries - 1), 100); // Retry until available
    }
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

  // Function to find an employee by ID
  const getEmployeeDataById = (nodeId: string): IEmployee | undefined => {
    return props.data?.find(emp => emp.EmpO365Email === nodeId);
  };

  const handleRenderComplete = (): void => {
    requestAnimationFrame(() => {
      resizeChartContainer();        // Recalculate layout without transform hacks
      centerOrgChart1(zoomLevel);    // Center chart with current zoom
    });
  };  

  function afterDOMUpdate(callback: () => void): void {
    requestAnimationFrame(callback);
  }

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

    const chartInstance = $orgChartElement.orgchart({
      data: organizationChartData ? organizationChartData : {},
      nodeContent: "id",
      nodeID: "id",
      zoom: true,
      pan: true,
      //verticalLevel: 4,
      initCollapsed: true,
      visibleLevel: getMatchedNodeTier(matchedNode),
      toggleSiblingsResp: true,
      //direction: "t2b",
      nodeTemplate: (data: any) => {
        const eNameFormatted = `${data.FirstName || ''} ${data.LastName || ''}`;
        const nodeId = data.EmpO365Email;
        //const nodeLevel = data.levelOffset && data.levelOffset > 0 ? data.level + data.levelOffset : data.level;
        const nodeLevel = (data.level ?? 0) + (data.levelOffset ?? 0);
        const acronym = data.KeyStaff ? "KEY" : data.DirectorateHead ? "ELT" : "";

        /*const jobTitle = data.JobTitle || "";
        const isLongJobTitle = jobTitle.length > 25; 

        const inlineStyle = isLongJobTitle ? 'min-height: 130px;' : '';*/
        const inlineStyle = 'min-height: 130px;';

        let nodeStyle;
        switch (nodeLevel) {
          case 1: nodeStyle = "nodeTemplate nodeTemplate1"; break;
          case 2: nodeStyle = "nodeTemplate nodeTemplate2"; break;
          case 3: nodeStyle = "nodeTemplate nodeTemplate3"; break;
          case 4: nodeStyle = "nodeTemplate nodeTemplate4"; break;
          case 5: nodeStyle = "nodeTemplate nodeTemplate5"; break;
          case 6: nodeStyle = "nodeTemplate nodeTemplate6"; break;
          case 7: nodeStyle = "nodeTemplate nodeTemplate7"; break;
          default: nodeStyle = "nodeTemplate"; break;
        }

        return `
          <div class="${nodeStyle}" style="${inlineStyle}" data-tier="${data.Tier}" data-id="${nodeId}" 
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
            
      }
    });

    chartInstanceRef.current = chartInstance;

    $orgChartElement.on('click', '.leftEdge, .rightEdge, .bottomEdge', () => {
      requestAnimationFrame(() => {
        syncAllToggleIcons();
      });
    });
    
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

    // Fake "renderComplete"
    requestAnimationFrame(() => {
      //setTimeout(() => {
        //console.log("✅ Safe fallback: renderComplete manually triggered");
        handleRenderComplete();
      //}, 0);
    });
            
    if (!$orgChartElement.data("touchEventSet")) {
      $orgChartElement.on("touchmove", (event: JQuery.Event) => {
        event.preventDefault();
      });
      $orgChartElement.data("touchEventSet", true);
    }
  };

  const updateEmployeeHighlighting = (node: IEmployee, emailId: string): { updatedNode: IEmployee; highlightedMatchedNode?: IEmployee } => {
    // Check if the current node matches the email ID
    const isMatch = node.EmpO365Email?.toLowerCase() === emailId;

    // Recursively update children first
    const updatedChildrenResults = node.children?.map((child) => updateEmployeeHighlighting(child, emailId)) || [];

    // Extract updated children and find the matched node if any
    const updatedChildren = updatedChildrenResults.map(result => result.updatedNode);
    const highlightedMatchedNode = isMatch ? { ...node, children: [] } : updatedChildrenResults.find(result => result.highlightedMatchedNode)?.highlightedMatchedNode;

    // Check if any child is highlighted (to ensure parent expansion)
    const isChildHighlighted = updatedChildrenResults.some(result => result.highlightedMatchedNode || result.updatedNode._highlighted);

    return {
      updatedNode: {
        ...node,
        _highlighted: isMatch,
        _expanded: isMatch || isChildHighlighted || node._expanded, // Expand node if matched or a child is highlighted
        children: updatedChildren, // Update children recursively
      },
      highlightedMatchedNode,
    };
  };

  // Second Pass: Assign tiers to each node based on their roles and parent-child relationships
  const assignTiers = (
    node: IEmployee | undefined,
    matchNode: IEmployee | undefined,
    parentTier: number = 0
  ): { tierAssignedHierarchy: IEmployee | undefined; tieredMatchedNode: IEmployee | undefined } => {
    if (!node) return { tierAssignedHierarchy: undefined, tieredMatchedNode: undefined };
  
    let tieredMatchedNode: IEmployee | undefined = undefined;
  
    // Assign the tier for the current node based on its parent tier
    if (node.parentId === '') {
      node.Tier = 3;
    } else if (parentTier === 3) {
      node.Tier = 4;
    } else if (parentTier === 4) {
      node.Tier = 5;
    } else if (parentTier === 5) {
      node.Tier = 6;
    } else if (parentTier >= 6) {
      node.Tier = 7;
    }
  
    // Match the node if matchNode is provided
    if (
      matchNode &&
      node.EmpO365Email?.trim().toLowerCase() ===
        matchNode.EmpO365Email?.trim().toLowerCase()
    ) {
      tieredMatchedNode = node;
    }
  
    // Recursively assign tiers to children and propagate matched node
    if (node.children) {
      node.children = node.children
        .map(child => {
          const { tierAssignedHierarchy: updatedChild, tieredMatchedNode: matchedInChild } = assignTiers(child, matchNode, node.Tier ?? 0);
          if (matchedInChild) {
            tieredMatchedNode = matchedInChild;
          }
          return updatedChild;
        })
        .filter((child): child is IEmployee => !!child); // Removes undefined values safely
    }
  
    return { tierAssignedHierarchy: node, tieredMatchedNode };
  };
  
  // Function to build hierarchy from flat employee list
const buildHierarchy = (employees: IEmployee[]): IEmployee | undefined => {
    if (!employees || employees.length === 0) return undefined;

    const employeeMap = new Map<string, IEmployee>();

    // Create a unique key for EACH leadership assignment.
    // Same person can therefore appear under multiple directorates.
    const getNodeKey = (emp: IEmployee): string => {
        return `${emp.id}__${emp.Directorate?.trim().toLowerCase() || "no-directorate"}`;
    };

    // Create every node separately.
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

        // Root employee
        if (!emp.parentId) {
            root = child;
            return;
        }

        // Find possible parent records.
        const possibleParents = employees.filter(
            parent => parent.id === emp.parentId
        );

        if (possibleParents.length === 0) {
            return;
        }

        // Prefer a parent from the SAME directorate.
        // This is important when one leader manages multiple directorates.
        const sameDirectorateParent = possibleParents.find(
            parent =>
                parent.Directorate?.trim().toLowerCase() ===
                emp.Directorate?.trim().toLowerCase()
        );

        const parentRecord =
            sameDirectorateParent || possibleParents[0];

        const parentKey = getNodeKey(parentRecord);
        const parent = employeeMap.get(parentKey);

        if (parent) {
            parent.children = parent.children ?? [];
            parent.children.push(child);
        }
    });

    return root;
};

  const renderDirectorateEmployeeOrgChart = (
    employeeListForOrgChart: IEmployee[], 
    matchedNode: IEmployee | undefined
  ): void => {
      if (!orgChartElement.current) return; // Guard clause
      teardownOrgChart();
      
      const hierarchy = buildHierarchy(employeeListForOrgChart);
      const { tierAssignedHierarchy: tierAssignedHierarchy, tieredMatchedNode } = assignTiers(hierarchy, matchedNode);
      
      let updatedTierAssignedHierarchy = tierAssignedHierarchy;
      let highlightedMatchedNode: IEmployee | undefined = undefined;

      if(tierAssignedHierarchy){
        if(tieredMatchedNode?.EmpO365Email?.length){
          const result  = updateEmployeeHighlighting(tierAssignedHierarchy, tieredMatchedNode.EmpO365Email);
          updatedTierAssignedHierarchy = result.updatedNode;
          highlightedMatchedNode = result.highlightedMatchedNode;
        }

         // Ensure DOM is fully updated before rendering
        //setTimeout(() => {
        requestAnimationFrame(() => {
          if (updatedTierAssignedHierarchy) {
            setOrgChartData(updatedTierAssignedHierarchy);
            // Inside renderDirectorateEmployeeOrgChart()
            zoomInitializedRef.current = false;
            applyInitialZoomIfNeeded();

            renderOrgChart(updatedTierAssignedHierarchy, highlightedMatchedNode);
            updateNodeStrokes();
            resizeChartContainer(); 
          }
        });
        //}, 0); // or use requestAnimationFrame
      }
  }

  const determineRootNodeSupervisor = (employeeList: IEmployee[], rootNode: IEmployee, supervisorList: IEmployee[] = []): IEmployee[] => {
   
    if(rootNode.SupervisorO365Email?.length){

        const parentNode = employeeList.find(n => n.EmpO365Email?.trim().toLowerCase() === rootNode.SupervisorO365Email?.trim().toLowerCase());
        
        if(parentNode){
          supervisorList.push(parentNode);

          const childrenNodes = employeeList.filter(n => (n.SupervisorO365Email === parentNode.EmpO365Email) && n.EmpO365Email !== rootNode.EmpO365Email);
          supervisorList.push(...childrenNodes);
          
          return determineRootNodeSupervisor(employeeList, parentNode, supervisorList);
        }
    }

    return supervisorList;
  }

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
  
  // Preprocess employee list to set expanded/collapsed states
  const preprocessEmployeeList = (directorate: string, employeeList: IEmployee[], employeesForOrgChart?: IEmployee[]): IEmployee[] => {
  
    let rootNode: IEmployee | undefined;

    // Create a deep copy of the employeeList to avoid mutation
    let processedData = JSON.parse(JSON.stringify(employeeList)) as IEmployee[];

    if (directorate?.length) {
      processedData = processedData.filter(node =>
        (node.Directorate?.trim().toLocaleLowerCase() || "") === directorate.trim().toLocaleLowerCase()
      );

      rootNode = processedData.find(node => node.DirectorateHead?((node.DirectorateHead.trim().toLowerCase() || "") === "yes")
                                                                :(node.KeyStaff? (node.KeyStaff.trim().toLowerCase() === "yes"): undefined));
      //rootNode = processedData.find(node => (node.DirectorateHead?.trim().toLowerCase() || "") === "yes");
      if (rootNode) {
        rootNode.parentId = "";
      }        
    }

    // Recursively mark nodes as expanded or collapsed based on employeesForOrgChart
    processedData.forEach(node => {
      // Ensure _expanded is false for all nodes before applying any specific logic
      node._expanded = false;
      markExpandedNodes(node, employeesForOrgChart, rootNode)
  });

  return processedData;
  };

  const populateEmployeeNameDropdown = (directorate: string): IEmployee[] => {

    let filteredOrderedProcessedEmpList: IEmployee[] = [];
    let newNameOptions: IComboBoxOption[] = [];

    if(!props.data) 
      return filteredOrderedProcessedEmpList;

    const processedEmpList = preprocessEmployeeList(directorate, props.data, undefined);

    if (processedEmpList?.length) {
    let directorateRootNode: IEmployee | undefined;
      if(directorate.trim().toLowerCase() === 'exec'){
        directorateRootNode = processedEmpList.find(node => node.parentId === '');
      } else{
        directorateRootNode = processedEmpList.find(node => node.parentId === '' 
          && (node.SupervisorO365Email && (node.SupervisorO365Email !== node.EmpO365Email)));
      }
      
      if(directorateRootNode)
      {
        const supervisorToBeRemoved = determineRootNodeSupervisor(processedEmpList, directorateRootNode);
        if(supervisorToBeRemoved?.length) {
          filteredOrderedProcessedEmpList = processedEmpList.filter(node => 
            !supervisorToBeRemoved.some(supervisor => supervisor.EmpO365Email === node.EmpO365Email)
          );
        }
        else
        {
          filteredOrderedProcessedEmpList = processedEmpList;
        }
      }
    }
    
    newNameOptions = filteredOrderedProcessedEmpList
    .filter(e => e.EmployeeName && e.EmpO365Email) // Ensure properties exist
    .sort((a, b) => (a.EmployeeName ?? "").localeCompare(b.EmployeeName ?? "")) // Handle undefined
    .map(e => ({
      key: e.EmpO365Email ?? "", // Provide a fallback (empty string)
      text: e.EmployeeName ?? "", // Provide a fallback (empty string)
    }));

    setNameOptions(newNameOptions); // Ensure state update with new array
    setIsNameAutocompleteEnabled(newNameOptions?.length > 0);

    return filteredOrderedProcessedEmpList;

  };

  const directorateOptionChange = (directorate?: string): void => {
    if(!directorate) return;

    const nameOptionsEmployeeList = populateEmployeeNameDropdown(directorate);

    const wrapper = document.getElementById("orgChartWrapper");
    if (wrapper) {
      wrapper.scrollTo({ top: 0, left: 0, behavior: "smooth" });
    }

    if(nameOptionsEmployeeList?.length){
      zoomInitializedRef.current = false;
      applyInitialZoomIfNeeded();
      renderDirectorateEmployeeOrgChart(nameOptionsEmployeeList, undefined);
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
    //setDirOptions(uniqueDirectorates.sort((a, b) => a.text.localeCompare(b.text)));
  };

  const resetZoomAndScroll = (): void => {

    zoomInitializedRef.current = false;
    applyInitialZoomIfNeeded();
    
    const wrapper = document.getElementById("orgChartWrapper");
    if (wrapper) {
      wrapper.scrollTo({ top: 0, left: 0, behavior: "auto" });
    }
  };

  /**
  * Selects and highlights the selected employee node
  * @param value - Selected Employee ComboBox option
  */
  const filterChartByName = (value: IComboBoxOption | null): void => {
    if (!value || value.key === undefined || value.key === null) return;
  
    const emailId = value.key.toString().toLowerCase();
  
    if (orgChartData) {
      const { updatedNode: updatedOrgChartData, highlightedMatchedNode } = updateEmployeeHighlighting(orgChartData, emailId);
      setOrgChartData(updatedOrgChartData);
  
      renderOrgChart(updatedOrgChartData, highlightedMatchedNode);
  
      // 1. Apply Zoom
      requestAnimationFrame(() => {
        const container = document.getElementById("chart-container");
        if (container) {
          container.style.transform = `scale(${zoomLevel})`;
          container.style.transformOrigin = "top left";
        }
        resizeChartContainer();
  
        // 2. Wait for layout + Update strokes
        requestAnimationFrame(() => {
          updateNodeStrokes();
  
          // 3. Then focus
          requestAnimationFrame(() => {
            waitAndFocusNode(emailId);
          });
        });
      });
    }
  };  

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
          resizeChartContainer();
        }}
        className="clickableLabel"
      >
        Reset
        <IconButton edge="end" size="small" aria-label="Reset" sx={{color:'#2A7386', marginLeft: "4px", padding: "2px"}} onClick={() => {
          setSelectedNameKey("");
          setSelectedNameOption(null);
          setSelectedDirKey("");
          filterChartByReset();
          resizeChartContainer();
        }}>
          <ClearRoundedIcon style={{fontSize:'8'}} />
        </IconButton>
      </Typography>
    );
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

  function filterChartByReset(): void {
     
    if (orgChartElement.current && orgChartData) {
      clearHighlighting();
      setOrgChartData(null);
      setEmployeesForOrgChart(undefined);

    /*  if(!props.data) return;
        populateDirectorateDropDown(props.data);*/

      if (orgChartElement.current) {
        teardownOrgChart();
        $(orgChartElement.current).empty();
      }

      resetZoomAndScroll();
      setIsNameAutocompleteEnabled(false);
    }
  }

  const reloadOrgChart = (optionalNameOption?: IComboBoxOption | null): void => {
    if (!orgChartElement.current || !selectedDirKey?.length) return;
  
    // Step 1: Full cleanup
    teardownOrgChart();
    $(orgChartElement.current).empty();
  
    // Step 2: Get name dropdown options for the selected directorate
    const nameOptionsEmployeeList = populateEmployeeNameDropdown(selectedDirKey);
    if (!nameOptionsEmployeeList?.length) return;
  
    // Step 3: Determine selected employee
    const activeOption = optionalNameOption !== undefined ? optionalNameOption : selectedNameOption;
    const selEmployee = props.data?.find(
      node =>
        node.EmpO365Email?.trim().toLowerCase() === activeOption?.key?.toString().trim().toLowerCase()
    );
  
    setSelectedNodeEmployee(selEmployee);
    setSelectedNameKey(activeOption?.key?.toString() ?? "");
  
    // Step 4: Reset zoom state and defer transform to after render
    zoomInitializedRef.current = false;
    setZoomLevel(initialZoomRef.current);           // update state
    zoomLevelRef.current = initialZoomRef.current;  // update ref
  
    // Step 5: Render the chart (match zoom + center timing)
    renderDirectorateEmployeeOrgChart(nameOptionsEmployeeList, selEmployee);
  
    // Step 6: Defer resize + zoom transform + scroll center until layout stabilizes
    requestAnimationFrame(() => {
      resizeChartContainer();
  
      requestAnimationFrame(() => {
        applyInitialZoomIfNeeded();
  
        // Only center if no specific employee is focused
        if (!selEmployee) {
          requestAnimationFrame(() => {
            centerOrgChart1(initialZoomRef.current);
          });
        }
      });
    });
  };  

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
            <title>Directorate View - ${selectedDirKey || selectedNameOption?.text || ""}</title>
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
             
  useEffect(() => {
    const link = document.createElement("link");
    link.href = "https://fonts.googleapis.com/icon?family=Material+Icons";
    link.rel = "stylesheet";
    document.head.appendChild(link);
  }, []);

  useEffect(() => {
    if (!props.data) return;
  
    populateDirectorateDropDown(props.data);
  
    if (selectedDirectorate?.length) {
      setSelectedDirKey(selectedDirectorate);
      const nameOptionsEmployeeList = populateEmployeeNameDropdown(selectedDirectorate);
  
      if (selectedNodeEmployee) {
        const selectedNodeEmployeeOption: IComboBoxOption = {
          key: selectedNodeEmployee.EmpO365Email ?? "",
          text: selectedNodeEmployee.EmployeeName ?? ""
        };
        setSelectedNameOption(selectedNodeEmployeeOption);
      }
  
      if (nameOptionsEmployeeList?.length) {
        renderDirectorateEmployeeOrgChart(nameOptionsEmployeeList, selectedNodeEmployee);
  
        if (selectedNodeEmployee?.EmpO365Email) {
          requestAnimationFrame(() => {
            waitAndFocusNode(selectedNodeEmployee.EmpO365Email ?? '');
          });
        }
        else {
          centerOrgChart1(zoomLevel);
        }
      }
    }
  
  }, [props.data]);

  useEffect(() => {
    const container = document.getElementById("chart-container");
    if (!container) return;
  
    let zoomTimeout: ReturnType<typeof setTimeout> | null = null;
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
  
      // Reset zoom lock after a short delay
      if (zoomTimeout) clearTimeout(zoomTimeout);
      zoomTimeout = setTimeout(() => {
        isZooming = false;
      }, 100);
    }, 60, { leading: true, trailing: false });
  
    container.addEventListener("wheel", throttledWheel, { passive: false });
  
    return () => {
      container.removeEventListener("wheel", throttledWheel);
      throttledWheel.cancel();
      if (zoomTimeout) clearTimeout(zoomTimeout);
    };
  }, []);
  
  
  useEffect(() => {
    const handleKeyZoom = (e: KeyboardEvent): void => {
      if (!(e.ctrlKey || e.metaKey)) return;
      if (e.key === '+' || e.key === '=' || e.key === '-') {
        e.preventDefault();
        const delta = (e.key === '+' || e.key === '=') ? ZOOM_STEP : -ZOOM_STEP;
        applyZoomChange(delta);
      }
    };
    window.addEventListener("keydown", handleKeyZoom);
    return () => window.removeEventListener("keydown", handleKeyZoom);
  }, []);
  
  /*useEffect(() => {
    zoomLevelRef.current = zoomLevelRef.current;
  }, [zoomLevelRef.current]);  */
  
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
                  maxWidth: 500,
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
                    setIsNameAutocompleteEnabled(false);
                    filterChartByReset();
                  } else {
                    const newDirKey = value?.key as string || "";
                    setSelectedDirKey(newDirKey);
                    setSelectedNameOption(null);
                    setSelectedNameKey("");
                    setIsNameAutocompleteEnabled(true);
                    directorateOptionChange(newDirKey);
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
                disabled={!isNameAutocompleteEnabled}
                value={selectedNameOption ? nameOptions.find((d) => d.key === selectedNameOption?.key) : null}
                filterOptions={(options, state) =>
                  options.filter((option) => option.text.toLowerCase().includes(state.inputValue.toLowerCase()))
                }
                onChange={(event, value, reason) => {
                  if (reason === "clear") { 
                    setSelectedNameOption(null);
                    setSelectedNameKey("");
                    requestAnimationFrame(() => {
                      reloadOrgChart(null);
                    });
                  } else {
                    setSelectedNameKey(value?.key as string || "");
                    setSelectedNameOption(value);
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
                    requestAnimationFrame(() => {
                      resetZoom();
                    });
                  }}
                  aria-label="Reset Zoom"
                >
                  <CenterFocusStrong style={{ fontSize: "20px" }} />
                </IconButton>
              </Tooltip>

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
              width: 'auto',
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
            }}
          >
            <div id="chart-inner"
               style={{
                width: '100%',
                height: '100%',
                minWidth: '100%',
                minHeight: '100%',
                display: "inline-block",
                cursor: "grab",
                position: "relative",
              }}
            >
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
}
