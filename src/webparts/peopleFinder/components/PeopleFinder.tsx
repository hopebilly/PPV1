import * as React from 'react';
import { ServiceScope } from "@microsoft/sp-core-library";
//import { useState } from "react";
import { SharedProvider } from '../../../contexts/SharedContext';
import { useSharedContext } from "../../../contexts/SharedContext";
import EmployeeSearchComponent from '../../employeeSearch/components/EmployeeSearchComponent';
//import PeopleorgchartComponent from '../../peopleorgchart/components/PeopleorgchartComponent';
import OrganizationChartComponent from '../../organizationChart/components/OrganizationChartComponent';
import { Tabs, Tab, Box, styled} from '@mui/material';
import AccountCircleIcon from '@material-ui/icons/AccountCircle';
import { PeopleTeam20Filled } from '@fluentui/react-icons';

export interface IPeopleFinderProps {
  serviceScope: ServiceScope;
}

const StyledTab = styled(Tab)({
  textTransform: 'none',
  display: 'block',
  fontSize: '0.8rem',
  marginBlockEnd: '15px',
  marginInlineStart: '0px',
  marginInlineEnd: '0px',
  fontWeight: 'bold',
  fontFamily: '"Segoe UI", sans-serif',
  color: '#2A7386'
});

const PeopleFinderController: React.FC<IPeopleFinderProps> = (props) => {

  const {selectedTab, setSelectedTab, setEmployeesForOrgChart, setSelectedNameKey, setSelectedNodeEmployee,
    setSelectedOrgChartView, 
  } = useSharedContext();

  const handleTabChange = (event: React.SyntheticEvent, newValue: number): void => {
    setSelectedTab(newValue);
    if(newValue === 0)
      setEmployeesForOrgChart(undefined);
      setSelectedNameKey('');
      setSelectedNodeEmployee(undefined);
      setSelectedOrgChartView("Button1");
  };

  return (
    
    <Box sx={{ width: '100%' }}>
    {/* Material-UI Tabs */}
    <Tabs
      value={selectedTab}
      onChange={handleTabChange}
      aria-label="tabbed navigation"
      indicatorColor="secondary"
      textColor="inherit"
      sx={{
        width: '100%',
        borderBottom: '2px solid #E0E0E0', // Light grey bottom border for tabs container
        borderRadius: '8px', // Curved edges for the tabs container
        overflow: 'hidden', // Prevents the indicator from going outside the container's edges
        backgroundColor: 'transparent', // Set custom background color for the tabs container
        height: '50px',
        minHeight: '50px',
        '& .MuiTabs-indicator': {
            height: '4px',
            backgroundColor: 'transparent',
        },
        display: 'flex',// Display tabs in a row (horizontal layout)
      }}
    >
      <StyledTab label="Employee Search"
        icon={<AccountCircleIcon />}
        iconPosition="start"
         sx={{
          width: '40%',
          borderRight: '2px solid #E0E0E0',
          backgroundColor: selectedTab === 0 ? '#F1F1F1' : 'transparent',
          borderTopLeftRadius: '8px',
          borderTopRightRadius: '8px',
          //color: selectedTab === 0 ? 'white' : '#e0e0e0', // Change text color for active tab
          color: '#2A7386', // text color 
          '&:hover': {
            backgroundColor: selectedTab === 0 ? '#f5f5f5' : 'transparent', // Slightly darker on hover for active tab
          },
          display: 'flex', // Use flexbox to center the content
          justifyContent: 'flex-start', // Start content horizontally
          alignItems: 'center', // Center content vertically
          textAlign: 'left', // Ensure label is centered
          height: '50px',
          minHeight: '50px'
        }}
      />
      <StyledTab label="Organization Chart"
        icon={ <PeopleTeam20Filled />}
        iconPosition="start"
         sx={{
          width: '40%',
          borderRight: '2px solid #E0E0E0', // Light grey border between tabs
          backgroundColor: selectedTab === 0 ? '#F1F1F1' : 'transparent',
          borderTopLeftRadius: '8px',
          borderTopRightRadius: '8px',
          //color: selectedTab === 1 ? 'white' : '#e0e0e0', // Change text color for active tab
          color: '#2A7386', // Change text color for active tab
          '&:hover': {
            backgroundColor: selectedTab === 1 ? '#f5f5f5' : 'transparent', // Slightly darker on hover for active tab
          },
          display: 'flex', // Use flexbox to center the content
          justifyContent: 'flex-start', // Left content horizontally
          alignItems: 'center', // Center content vertically
          textAlign: 'left', // Ensure label is centered
          height: '50px',
          minHeight: '50px'
        }} 
      />
    </Tabs>

    {/* Tab Content */}
    <Box sx={{ padding: '20px', backgroundColor: '#f5f5f5' }}>
      {selectedTab === 0 && <EmployeeSearchComponent serviceScope={props.serviceScope}/>}
      {selectedTab === 1 && <OrganizationChartComponent serviceScope={props.serviceScope} />}
    </Box>
  </Box>
    
  );
};

export const PeopleFinder: React.FC<IPeopleFinderProps> = (props) => {
  return (
    <SharedProvider>
      <PeopleFinderController serviceScope={props.serviceScope} />
    </SharedProvider>
  );
}
