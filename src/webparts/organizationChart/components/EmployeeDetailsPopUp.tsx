require("./DirectorateOrgChart.css");
import { ServiceScope } from "@microsoft/sp-core-library";
import { FC } from "react";
import { useSharedContext } from "../../../contexts/SharedContext";
import { IconButton, Typography, Modal, Grid, Box, Button} from '@mui/material';
import { CalendarLtr20Regular, Mail20Regular, Chat20Regular, BuildingSkyscraper20Regular, 
    PeopleCommunity20Regular, People20Regular, Location20Regular, Call20Regular} from '@fluentui/react-icons';
import ClearRoundedIcon from '@material-ui/icons/ClearRounded';

interface IEmployeeDetailsPopUpprops {
    serviceScope: ServiceScope;
}

export const EmployeeDetailsPopUp: FC<IEmployeeDetailsPopUpprops> = (props) => {

    const {setSelectedNodeEmployee, setSelectedOrgChartView, setSelectedDirectorate, 
        selectedEmployee, setSelectedEmployee, isEmployeeModalOpen, setEmployeeModalOpen} = useSharedContext();

      // Split by the comma separating the last name and first/middle name
   // If firstName has middle names or initials, we can directly return it as the formatted string
   const parseName = (name: string): string => {
    
    const [lastName, firstName] = name.split(",").map(part => part.trim());
    return `${firstName} ${lastName}`;
    };

    // Remove all non-numeric characters
    // Return unformatted phone if it doesn't match the pattern
    const formatPhoneNumber = (phone: string): string => {
        const cleaned = ('' + phone).replace(/\D/g, '');
        const match = cleaned.match(/^(\d{3})(\d{3})(\d{4})$/);
        if (match) {
        return `(${match[1]}) ${match[2]}-${match[3]}`;
        }
        return phone;
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

    //Modal for Employee Details on click of Org chart employee node

    return (
        
      <Modal open={isEmployeeModalOpen} onClose={handleEmployeeNameCloseModal} aria-labelledby="employee-details-modal"
            aria-describedby="modal-to-display-employee-details" disableRestoreFocus>
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
        {selectedEmployee && (
          <img
            src={`https://nih.sharepoint.com/_layouts/15/userphoto.aspx?size=L&username=${selectedEmployee.EmpO365Email}`}
            alt={`${selectedEmployee.EmployeeName} Photo`}
            style={{
              position: 'absolute',
              top: '-50px',
              left: '20px',
              borderRadius: '50%',
              width: '70px',
              height: '70px',
              border: '3px solid #2A7386',
              zIndex: 1,
            }}
          />
        )}
        {selectedEmployee && selectedEmployee.DirectorateHead && (
            <Button
            onClick={() => {
              setSelectedOrgChartView("Button2");
              if (selectedEmployee){
                setSelectedDirectorate(selectedEmployee.Directorate ?? '');
                setSelectedNodeEmployee(selectedEmployee);
              }
            }}
            sx={{
              position: 'absolute',
              top: '44px', // Align with the employee name
              right: '20px', // Keep towards the right edge
              backgroundColor: '#f5f5f5', // Light gray background
              color: '#2A7386', // Text color matches the typography
              fontSize: '0.8rem', // Font size from Typography
              fontWeight: 'bold', // Bold text from Typography
              borderRadius: '8px', // Rounded corners for a sleek look
              padding: '8px 16px', // Padding for the button
              boxShadow: 'none', // No box shadow for a flat design
              '&:hover': {
                //backgroundColor: '#e0e0e0', // Slightly darker on hover
                backgroundColor: '#E2E2E2',
              },
            }}
          >
            Organization Chart
          </Button>
        )}
      
        {/* Close Button */}
        <IconButton
          onClick={handleEmployeeNameCloseModal}
          sx={{
            position: 'absolute',
            top: 8,
            right: 8,
            color: 'gray',
            zIndex: 2,
          }}
        >
          <ClearRoundedIcon style={{ fontSize: 18 }} />
        </IconButton>

        {selectedEmployee && (
          <>
            <Typography
              variant="h6"
              sx={{
                mb: 0.5,
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
                mb: 2,
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
                    {selectedEmployee.SupervisorName? parseName(selectedEmployee.SupervisorName): 'Not Available'}
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
    );
}