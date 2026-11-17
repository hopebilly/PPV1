import * as React from "react";
import styles from './EmployeeSearch.module.scss';
import { useState } from "react";
import { useSharedContext } from "../../../contexts/SharedContext";
import { Table, TableBody, TableCell, TableContainer, TableHead, TableRow, Paper, Tooltip, InputAdornment, makeStyles} from "@material-ui/core";
import { Button } from '@mui/material';
import { ArrowDownward, ArrowUpward } from "@material-ui/icons";
import { ServiceScope } from "@microsoft/sp-core-library";
import { IEmployee } from "../../../model";
import { Icon} from '@fluentui/react/lib/Icon';
import ClearRoundedIcon from '@material-ui/icons/ClearRounded';
import { Pagination, FormControl, Select, MenuItem, SelectChangeEvent, Box, IconButton, TextField, Modal, 
  Typography, Grid} from '@mui/material';
import Divider, { dividerClasses } from '@mui/material/Divider';
import { Panel, PanelType } from "@fluentui/react";
import { CoreOrgChart } from "../../peopleorgchart/components/CoreOrgChart";
import { CalendarLtr20Regular, Mail20Regular, Organization20Filled, Chat20Regular, BuildingSkyscraper20Regular, PeopleCommunity20Regular,
  People20Regular, Location20Regular, Call20Regular
 } from '@fluentui/react-icons';
import Excel from "exceljs";
import { saveAs } from "file-saver";
import jsPDF from "jspdf";
import  autoTable from 'jspdf-autotable';

//import AccountCircleIcon from '@material-ui/icons/AccountCircle';

export interface IEmployeeSearchProps {
  serviceScope: ServiceScope;
  employees: IEmployee[];
}

interface fillPattern {
  type: "pattern", 
  pattern: Excel.FillPattern['pattern'],
  fgColor: {argb: string}
}

const columns = [
    { Header: "", accessor: "OrgChart" },
    { Header: "Name", accessor: "EmployeeName" },
    { Header: "Job Title", accessor: "JobTitle" },
    { Header: "Directorate", accessor: "DirectorateName" },
    { Header: "Sub Group", accessor: "Department" },
    { Header: "Email", accessor: "EmpO365Email" },
    { Header: "Location", accessor: "City" }
];

const useStyles = makeStyles((theme) => ({
  tableHeaderCell: {
    wordBreak: 'break-word',
    whiteSpace: 'normal',
    maxWidth: 200,
    fontSize: '11px',
    padding: '8px',
    '&:hover': {
      backgroundColor: '#f5f5f5',
      color: '#2A7386',
      cursor: 'pointer',
    },
    '&:active': {
     backgroundColor: '#e0e0e0',
      color: '#2A7386',
    }
  },
  filterTextField: {
    marginRight: '8px',
    underline: 'bottom-border',
    '& .MuiInputBase-input': {
      color: '#17414c',
      fontSize: '14px',
    },
    backgroundColor: '#fff',
    '& .MuiInputBase-root': {
      backgroundColor: '#fff',
    },
    '& .MuiInput-underline:before': {
      borderBottomColor: '#2A7386', // Initial (inactive) underline color
    },
    '& .MuiInput-underline:hover:before': {
      borderBottomColor: '#17414c', // Underline color when hovering
    },
    '& .MuiInput-underline:after': {
      borderBottomColor: '#17414c', // Underline color when focused
    },
    '& .MuiOutlinedInput-root': {
      paddingRight: '36px',
    },
  },
  resetButton: {
    marginLeft: 'auto',
    border: 'none',
    fontWeight: 'bold',
    borderRadius: '10%',
    backgroundColor: '#fff',
    color: '#2A7386' 
  },
  clickableLabel: {
    cursor: 'pointer',
    padding: '4px 8px',
    border: '1px solid #17414c',
    borderRadius: '8px',
    backgroundColor: '#fff',
    color: '#2A7386',
    display: 'inline-block',
    fontSize: '10px',
    textAlign: 'center',
    boxShadow: '0 2px 4px rgba(0, 0, 0, 0.1)',
    '&:hover': {
      backgroundColor: '#f3f3f2',
      boxShadow: '0 4px 8px rgba(0, 0, 0, 0.1), 0 2px 4px rgba(0, 0, 0, 0.2)',
    },
    '&:active': {
      backgroundColor: '#d0d0d0',
      boxShadow: 'none',
    },
  },
  colHeader: {
    display: 'flex', 
    alignItems: 'center',
    color: '#2A7386', 
    fontSize: '14px', 
    fontWeight: 'bold' 
  }
}));

export const FilterableTable: React.FC<IEmployeeSearchProps> = (props) => {

  const classes = useStyles();
  const [filters, setFilters] = useState({
    name: '',
    jobTitle: '',
    directorate: '',
    subgroup: '',
    email: '',
    location: ''
  });
  const [sortBy, setSortBy] = useState<string | null>("");
  const [sortDirection, setSortDirection] = useState<"asc" | "desc">("asc");
  const [page, setPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(10);
  const [selectedEmployee, setSelectedEmployee] = React.useState<IEmployee | undefined>();
  const [isOpen, setIsOpen] = React.useState(false);
  const [isEmployeeModalOpen, setIsEmployeeModalOpen] = useState(false);
  //const orgChartEmployees : IEmployee[] = [];
  const {setSelectedTab, employeesForOrgChart, 
                          //setEmployeesForOrgChart, 
                          setSelectedNodeEmployee, setSelectedOrgChartView, setSelectedDirectorate} = useSharedContext();
  
  const isEmptyMatch = (value: string | null | undefined, filter: string): boolean => {
    const normalizedFilter = filter.trim().toLowerCase();
  
    // No filter applied
    if (!normalizedFilter) return true;
  
    // User wants to see only empty/null/undefined values
    if (normalizedFilter === "#empty") {
      return value === null || value === undefined || value.trim() === "";
    }
  
    // Normal substring match
    return value?.toLowerCase().includes(normalizedFilter) ?? false;
  };

  const filteredData = React.useMemo(() => {
    return props.employees.filter((row) => {
      return (
        isEmptyMatch(row.EmployeeName, filters.name) &&
        isEmptyMatch(row.JobTitle, filters.jobTitle) &&
        isEmptyMatch(row.DirectorateName, filters.directorate) &&
        isEmptyMatch(row.Department, filters.subgroup) &&
        isEmptyMatch(row.EmpO365Email, filters.email) &&
        isEmptyMatch(row.City, filters.location)
      );
    });
  }, [props.employees, filters]);                         
                          
  // Sort the filtered data
  const sortedData = filteredData.sort((a, b) => {
  
    let valueA : string | number | boolean | Date | IEmployee[] | undefined;
    let valueB : string | number | boolean | Date | IEmployee[] | undefined;

    if (sortBy === 'EmployeeName') {
      valueA = (a.EmployeeName !== null && a.EmployeeName !== undefined)?(a.EmployeeName.toString().trim().split(' ').pop()): '';
      valueB = (b.EmployeeName !== null && b.EmployeeName !== undefined)?(b.EmployeeName.toString().trim().split(' ').pop()): '';
    }
    else if(sortBy === 'City')
    {
        valueA = (a.City !== null && a.City !== undefined)? a.City.toString().trim().toLowerCase(): " ";
        valueB = (b.City !== null && b.City !== undefined)? b.City.toString().trim().toLowerCase():" ";
    }
    else {
        valueA = a[sortBy as keyof IEmployee];
        valueB = b[sortBy as keyof IEmployee];

        if (typeof valueA === "string") {
          valueA = valueA.toLowerCase();
        }

        if (typeof valueB === "string") {
          valueB = valueB.toString().toLowerCase();
        }
    }

      valueA = (valueA !== null && valueA !== undefined)? valueA : '';
      valueB = (valueB !== null && valueB !== undefined)? valueB : '';

      if (sortDirection === "asc") {
        return valueA < valueB ? -1 : valueA > valueB ? 1 : 0;
      } else {
        return valueA < valueB ? 1 : valueA > valueB ? -1 : 0;
      }
    }
  );

  // Handle sorting when a column header is clicked
  const handleSort = (column: string): void => {
    if (sortBy === column) {
      setSortDirection((prevSortDirection) =>
        prevSortDirection === "asc" ? "desc" : "asc"
      );
    } else {
      setSortBy((prevSortBy) => {
        if (prevSortBy === column) {
          return null; 
        } else {
          return column;
        }
    });
      setSortDirection((prevSortDirection) =>
        prevSortDirection === "asc" ? "desc" : "asc"
      );
    
    }
  };

  // Handle page change
  const handleChangePage = (event: React.ChangeEvent<unknown> | null, newPage: number): void => {
    setPage(newPage);
    setRowsPerPage(rowsPerPage);
  };

  // Handle rows per page change 
    const handleChangeRowsPerPage = (event: SelectChangeEvent<number>): void => {
      setRowsPerPage(Number(event.target.value));
      setPage(1);
  };

  //Reset all Filters
  const handleResetFilters = (): void => {
    setFilters({
      name: '',
      jobTitle: '',
      directorate: '',
      subgroup: '',
      email: '',
      location: ''
    });    
  };

  const ResetFiltersLabel = (): JSX.Element  => {
    
    return (
      <Typography
        variant="body2"
        component="span"
        onClick={handleResetFilters}
        className={classes.clickableLabel} 
      >
        Reset Filters
        <IconButton edge="end" size="small" sx={{color:'#2A7386'}} onClick={handleResetFilters}>
          <ClearRoundedIcon style={{fontSize:'8'}} />
        </IconButton>
      </Typography>
    );
  };

  // exports filtered data to excel
  const exportToExcel = async(): Promise<void> => {
    
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet('EmployeeData');
    
    worksheet.columns = [
      {header: "Employee Name", key: "EmployeeName", width:30},
      {header: "Job Title", key: "JobTitle", width:30},
      {header: "Directorate", key: "DirectorateName", width:30},
      {header: "Department", key: "Department", width:30},
      {header: "Email", key: "EmpO365Email", width:30},
      {header: "City", key: "City", width:30},
      
    ]
    worksheet.getRow(1).font = {bold:true}
    
    filteredData.forEach((item,index) => {
      const row = worksheet.addRow({
        EmployeeName: item.EmployeeName, 
        JobTitle: item.JobTitle,
        DirectorateName: item.DirectorateName,
        Department: item.Department,
        EmpO365Email: item.EmpO365Email,
        City: item.City,

    });
    const fill: fillPattern = {
      type:"pattern", pattern:"solid",
      fgColor:{
        argb:index % 2 === 0 ? "f1f0f2" : "FFFFFFFF"
      }
    };
    row.eachCell({includeEmpty: true}, (cell, colNumber) =>{
      if (colNumber <= 6){
        cell.fill = fill
      }
    })
    });  
  
    const excelBuffer = await workbook.xlsx.writeBuffer();
    const data = new Blob([excelBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    saveAs(data, "filteredemployee.xlsx");
  };

  // Export to PDF 
  const exportToPDF = async(): Promise<void> => {

    if (filteredData.length===0){
      console.log ('No data to export');
      return;
    }
    const doc = new jsPDF();
    doc.text('Filtered Employee Data', 14,10);
    const tableColumn =  [
      "Employee Name",
      "Job Title",
      "Directorate",
      "Department",
      "Email",
      "City",
    ];

    const tableRows: string[][] = [];
    filteredData.forEach(item=>{
      tableRows.push([item.EmployeeName?.toString()|| 'N/A', item.JobTitle?.toString()|| 'N/A', item.Directorate?.toString()|| 'N/A', item.Department?.toString()|| 'N/A', item.EmpO365Email?.toString()|| 'N/A', item.City?.toString()|| 'N/A'])
    });
    autoTable(doc,{
      head:[tableColumn],
      body: tableRows,
      startY: 20,

    });
    doc.save('Filtered_Employee_data.pdf')
  }

 /* const getOrgChartEmployees = (selectedEmployee: IEmployee, rootOrgNode: IEmployee): void => {

    if(selectedEmployee.EmpO365Email !== rootOrgNode?.EmpO365Email)
    {
      orgChartEmployees.push(selectedEmployee); 
      const supervisorEmployee: IEmployee | undefined = props.employees?.find(e => e.EmpO365Email === selectedEmployee.SupervisorO365Email);
      if(supervisorEmployee){
        getOrgChartEmployees(supervisorEmployee, rootOrgNode);
      }
    }
    setEmployeesForOrgChart(orgChartEmployees);
  }*/

  // On Click of icon open Org Chart Panel
  const handleOrgChartClick = (employee: IEmployee): void => {

    setSelectedEmployee(employee);
    
    if(employee){
      setSelectedNodeEmployee(employee);
    }

    if(employee.Directorate){
      setSelectedDirectorate(employee.Directorate);
    }

    setSelectedTab(1);

    if((employee && 
      (employee?.KeyStaff?.trim().toLowerCase() === "yes".trim() ||
       employee?.DirectorateHead?.trim().toLowerCase() === "yes".trim()))){
       setSelectedOrgChartView("Button1");
    }
    else {
        setSelectedOrgChartView("Button2");
    }
  }

  const handleEmployeeNameClick = (employee: IEmployee): void => {
    setSelectedEmployee(employee);
    setIsEmployeeModalOpen(true);
  };

  const handleEmployeeNameCloseModal = (): void => {
    setSelectedEmployee(undefined);
    setIsEmployeeModalOpen(false);
  };

    // Split by the comma separating the last name and first/middle name
   // If firstName has middle names or initials, we can directly return it as the formatted string
   const parseName = (name: string): string => {
    
    const [lastName, firstName] = name.split(",").map(part => part.trim());
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

  // Slice rows based on the current page
  const startIndex = (page - 1) * rowsPerPage;
  const paginatedData = sortedData.slice(startIndex, startIndex + rowsPerPage);

  const totalPages = Math.ceil(sortedData.length / rowsPerPage);

  // Display range of records
  const startRecord = startIndex + 1;
  const endRecord = startIndex + paginatedData.length;

  /*console.log("Filtered:", filteredData.length, filteredData);
  console.log("Sorted:", sortedData.length, sortedData);
  console.log("Paginated:", paginatedData.length, paginatedData);

  const ids = paginatedData.map(emp => emp.id);
  const uniqueIds = new Set(ids);
  console.log("Unique IDs in page:", uniqueIds.size, "vs Total IDs:", ids.length);*/

  React.useEffect(() => {
    const maxPage = Math.ceil(filteredData.length / rowsPerPage);
    if (page > maxPage && maxPage > 0) {
      setPage(maxPage); // or setPage(1)
    }
  }, [filteredData, page, rowsPerPage]);  

  return (
  <div>
    
    <TableContainer component={Paper} style={{width:"100%", maxWidth: "100%", overflowX: "auto", margin: '5 0 px auto' }}>
      <Box
        sx={{
          display: 'flex',
          flexDirection: 'row', // Stack the content vertically
          alignItems: 'flex-end', // Align items to the right
          justifyContent: 'flex-end',
          height: 'auto', // Adjust height to fit content
          border: '1px solid rgba(0, 0, 0, 0.12)',
          borderColor: 'divider',
          borderRadius: 1,
          bgcolor: 'background.paper',
          color: 'text.secondary',
          margin: '5px 5px 0px 0px',
        }}
      >
          {/* Box for Export to Excel and Export to PDF Buttons */}
         
            <Typography variant="body2" style={{ marginRight: '10px', fontSize: '0.8rem', lineHeight: '3.0' }}>
              Export
            </Typography>
            {/* Export to Excel Button */}
            <Tooltip title="Export to Excel">
              <IconButton
                color="primary"
                onClick={exportToExcel}
                aria-label="Download to Excel"
                sx={{ marginRight: '10px' }}
              >
                <Icon
                  iconName="ExcelLogo"
                  style={{ cursor: "pointer", fontSize: 20, color: "#2A7386" }}
                />
              </IconButton>
            </Tooltip>
            <Divider orientation="vertical" flexItem />
            {/* Export to PDF Button */}
            <Tooltip title="Export to PDF">
              <IconButton
                color="primary"
                onClick={exportToPDF}
                aria-label="Download to PDF"
                sx={{ marginRight: '10px' }}
              >
                <Icon
                  iconName="PDF"
                  style={{ cursor: "pointer", fontSize: 20, color: "#2A7386" }}
                />
              </IconButton>
            </Tooltip>
       
      </Box>

      <Box
      sx={{
        display: 'flex',
        flexDirection: 'row', // Stack the content vertically
        alignItems: 'center',        // vertical alignment
        justifyContent: 'flex-end',  // aligns everything to the right
        flexWrap: 'wrap',            // optional for small screens
        gap: 1.5,                    // spacing between elements
        padding: '12px 12px 8px 12px',  // top padding + side padding
        marginBottom: '8px',           // subtle space below the box
        height: 'auto', // Adjust height to fit content
        border: '1px solid',
        borderColor: 'divider',
        borderRadius: 1,
        bgcolor: 'background.paper',
        color: 'text.secondary',
        margin: '0px 5px 0px 0px',
      }}
    >
      {(Object.values(filters).some(f => f.trim().length > 0)) && (
      <>
        <ResetFiltersLabel />
        <Divider orientation="vertical" flexItem sx={{ mx: 1 }} />
      </>
    )}

    <Typography variant="body2" sx={{ fontSize: '0.8rem', lineHeight: '1.8'}}>
      Rows per page:
    </Typography>

    <FormControl sx={{ minWidth: 40 }}>
      <Select
        value={rowsPerPage}
        onChange={handleChangeRowsPerPage}
        size="small"
        displayEmpty
        MenuProps={{
          PaperProps: {
            sx: {
              width: 70,
              fontSize: '0.75rem',
              '& .MuiMenuItem-root': {
                fontSize: '0.75rem',
                '&:hover': { backgroundColor: '#e0e0e0' },
              },
              '& .MuiMenuItem-root.Mui-selected': {
                fontSize: '0.75rem',
                backgroundColor: '#d3d3d3',
                '&:hover': { backgroundColor: '#b0b0b0' },
              },
            },
          },
        }}
        sx={{
          fontSize: '0.75rem',
          padding: '0 2px',
          border: '1px solid #2A7386',
          borderRadius: '4px',
          height: '28px',
          '& .MuiSelect-select': {
            padding: '0 2px',
          },
          '& .MuiSelect-icon': {
            top: '30%',
            right: '4px',
            transform: 'translateY(-30%)',
          },
          '& .MuiOutlinedInput-notchedOutline': {
            borderWidth: 0,
          },
          '&.Mui-focused .MuiOutlinedInput-notchedOutline': {
            borderColor: 'transparent',
          },
        }}
        inputProps={{ 'aria-label': 'Rows per page' }}
      >
        <MenuItem value={5}>5</MenuItem>
        <MenuItem value={10}>10</MenuItem>
        <MenuItem value={25}>25</MenuItem>
        <MenuItem value={100}>100</MenuItem>
        <MenuItem value={-1}>Show All</MenuItem>
      </Select>
    </FormControl>

    <Typography variant="body2" sx={{ fontSize: '0.75rem' }}>
      {`${startRecord}-${endRecord} of ${sortedData.length}`}
    </Typography>

      
  </Box>

    <Table style={{ minWidth: "100%", tableLayout: 'auto' }}>
      <TableHead>
        <TableRow>

          <TableCell style={{width:'10%'}} className={classes.tableHeaderCell}>
            <div style={{ display: 'flex', alignItems: 'center' }} />
        
          </TableCell>
        
          <TableCell style={{width:'15%'}} className={classes.tableHeaderCell}>
            <div style={{ display: 'flex', alignItems: 'center' }}>
              <Tooltip title="Name">
                <TextField
                  value={filters.name}
                  onChange={(e) => {
                    setFilters(prev => ({ ...prev, name: e.target.value }))
                  }}
                  placeholder="Name"
                  variant="standard"
                  size="small"
                  fullWidth
                  className={classes.filterTextField}
                  InputProps={{
                    startAdornment: (
                      <InputAdornment position="start">
                        <Icon iconName="Filter" 
                              style={
                                      { fontSize: '15px', 
                                      color: '#2A7386', 
                                      borderRadius: '50%', 
                                      padding: '4px',  
                                      }
                                    } />
                      </InputAdornment>
                    ),
                    endAdornment: filters.name && (
                      <InputAdornment position="end">
                        <IconButton edge="end" size="small" sx={{color:'#2A7386', padding: 0, marginRight: '4px'}} onClick={() => setFilters(prev => ({ ...prev, name: '' }))}>
                          <ClearRoundedIcon style={{fontSize:'12'}} />
                        </IconButton>
                      </InputAdornment>
                    ),
                  }}
                />
              </Tooltip>
              <Tooltip title="Sort">
                <IconButton size="small" style={{ display: 'inline' }} onClick={() => handleSort(columns.filter((column) => column.accessor === 'EmployeeName')[0].accessor)}> 
                  {sortBy === (columns.filter((column) => column.accessor === 'EmployeeName')[0].accessor) && sortDirection === "desc" ? <ArrowDownward fontSize="small" /> : <ArrowUpward fontSize="small" />}
                </IconButton>
              </Tooltip>
            </div>
          </TableCell>

          <TableCell style={{width:'15%'}} className={classes.tableHeaderCell}>
            <div style={{ display: 'flex', alignItems: 'center' }}>
              <Tooltip title="Job Title (type #empty for blanks)">
                <TextField
                  value={filters.jobTitle}
                  onChange={(e) => {
                    setFilters(prev => ({ ...prev, jobTitle: e.target.value }))
                  }}
                  placeholder="Job Title"
                  variant="standard"
                  size="small"
                  fullWidth
                  className={classes.filterTextField} 
                  InputProps={{
                    startAdornment: (
                      <InputAdornment position="start">
                        <Icon iconName="Filter" 
                              style={
                                { fontSize: '15px', 
                                color: '#2A7386', 
                                borderRadius: '50%', 
                                padding: '4px',  
                                }
                              } />
                      </InputAdornment>
                    ),
                    endAdornment: filters.jobTitle && (
                      <InputAdornment position="end">
                        <IconButton edge="end" size="small" sx={{color:'#2A7386', padding: 0, marginRight: '4px'}}  onClick={() => setFilters(prev => ({ ...prev, jobTitle: '' }))}>
                          <ClearRoundedIcon  style={{fontSize:'12'}} />
                        </IconButton>
                      </InputAdornment>
                    ),
                  }}
                />
              </Tooltip>
              <Tooltip title="Sort">
                <IconButton size="small" onClick={() => handleSort(columns.filter((column) => column.accessor === 'JobTitle')[0].accessor)}>
                  {sortBy === (columns.filter((column) => column.accessor === 'JobTitle')[0].accessor) && sortDirection === "desc" ? <ArrowDownward fontSize="small" /> : <ArrowUpward fontSize="small" />}
                </IconButton>
              </Tooltip>
            </div>
          </TableCell>
  
          <TableCell style={{width:'15%'}} className={classes.tableHeaderCell}>
            <div style={{ display: 'flex', alignItems: 'center' }}>
              <Tooltip title="Directorate (type #empty for blanks)">
                <TextField
                  value={filters.directorate}
                  onChange={(e) => {
                    setFilters(prev => ({ ...prev, directorate: e.target.value }))
                  }}
                  placeholder="Directorate"
                  variant="standard"
                  size="small"
                  fullWidth
                  className={classes.filterTextField}
                  InputProps={{
                    startAdornment: (
                      <InputAdornment position="start">
                        <Icon iconName="Filter" 
                              style={
                                { fontSize: '15px', 
                                color: '#2A7386', 
                                borderRadius: '50%', 
                                padding: '4px',  
                                }
                              } />
                      </InputAdornment>
                    ),
                    endAdornment: filters.directorate && (
                      <InputAdornment position="end">
                        <IconButton edge="end" size="small" sx={{color:'#2A7386', padding: 0, marginRight: '4px'}} onClick={() => setFilters(prev => ({ ...prev, directorate: '' }))}>
                          <ClearRoundedIcon  style={{fontSize:'12'}} />
                        </IconButton>
                      </InputAdornment>
                    ),
                  }}
                />
              </Tooltip>
              <Tooltip title="Sort">
                <IconButton size="small" onClick={() => handleSort(columns.filter((column) => column.accessor === 'DirectorateName')[0].accessor)}>
                  {sortBy === (columns.filter((column) => column.accessor === 'DirectorateName')[0].accessor) && sortDirection === "desc" ? <ArrowDownward fontSize="small" /> : <ArrowUpward fontSize="small" />}
                </IconButton>
              </Tooltip>
            </div>
          </TableCell>
          
          <TableCell style={{width:'15%'}} className={classes.tableHeaderCell}>
          <div style={{ display: 'flex', alignItems: 'center' }}>
              <Tooltip title="Sub Group (type #empty for blanks)">
                <TextField
                    value={filters.subgroup}
                    onChange={(e) => {
                      setFilters(prev => ({ ...prev, subgroup: e.target.value }))
                    }}
                    placeholder="Sub Group"
                    variant="standard"
                    size="small"
                    fullWidth
                    className={classes.filterTextField}
                    InputProps={{
                      startAdornment: (
                        <InputAdornment position="start">
                          <Icon iconName="Filter" 
                                style={
                                  { fontSize: '15px', 
                                  color: '#2A7386', 
                                  borderRadius: '50%', 
                                  padding: '4px',  
                                  }
                                } />
                        </InputAdornment>
                      ),
                      endAdornment: filters.subgroup && (
                        <InputAdornment position="end">
                          <IconButton edge="end" size="small" sx={{color:'#2A7386', padding: 0, marginRight: '4px'}} onClick={() => setFilters(prev => ({ ...prev, subgroup: '' }))}>
                            <ClearRoundedIcon  style={{fontSize:'12'}} />
                          </IconButton>
                        </InputAdornment>
                      ),
                    }}
                />
              </Tooltip>
              <Tooltip title="Sort">
                <IconButton size="small" onClick={() => handleSort(columns.filter((column) => column.accessor === 'Department')[0].accessor)}>
                  {sortBy === (columns.filter((column) => column.accessor === 'Department')[0].accessor) && sortDirection === "desc" ? <ArrowDownward fontSize="small" /> : <ArrowUpward fontSize="small" />}
                </IconButton>
              </Tooltip>
            </div>
          </TableCell>

          <TableCell style={{width:'15%'}} className={classes.tableHeaderCell}>
            <div style={{ display: 'flex', alignItems: 'center' }}>
              <Tooltip title="Email (type #empty for blanks)">
                <TextField
                  value={filters.email}
                  onChange={(e) => {
                    setFilters(prev => ({ ...prev, email: e.target.value }))
                  }}
                  placeholder="Email"
                  variant="standard"
                  size="small"
                  fullWidth
                  className={classes.filterTextField}
                  InputProps={{
                    startAdornment: (
                      <InputAdornment position="start">
                        <Icon iconName="Filter" 
                              style={
                                { fontSize: '15px', 
                                color: '#2A7386', 
                                borderRadius: '50%', 
                                padding: '4px',  
                                }
                              } />
                      </InputAdornment>
                    ),
                    endAdornment: filters.email && (
                      <InputAdornment position="end">
                        <IconButton edge="end" size="small" sx={{color:'#2A7386', padding: 0, marginRight: '4px'}} onClick={() => setFilters(prev => ({ ...prev, email: '' }))}>
                          <ClearRoundedIcon  style={{fontSize:'12'}} />
                        </IconButton>
                      </InputAdornment>
                    ),
                  }}
                />
              </Tooltip>
              <Tooltip title="Sort">
                <IconButton size="small" onClick={() => handleSort(columns.filter((column) => column.accessor === 'EmpO365Email')[0].accessor)}>
                  {sortBy === (columns.filter((column) => column.accessor === 'EmpO365Email')[0].accessor) && sortDirection === "desc" ? <ArrowDownward fontSize="small" /> : <ArrowUpward fontSize="small" />}
                </IconButton>
              </Tooltip>
            </div>
          </TableCell>

          <TableCell style={{width:'15%'}} className={classes.tableHeaderCell}>
            <div style={{ display: 'flex', alignItems: 'center' }}>
              <Tooltip title="Location (type #empty for blanks)">
                <TextField
                  value={filters.location}
                  onChange={(e) => {
                    setFilters(prev => ({ ...prev, location: e.target.value }))
                  }}
                  placeholder="Location"
                  variant="standard"
                  size="small"
                  fullWidth
                  className={classes.filterTextField}
                  InputProps={{
                    startAdornment: (
                      <InputAdornment position="start">
                        <Icon iconName="Filter" 
                              style={
                                { fontSize: '15px', 
                                color: '#2A7386', 
                                borderRadius: '50%', 
                                padding: '4px',  
                                }
                              } />
                      </InputAdornment>
                    ),
                    endAdornment: filters.location && (
                      <InputAdornment position="end">
                        <IconButton edge="end" size="small" sx={{color:'#2A7386', padding: 0, marginRight: '4px'}} onClick={() => setFilters(prev => ({ ...prev, location: '' }))}>
                          <ClearRoundedIcon  style={{fontSize:'12'}} />
                        </IconButton>
                      </InputAdornment>
                    ),
                  }}
                />
              </Tooltip>
              <Tooltip title="Sort">
                <IconButton size="small" onClick={() => handleSort(columns.filter((column) => column.accessor === 'City')[0].accessor)}>
                  {sortBy === (columns.filter((column) => column.accessor === 'City')[0].accessor) && sortDirection === "desc" ? <ArrowDownward fontSize="small" /> : <ArrowUpward fontSize="small" />}
                </IconButton>
              </Tooltip>
            </div>
          </TableCell>
        </TableRow>
      </TableHead>

      <TableBody>
        {paginatedData.map((row, index) => (
          <TableRow 
              key={`${row.EmpO365Email}-${index}`}
              style= {{
                backgroundColor: index % 2 === 0 ? '#f5f5f5' : 'white',
              }}>

              {columns.map((column) => (
                (column.accessor === "OrgChart")?
                (
                    <TableCell>
                      <div className="styles.iconCircle">
                      <Organization20Filled style={{backgroundColor: '#bfd8db', padding: '4px', borderRadius: '50%'}} className={styles.orgIcon} onClick={() => handleOrgChartClick(row)} />
                      </div>
                    </TableCell>
                ): column.accessor === "EmployeeName" ?
                (
                  <TableCell key={column.accessor} style={{fontSize:'11px', cursor: 'pointer', color: '#2A7386' }}
                              onClick={() => handleEmployeeNameClick(row)}>
                    {row[column.accessor as keyof IEmployee]}
                  </TableCell>
                ):
                (
                  <TableCell key={column.accessor} style={{fontSize:'11px'}}>
                    {row[column.accessor as keyof IEmployee]}
                  </TableCell>
                )
              ))}
          </TableRow>
        ))}
      </TableBody>
    </Table>

      <Box
      sx={{
        display: 'flex',
        alignItems: 'center',
        height: '40px',
        border: 'none',
        borderColor: 'divider',
        borderRadius: 1,
        bgcolor: 'background.paper',
        color: 'text.secondary',
        justifyContent: 'flex-end',
        [`& .${dividerClasses.root}`]: {
          mx: 0.5,
        },
        margin: '5px 5px'
      }}
      >
          <Box sx={{ display: 'flex', flexGrow: 1, alignItems: 'center', marginBottom: '20px'}}>
            <Pagination
              size="small"
              count={totalPages}
              page={page}
              onChange={handleChangePage}
              color="primary"
              defaultPage={1}
              siblingCount={1}
              boundaryCount={1}
              shape="rounded"
              style={{ marginTop: '20px', display: 'flex', justifyContent: 'center' }}
              sx={{
                '& .MuiPaginationItem-root': {
                  backgroundColor: '#f0f0f0', // Set background color of page numbers
                  color: '#2A7386', // Set text color for page numbers
                  '&:hover': {
                    backgroundColor: '#cfcfcf', // Hover effect for page numbers
                  },
                  '&.Mui-selected': {
                    backgroundColor: '#2A7386', // Selected page number background color
                    color: '#fff', // Selected page number text color
                  },
                },
              }}
            />
        </Box>

        <Box sx={{ display: 'flex', alignItems: 'center', margin: '10px'}}>
          
          <Typography variant="body2" style={{ marginRight: '10px', fontSize: '0.8rem'}}>
            Rows per page: 
          </Typography>

          <FormControl style={{ minWidth: 'auto', marginRight: '10px'}}>
                <Select
                  value={rowsPerPage}
                  onChange={handleChangeRowsPerPage}
                  size="small"
                  displayEmpty
                  MenuProps={{
                    PaperProps: {
                      sx: {
                        width: 70,
                        fontSize: '0.75rem',
                        '& .MuiMenuItem-root': {
                          fontSize: '0.75rem',
                          '&:hover': {
                            backgroundColor: '#e0e0e0',
                          },
                        },
                        '& .MuiMenuItem-root.Mui-selected': {
                          fontSize: '0.75rem',
                          backgroundColor: '#d3d3d3',  // Set the selected option background color to light grey
                          '&:hover': {
                            backgroundColor: '#b0b0b0',  // Set a slightly darker grey on hover
                          },
                        },
                      }
                    }
                  }}
                  sx={{
                    fontSize: '0.75rem', // Reduce font size of the select itself
                    padding: '0 2px', // Reduce the padding inside the select box
                    border: '1px solid #2A7386', // Thin grey border
                    minWidth: 'fit-content',
                    borderRadius: '4px',
                    height: '28px',  // Adjust height to reduce space
                    '& .MuiSelect-select': {
                      padding: '0 2px', // Adjust padding inside the select box to reduce space further
                    },
                    '& .MuiSelect-icon': {
                      top: '30%', // Align the dropdown icon vertically
                      right: '4px', // Adjust horizontal position of the arrow
                      transform: 'translateY(-30%)', // Correct vertical alignment
                    },
                    // Remove the outline (border) when selected
                    '& .MuiOutlinedInput-notchedOutline': {
                      borderWidth: '0px', // Ensure border width is consistent and smaller
                    },
                    // Remove the primary color border when focused
                    '&.Mui-focused .MuiOutlinedInput-notchedOutline': {
                      borderColor: 'transparent', // Make the focused border transparent
                    },
                  }}
                  inputProps={{
                    'aria-label': 'Rows per page', // Optional for accessibility
                  }}
                >
                  <MenuItem value={5}>5</MenuItem>
                  <MenuItem value={10}>10</MenuItem>
                  <MenuItem value={25}>25</MenuItem>
                  <MenuItem value={100}>100</MenuItem>
                  <MenuItem value={-1}>Show All</MenuItem>
                </Select>
          </FormControl>
            
          <Typography variant="body2" style={{ marginRight: '10px', fontSize: '0.75rem'  }}>
            {`${startRecord}-${endRecord} of ${sortedData.length}`}
          </Typography>
        </Box>

      </Box>
    </TableContainer>
    
      <Panel isOpen={isOpen} onDismiss={() => setIsOpen(false)} type={PanelType.medium} closeButtonAriaLabel="Close" headerText={selectedEmployee ? `Organization Chart - ${selectedEmployee.EmployeeName} ` : "Org Chart"}>
          <CoreOrgChart serviceScope={props.serviceScope} data={employeesForOrgChart}  showHighlighers={false} />
      </Panel> 
          
       {/* Modal for Employee Details */}
        {/* Modal for Employee Details on click of Org chart employee node*/}
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
            {/* Employee Picture - Positioned on Top Edge */}
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

            {selectedEmployee && (
              <Button
                onClick={() => handleOrgChartClick(selectedEmployee)}
                sx={{
                  position: 'absolute',
                  top: '44px',
                  right: '20px',
                  backgroundColor: '#f5f5f5',
                  color: '#2A7386',
                  fontSize: '0.8rem',
                  fontWeight: 'bold',
                  borderRadius: '8px',
                  padding: '8px 16px',
                  boxShadow: 'none',
                  '&:hover': {
                    backgroundColor: '#E2E2E2',
                  },
                }}
              >
                Organization Chart
              </Button>
            )}
          
          {/* <DefaultButton
              onClick={() => {
                setSelectedOrgChartView("Button2");
                if (selectedEmployee)
                  setSelectedDirectorate(selectedEmployee.Directorate ?? '');
              }}
              styles={{
                root: {
                  position: 'absolute',
                  top: '44px', // Align with the employee name
                  right: '20px', // Keep towards the right edge
                  backgroundColor: '#f5f5f5', // Light gray background
                  color: '#2A7386', // Text color matches the typography
                  fontSize: '0.8rem', // Font size from Typography
                  fontWeight: 'bold', // Bold text from Typography
                  borderRadius: '8px', // Rounded corners for a sleek look
                  padding: '8px 16px', // Padding for the button
                  border: 'none', // Remove default border
                  boxShadow: 'none', // No shadow for a flat design
                },
                rootHovered: {
                  backgroundColor: '#e0e0e0', // Slightly darker on hover
                },
              }}
            >
              View Organization Chart
            </DefaultButton> */}

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
                  {selectedEmployee.EmployeeName}
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
  </div>
  );
};