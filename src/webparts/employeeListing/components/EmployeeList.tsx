import * as React from "react";
import { sp } from "@pnp/sp/presets/all"; // PnPjs import
import {
  DetailsList,
  IColumn,
  DetailsListLayoutMode,
} from "@fluentui/react/lib/DetailsList";
import { TextField, PrimaryButton, IconButton, Dialog, DialogFooter, DatePicker, Dropdown, IDropdownOption } from "@fluentui/react";
import { DefaultButton } from "@fluentui/react/lib/Button";

interface IEmployee {
  Id: number; // Add Id for delete operation
  Name: string;
  DOB: string;
  Department: string;
  Experience: number;
}

const EmployeeList: React.FC = () => {
  const [employees, setEmployees] = React.useState<IEmployee[]>([]);
  const [filteredEmployees, setFilteredEmployees] = React.useState<IEmployee[]>([]);
  const [searchQuery, setSearchQuery] = React.useState<string>(""); // Search query state
  const [isSortedDescending, setIsSortedDescending] = React.useState<boolean>(false);
  
  // Edit dialog state
  const [isDialogOpen, setIsDialogOpen] = React.useState<boolean>(false);
  const [selectedEmployee, setSelectedEmployee] = React.useState<IEmployee | null>(null);

  // Confirmation dialog state
  const [isConfirmationDialogOpen, setIsConfirmationDialogOpen] = React.useState<boolean>(false);
  
  // Department options (You can fetch this from SharePoint or hardcode them if you prefer)
  const departmentOptions: IDropdownOption[] = [
    { key: "HR", text: "HR" },
    { key: "IT", text: "IT" },
    { key: "Sales", text: "Sales" },
    // Add other departments as needed
  ];

  // Fetch employee data from the SharePoint list
  React.useEffect(() => {
    const fetchEmployees = async () => {
      try {
        const items = await sp.web.lists
          .getByTitle("Q-14_Employees")
          .items.select("Id", "Name1", "DOB", "Department1", "Experience")
          .get();

        // Map the SharePoint list items to our employee format
        const formattedEmployees = items.map((item: any) => ({
          Id: item.Id,
          Name: item.Name1,
          DOB: new Date(item.DOB).toLocaleDateString(),
          Department: item.Department1,
          Experience: item.Experience,
        }));

        setEmployees(formattedEmployees);
        setFilteredEmployees(formattedEmployees); // Initially show all employees
      } catch (error) {
        console.error("Error fetching employees:", error);
      }
    };

    fetchEmployees();
  }, []);

  // Sorting function to toggle between ascending and descending order
  const onColumnClick = (column: IColumn): void => {
    if (column.key === "name") {
      const newIsSortedDescending = !isSortedDescending;
      setIsSortedDescending(newIsSortedDescending);

      const sortedEmployees = [...filteredEmployees].sort((a, b) => {
        const aName = a.Name.toLowerCase();
        const bName = b.Name.toLowerCase();

        if (aName < bName) {
          return newIsSortedDescending ? 1 : -1;
        }
        if (aName > bName) {
          return newIsSortedDescending ? -1 : 1;
        }
        return 0;
      });

      setFilteredEmployees(sortedEmployees);
    }
  };

  // Handle search button click to filter employees by Name
  const onSearch = () => {
    const filtered = employees.filter((employee) =>
      employee.Name.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1
    );
    setFilteredEmployees(filtered);
  };

  // Handle delete operation
  const onDelete = async (id: number) => {
    try {
      // Delete the employee from the SharePoint list
      // await sp.web.lists.getByTitle("Q-14_Employees").items.getById(id).delete();

      // Filter out the deleted employee from the state
      setEmployees((prevEmployees) =>
        prevEmployees.filter((employee) => employee.Id !== id)
      );
      setFilteredEmployees((prevEmployees) =>
        prevEmployees.filter((employee) => employee.Id !== id)
      );
    } catch (error) {
      console.error("Error deleting employee:", error);
    }
  };

  // Open the Edit dialog
  const onEdit = (employee: IEmployee) => {
    setSelectedEmployee(employee);
    setIsDialogOpen(true);
  };

  // Close the Edit dialog
  const closeDialog = () => {
    setIsDialogOpen(false);
    setSelectedEmployee(null);
  };

  // Open the Confirmation dialog
  const openConfirmationDialog = () => {
    setIsConfirmationDialogOpen(true);
  };

  // Close the Confirmation dialog
  const closeConfirmationDialog = () => {
    setIsConfirmationDialogOpen(false);
  };

  // Handle the form submit for updating the employee
  const onSaveEdit = async () => {
    if (!selectedEmployee) return;
    openConfirmationDialog(); // Show confirmation dialog before updating
  };

  // Save the update to SharePoint if confirmed
  const onConfirmSave = async () => {
    if (!selectedEmployee) return;

    try {
      // Update the employee details in SharePoint
      // await sp.web.lists
      //   .getByTitle("Q-14_Employees")
      //   .items.getById(selectedEmployee.Id)
      //   .update({
      //     Name1: selectedEmployee.Name,
      //     DOB: new Date(selectedEmployee.DOB).toISOString(),
      //     Department1: selectedEmployee.Department,
      //     Experience: selectedEmployee.Experience,
      //   });

      // Update the state with the new data
      setEmployees((prevEmployees) =>
        prevEmployees.map((emp) =>
          emp.Id === selectedEmployee.Id ? selectedEmployee : emp
        )
      );
      setFilteredEmployees((prevEmployees) =>
        prevEmployees.map((emp) =>
          emp.Id === selectedEmployee.Id ? selectedEmployee : emp
        )
      );

      closeDialog(); // Close the dialog after saving
      closeConfirmationDialog(); // Close the confirmation dialog
    } catch (error) {
      console.error("Error updating employee:", error);
    }
  };

  // Cancel the update and close the dialog
  const onCancelSave = () => {
    closeConfirmationDialog(); // Close the confirmation dialog without saving
  };

  // Define the columns for the DetailsList
  const columns: IColumn[] = [
    {
      key: "actions",
      name: "Actions",
      fieldName: "Actions",
      minWidth: 50,
      maxWidth: 100,
      isMultiline: false,
      onRender: (item: IEmployee) => (
        <div>
          <IconButton
            iconProps={{ iconName: "Edit" }}
            title="Edit"
            onClick={() => onEdit(item)} // Open edit dialog
          />
          <IconButton
            iconProps={{ iconName: "Delete" }}
            title="Delete"
            onClick={() => onDelete(item.Id)}
            style={{ marginLeft: 10 }}
          />
        </div>
      ),
    },
    {
      key: "name",
      name: "Name",
      fieldName: "Name",
      minWidth: 100,
      maxWidth: 200,
      isMultiline: false,
      isSorted: true,
      isSortedDescending: isSortedDescending,
      sortAscendingAriaLabel: "Sorted A to Z",
      sortDescendingAriaLabel: "Sorted Z to A",
      onColumnClick: (ev, column) => onColumnClick(column),
    },
    {
      key: "dob",
      name: "Date of Birth",
      fieldName: "DOB",
      minWidth: 100,
      maxWidth: 200,
      isMultiline: false,
    },
    {
      key: "department",
      name: "Department",
      fieldName: "Department",
      minWidth: 100,
      maxWidth: 200,
      isMultiline: false,
    },
    {
      key: "experience",
      name: "Experience",
      fieldName: "Experience",
      minWidth: 100,
      maxWidth: 200,
      isMultiline: false,
    },
  ];

  return (
    <div>
      <div style={{ margin: 20, display: "flex" }}>
        <TextField
          value={searchQuery}
          onChange={(e, newValue) => setSearchQuery(newValue || "")}
          placeholder="Enter name here"
          styles={{ root: { maxWidth: 300 } }}
        />
        <PrimaryButton text="Search" onClick={onSearch} style={{ marginLeft: 10 }} />
      </div>

      <DetailsList
        items={filteredEmployees}
        columns={columns}
        setKey="set"
        layoutMode={DetailsListLayoutMode.fixedColumns}
      />

      {/* Edit Dialog */}
      {selectedEmployee && (
        <Dialog
          hidden={!isDialogOpen}
          onDismiss={closeDialog}
          dialogContentProps={{
            title: "Edit Employee",
          }}
        >
          <TextField
            label="Name"
            value={selectedEmployee.Name}
            onChange={(e, newValue) => setSelectedEmployee({ ...selectedEmployee, Name: newValue || "" })}
          />
          <DatePicker
            label="Date of Birth"
            value={new Date(selectedEmployee.DOB)}
            onSelectDate={(date) => date && setSelectedEmployee({ ...selectedEmployee, DOB: date.toLocaleDateString() })}
            // formatDate={(date) => date && date.toLocaleDateString()}
          />
          <Dropdown
            label="Department"
            selectedKey={selectedEmployee.Department}
            options={departmentOptions}
            onChange={(e, option) => option && setSelectedEmployee({ ...selectedEmployee, Department: option.key as string })}
          />
          <TextField
            label="Experience"
            value={selectedEmployee.Experience.toString()}
            onChange={(e, newValue) => setSelectedEmployee({ ...selectedEmployee, Experience: parseInt(newValue || "0", 10) })}
            type="number"
          />
          <DialogFooter>
            <PrimaryButton text="Save" onClick={onSaveEdit} />
            <DefaultButton text="Cancel" onClick={closeDialog} />
          </DialogFooter>
        </Dialog>
      )}

      {/* Confirmation Dialog */}
      <Dialog
        hidden={!isConfirmationDialogOpen}
        onDismiss={closeConfirmationDialog}
        dialogContentProps={{
          title: "Are you sure, you want to update the details?",
        }}
      >
        <DialogFooter>
          <PrimaryButton text="Yes" onClick={onConfirmSave} />
          <DefaultButton text="No" onClick={onCancelSave} />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default EmployeeList;
