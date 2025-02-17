import * as React from "react";
import { sp } from "@pnp/sp/presets/all";
import {
  DetailsList,
  IColumn,
  DetailsListLayoutMode,
} from "@fluentui/react/lib/DetailsList";
import {
  TextField,
  PrimaryButton,
  IconButton,
  Dialog,
  DialogFooter,
  DatePicker,
  Dropdown,
  IDropdownOption,
} from "@fluentui/react";
import { DefaultButton } from "@fluentui/react/lib/Button";

interface IEmployee {
  Id: number;
  Name: string;
  DOB: string;
  Department: string;
  Experience: number;
}

interface IState {
  employees: IEmployee[];
  filteredEmployees: IEmployee[];
  searchQuery: string;
  isSortedDescending: boolean;
  isDialogOpen: boolean;
  selectedEmployee: IEmployee | null;
  isConfirmationDialogOpen: boolean;
  isConfirmationDialogOpenEdit: boolean;
  isAddDialogOpen: boolean;
  newEmployee: IEmployee;
}

class EmployeeList extends React.Component<{}, IState> {
  constructor(props: {}) {
    super(props);
    this.state = {
      employees: [],
      filteredEmployees: [],
      searchQuery: "",
      isSortedDescending: false,
      isDialogOpen: false,
      selectedEmployee: null,
      isConfirmationDialogOpen: false,
      isConfirmationDialogOpenEdit: false,
      isAddDialogOpen: false,
      newEmployee: { Id: 0, Name: "", DOB: "", Department: "", Experience: 0 },
    };
  }

  departmentOptions: IDropdownOption[] = [
    { key: "HR", text: "HR" },
    { key: "IT", text: "IT" },
    { key: "Sales", text: "Sales" },
  ];

  componentDidMount() {
    const fetchEmployees = async () => {
      try {
        const items = await sp.web.lists
          .getByTitle("Q-14_Employees")
          .items.select("Id", "Name1", "DOB", "Department1", "Experience")
          .get();

        const formattedEmployees = items.map((item: any) => ({
          Id: item.Id,
          Name: item.Name1,
          DOB: new Date(item.DOB).toLocaleDateString(),
          Department: item.Department1,
          Experience: item.Experience,
        }));

        this.setState({
          employees: formattedEmployees,
          filteredEmployees: formattedEmployees,
        });
      } catch (error) {
        console.error("Error fetching employees:", error);
      }
    };

    fetchEmployees();
  }

  onColumnClick = (column: IColumn): void => {
    const { isSortedDescending, filteredEmployees } = this.state;
    if (column.key === "name") {
      const newIsSortedDescending = !isSortedDescending;
      this.setState({ isSortedDescending: newIsSortedDescending });

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

      this.setState({ filteredEmployees: sortedEmployees });
    }
  };

  onSearch = () => {
    const { employees, searchQuery } = this.state;
    const filtered = employees.filter(
      (employee) =>
        employee.Name.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1
    );
    this.setState({ filteredEmployees: filtered });
  };

  onDelete = async (id: number) => {
    try {
      // await sp.web.lists.getByTitle("Q-14_Employees").items.getById(id).delete();

      this.setState((prevState) => ({
        employees: prevState.employees.filter((employee) => employee.Id !== id),
        filteredEmployees: prevState.filteredEmployees.filter(
          (employee) => employee.Id !== id
        ),
      }));
    } catch (error) {
      console.error("Error deleting employee:", error);
    }
  };

  onEdit = (employee: IEmployee) => {
    this.setState({
      selectedEmployee: employee,
      isDialogOpen: true,
    });
  };

  closeDialog = () => {
    this.setState({
      isDialogOpen: false,
      selectedEmployee: null,
    });
  };

  openConfirmationDialog = () => {
    this.setState({ isConfirmationDialogOpen: true });
  };

  closeConfirmationDialog = () => {
    this.setState({ isConfirmationDialogOpen: false });
  };

  openConfirmationDialogEdit = () => {
    this.setState({ isConfirmationDialogOpenEdit: true });
  };

  closeConfirmationDialogEdit = () => {
    this.setState({ isConfirmationDialogOpenEdit: false });
  };

  onSaveEdit = async () => {
    const { selectedEmployee } = this.state;
    if (!selectedEmployee) return;
    this.openConfirmationDialogEdit();
  };

  onConfirmSave = async () => {
    const { selectedEmployee } = this.state;
    if (!selectedEmployee) return;

    try {
      // await sp.web.lists
      //   .getByTitle("Q-14_Employees")
      //   .items.getById(selectedEmployee.Id)
      //   .update({
      //     Name1: selectedEmployee.Name,
      //     DOB: new Date(selectedEmployee.DOB).toISOString(),
      //     Department1: selectedEmployee.Department,
      //     Experience: selectedEmployee.Experience,
      //   });

      this.setState((prevState) => ({
        employees: prevState.employees.map((emp) =>
          emp.Id === selectedEmployee.Id ? selectedEmployee : emp
        ),
        filteredEmployees: prevState.filteredEmployees.map((emp) =>
          emp.Id === selectedEmployee.Id ? selectedEmployee : emp
        ),
        isDialogOpen: false,
        isConfirmationDialogOpenEdit: false,
      }));
    } catch (error) {
      console.error("Error updating employee:", error);
    }
  };

  onCancelSave = () => {
    this.setState({ isConfirmationDialogOpenEdit: false });
  };

  onAddHandle = () => {
    this.setState({ isAddDialogOpen: true });
  };

  closeAddDialog = () => {
    this.setState({ isAddDialogOpen: false });
  };

  onSaveAdd = async () => {
    this.openConfirmationDialog();
  };

  onConfirmSaveAdd = async () => {
    const { newEmployee } = this.state;
    try {
      const addedEmployee = await sp.web.lists
        .getByTitle("Q-14_Employees")
        .items.add({
          Name1: newEmployee.Name,
          DOB: new Date(newEmployee.DOB).toISOString(),
          Department1: newEmployee.Department,
          Experience: newEmployee.Experience,
        });

      this.setState((prevState) => ({
        employees: [
          ...prevState.employees,
          {
            Id: addedEmployee.data.Id,
            Name: newEmployee.Name,
            DOB: new Date(newEmployee.DOB).toLocaleDateString(),
            Department: newEmployee.Department,
            Experience: newEmployee.Experience,
          },
        ],
        filteredEmployees: [
          ...prevState.filteredEmployees,
          {
            Id: addedEmployee.data.Id,
            Name: newEmployee.Name,
            DOB: new Date(newEmployee.DOB).toLocaleDateString(),
            Department: newEmployee.Department,
            Experience: newEmployee.Experience,
          },
        ],
        isAddDialogOpen: false,
        isConfirmationDialogOpen: false,
      }));
    } catch (error) {
      console.error("Error adding employee:", error);
    }
  };

  render() {
    const {
      filteredEmployees,
      searchQuery,
      isSortedDescending,
      isDialogOpen,
      selectedEmployee,
      isConfirmationDialogOpen,
      isConfirmationDialogOpenEdit,
      isAddDialogOpen,
    } = this.state;

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
              onClick={() => this.onEdit(item)}
            />
            <IconButton
              iconProps={{ iconName: "Delete" }}
              title="Delete"
              onClick={() => this.onDelete(item.Id)}
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
        isSortedDescending,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        onColumnClick: (ev, column) => this.onColumnClick(column),
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
            onChange={(e, newValue) =>
              this.setState({ searchQuery: newValue || "" })
            }
            placeholder="Enter name here"
            styles={{ root: { maxWidth: 300 } }}
          />
          <PrimaryButton
            text="Search"
            onClick={this.onSearch}
            style={{ marginLeft: 10 }}
          />
          <PrimaryButton
            text="Add Employee"
            onClick={this.onAddHandle}
            style={{ marginLeft: 20 }}
          />
        </div>

        <DetailsList
          items={filteredEmployees}
          columns={columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.fixedColumns}
        />

        {selectedEmployee && (
          <Dialog
            hidden={!isDialogOpen}
            onDismiss={this.closeDialog}
            dialogContentProps={{
              title: "Edit Employee",
            }}
          >
            <TextField
              label="Name"
              value={selectedEmployee.Name}
              onChange={(e, newValue) =>
                this.setState({
                  selectedEmployee: {
                    ...selectedEmployee!,
                    Name: newValue || "",
                  },
                })
              }
            />
            <DatePicker
              label="Date of Birth"
              value={new Date(selectedEmployee.DOB)}
              onSelectDate={(date) =>
                date &&
                this.setState({
                  selectedEmployee: {
                    ...selectedEmployee!,
                    DOB: date.toLocaleDateString(),
                  },
                })
              }
            />
            <Dropdown
              label="Department"
              selectedKey={selectedEmployee.Department}
              options={this.departmentOptions}
              onChange={(e, option) =>
                option &&
                this.setState({
                  selectedEmployee: {
                    ...selectedEmployee!,
                    Department: option.key as string,
                  },
                })
              }
            />
            <TextField
              label="Experience"
              value={selectedEmployee.Experience.toString()}
              onChange={(e, newValue) =>
                this.setState({
                  selectedEmployee: {
                    ...selectedEmployee!,
                    Experience: parseInt(newValue || "0", 10),
                  },
                })
              }
              type="number"
            />
            <DialogFooter>
              <PrimaryButton text="Save" onClick={this.onSaveEdit} />
              <DefaultButton text="Cancel" onClick={this.closeDialog} />
            </DialogFooter>
          </Dialog>
        )}

        {isConfirmationDialogOpenEdit && (
          <Dialog
            hidden={!isConfirmationDialogOpenEdit}
            onDismiss={this.closeConfirmationDialogEdit}
            dialogContentProps={{
              title: "Confirm Action",
              subText: "Do you want to save the changes?",
            }}
          >
            <DialogFooter>
              <PrimaryButton onClick={this.onConfirmSave} text="Yes" />
              <DefaultButton
                onClick={this.closeConfirmationDialogEdit}
                text="No"
              />
            </DialogFooter>
          </Dialog>
        )}

        {isAddDialogOpen && (
          <Dialog
            hidden={!isAddDialogOpen}
            onDismiss={this.closeAddDialog}
            dialogContentProps={{
              title: "Add New Employee",
            }}
          >
            <TextField
              label="Name"
              value={this.state.newEmployee.Name}
              onChange={(e, newValue) =>
                this.setState({
                  newEmployee: {
                    ...this.state.newEmployee,
                    Name: newValue || "",
                  },
                })
              }
            />
            <DatePicker
              label="Date of Birth"
              value={
                this.state.newEmployee.DOB
                  ? new Date(this.state.newEmployee.DOB)
                  : undefined
              }
              onSelectDate={(date) =>
                date &&
                this.setState({
                  newEmployee: {
                    ...this.state.newEmployee,
                    DOB: date.toLocaleDateString(),
                  },
                })
              }
            />
            <TextField
              label="Experience"
              value={this.state.newEmployee.Experience.toString()}
              onChange={(e, newValue) =>
                this.setState({
                  newEmployee: {
                    ...this.state.newEmployee,
                    Experience: parseInt(newValue || "0", 10),
                  },
                })
              }
            />
            <Dropdown
              label="Department"
              selectedKey={this.state.newEmployee.Department}
              options={this.departmentOptions}
              onChange={(e, option) =>
                this.setState({
                  newEmployee: {
                    ...this.state.newEmployee,
                    Department: option?.key as string,
                  },
                })
              }
            />

            <DialogFooter>
              <PrimaryButton onClick={this.onSaveAdd} text="Add" />
              <DefaultButton onClick={this.closeAddDialog} text="Cancel" />
            </DialogFooter>
          </Dialog>
        )}

        {isConfirmationDialogOpen && (
          <Dialog
            hidden={!isConfirmationDialogOpen}
            onDismiss={this.closeConfirmationDialog}
            dialogContentProps={{
              title: "Confirm Action",
              subText: "Do you want to add new employee?",
            }}
          >
            <DialogFooter>
              <PrimaryButton onClick={this.onConfirmSaveAdd} text="Yes" />
              <DefaultButton onClick={this.closeConfirmationDialog} text="No" />
            </DialogFooter>
          </Dialog>
        )}
      </div>
    );
  }
}

export default EmployeeList;
