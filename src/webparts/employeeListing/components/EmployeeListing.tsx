import * as React from 'react';
// import styles from './EmployeeListing.module.scss';
import type { IEmployeeListingProps } from './IEmployeeListingProps';
import { sp } from '@pnp/sp';
import EmployeeList from './EmployeeList';

export interface IEmployeeListingState { myName: string; }

export default class EmployeeListing extends React.Component<IEmployeeListingProps, IEmployeeListingState> {

  constructor(props: IEmployeeListingProps, state: IEmployeeListingState) {
    super(props);
    this.state = { myName: "Employee Listing" };
  }


  public render(): React.ReactElement<IEmployeeListingProps> {
    
    const { myName } = this.state;

    return (
      <section>
        <h2>{myName}</h2>
        <div>
          <EmployeeList />
        </div>
      </section>
    );
  }

  async componentDidMount()  {
    let data = await sp.web.lists.getByTitle("Q-14_Employees").items.get();
    console.log("Employee Data:", data);
  }
}
