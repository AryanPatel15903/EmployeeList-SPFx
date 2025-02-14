import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import {sp} from '@pnp/sp/presets/all';

import * as strings from 'EmployeeListingWebPartStrings';
import EmployeeListing from './components/EmployeeListing';
import { IEmployeeListingProps } from './components/IEmployeeListingProps';

export interface IEmployeeListingWebPartProps {
  description: string;
}

export default class EmployeeListingWebPart extends BaseClientSideWebPart<IEmployeeListingWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IEmployeeListingProps> = React.createElement(
      EmployeeListing,
      {
        description: this.properties.description,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    });

    return super.onInit().then(() => {
      // this.fetchEmployeeListDetails(); 
    });
  }



  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
