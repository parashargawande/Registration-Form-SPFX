import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyFieldDateTimePicker, DateConvention, TimeConvention, IDateTimeFieldValue } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';
import { setup as pnpSetup } from "@pnp/common";
import pnp, { Web } from 'sp-pnp-js';  

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  
} from '@microsoft/sp-property-pane';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'RegistrationFormWebPartStrings';
import RegistrationForm from './components/Form/RegistrationForm';
import { IRegistrationFormProps } from './components/Form/IRegistrationFormProps';

export interface IRegistrationFormWebPartProps {
  startDate: IDateTimeFieldValue;
  datetime: IDateTimeFieldValue;
  lists: string ;
  additionalInformation: string;
  homePageUrl: string;
}

export default class RegistrationFormWebPart extends BaseClientSideWebPart<IRegistrationFormWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IRegistrationFormProps> = React.createElement(
      RegistrationForm,
      {
        lists: this.properties.lists,
        context: this.context,
        domElement: this.domElement,
        startDate: this.properties.startDate,
        endDate:this.properties.datetime,
        additionalInformation : this.properties.additionalInformation,
        homePageUrl:this.properties.homePageUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onInit(): Promise<void>{
    return super.onInit().then( _ =>{
      pnp.setup({  
        spfxContext: this.context  
      }); 
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Data Source",
              groupFields: [
                PropertyFieldListPicker('lists', {
                  label: 'Select a list',
                  selectedList: this.properties.lists,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyPaneTextField('homePageUrl', {
                  label: 'Home Page Url',
                  description:'url to redirect user if request is approved'
                })
              ],
            },
            {
              groupName: "Event Information",
              groupFields: [
                PropertyFieldDateTimePicker('startDate', {
                  label: 'Select Start Date',
                  initialDate: this.properties.startDate,
                  dateConvention: DateConvention.Date,
                  timeConvention: TimeConvention.Hours12,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'startDate',
                  showLabels: false
                }),
                PropertyFieldDateTimePicker('datetime', {
                  label: 'Select End Date',
                  initialDate: this.properties.datetime,
                  dateConvention: DateConvention.Date,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'endDate',
                  showLabels: false
                }),
                PropertyPaneTextField('additionalInformation', {
                  label: 'Additional Information'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
