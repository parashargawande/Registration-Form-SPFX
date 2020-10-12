import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDateTimeFieldValue } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';

export interface IRegistrationFormProps {
  lists: string;
  context: WebPartContext;
  domElement:HTMLElement;
  startDate: IDateTimeFieldValue;
  endDate: IDateTimeFieldValue;
  additionalInformation: string;
  homePageUrl: string;
}
