declare interface IEventDetailsWebPartStrings {
  listItemFieldLabel: string;
  TitleViewFieldLabel: string;
  StartViewFieldLabel: string;
  EndViewFieldLabel: string;
  LocationViewFieldLabel: string;
  ViewFieldsGroup: string;
  RegisterViewFieldLabel:string;
  SaveEventViewFieldLabel:string;
  SingUp:string;
  UnSubscribe:string;
  SaveEvent:string;
  eventlistid:string;
}

declare module 'EventDetailsWebPartStrings' {
  const strings: IEventDetailsWebPartStrings;
  export = strings;
}
