declare interface IMyAgendaWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  ViewAll: string;
  NoMeetings: string;
  Loading: string;
}

declare module 'MyAgendaWebPartStrings' {
  const strings: IMyAgendaWebPartStrings;
  export = strings;
}
