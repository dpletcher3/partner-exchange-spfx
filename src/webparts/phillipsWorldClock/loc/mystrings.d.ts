declare interface IPhillipsWorldClockWebPartStrings {
  PropertyPaneDescription: string;
  ContentGroupName: string;
  SectionTitleFieldLabel: string;
  ClocksFieldLabel: string;
  ClocksPanelHeader: string;
  ClocksManageButtonLabel: string;
  ClockTitleColumnLabel: string;
  ClockTimezoneColumnLabel: string;
  ClockTimezoneInvalidError: string;
  ClockTimezoneRequiredError: string;
  TimezoneHelpLinkText: string;
  UnconfiguredMessage: string;
}

declare module 'PhillipsWorldClockWebPartStrings' {
  const strings: IPhillipsWorldClockWebPartStrings;
  export = strings;
}
