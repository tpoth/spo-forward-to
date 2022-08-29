declare interface IForwardToWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ForwardToUrlFieldLabel: string;
  ForwardingActiveFieldLabel: string;
  ForwardingDelayFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'ForwardToWebPartStrings' {
  const strings: IForwardToWebPartStrings;
  export = strings;
}
