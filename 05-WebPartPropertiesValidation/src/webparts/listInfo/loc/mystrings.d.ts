declare interface IListInfoWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  ListNameFieldLabel: string;
}

declare module 'ListInfoWebPartStrings' {
  const strings: IListInfoWebPartStrings;
  export = strings;
}
