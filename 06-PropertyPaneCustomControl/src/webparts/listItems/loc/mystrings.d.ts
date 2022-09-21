declare interface IListItemsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  ListFieldLabel: string;
  ItemFieldLabel: string;
}

declare module 'ListItemsWebPartStrings' {
  const strings: IListItemsWebPartStrings;
  export = strings;
}