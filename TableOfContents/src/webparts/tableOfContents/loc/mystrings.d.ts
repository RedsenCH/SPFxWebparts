declare interface ITableOfContentsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;

  GroupFieldGeneralSettingsTitle: string;
  GroupFieldCustomStylesTitle: string;
  GroupFieldMobileViewTitle: string;
  GroupFieldDisplayTitle: string;
  
  PropertyPaneDescription: string;
  ShowHeading2FieldLabel: string;
  ShowHeading3FieldLabel: string;
  ShowHeading4FieldLabel: string;

  MenuType: string;
  MenuTypeStandard: string;
  MenuTypeFloating: string;

  floatingMenuButtonIcon: string;
  floatingMenuLabel: string;
  floatingMenuOpenOnLoad: boolean;
  floatingMenuButtonIconColor: string;
  floatingMenuButtonBackgroundColor: string;

  HideInMobileViewLabel: string;
  CustomStyles: string;
}

declare module 'TableOfContentsWebPartStrings' {
  const strings: ITableOfContentsWebPartStrings;
  export = strings;
}
