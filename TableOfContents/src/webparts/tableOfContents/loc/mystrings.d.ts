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
    ShowHeading1FieldLabel: string;
    ShowHeading2FieldLabel: string;
    ShowHeading3FieldLabel: string;

    MenuType: string;
    MenuTypeStandard: string;
    MenuTypeFloating: string;

    floatingMenuButtonIcon: string;
    floatingMenuLabel: string;
    floatingMenuOpenOnLoad: boolean;
    floatingMenuButtonIconColor: string;
    floatingMenuButtonBackgroundColor: string;
    floatingMenuBackgroundColor: string;
    floatingMenuLinksColor: string;
    floatingMenuLinkIcon: string;
    floatingMenuLinkTextDecoration: string;
    floatingMenuLinkTextDecorationNone: string;
    floatingMenuLinkTextDecorationUnderline: string;
    floatingMenuBorderColor: string;

    HideInMobileViewLabel: string;
    CustomStyles: string;
}

declare module "TableOfContentsWebPartStrings" {
    const strings: ITableOfContentsWebPartStrings;
    export = strings;
}
