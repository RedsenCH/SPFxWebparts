import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { DisplayMode } from "@microsoft/sp-core-library";

export interface ITableOfContentsProps {
    isDarkTheme: boolean;

    environmentMessage: string;
    hasTeamsContext: boolean;
    userDisplayName: string;

    theme: IReadonlyTheme | undefined;

    title: string;
    displayMode: DisplayMode;

    updateProperty: (value: string) => void;

    showHeading2: boolean;
    showHeading3: boolean;
    showHeading4: boolean;

    hideInMobileView: boolean;

    floatingMenu: boolean;
    floatingMenuOpenOnLoad: boolean;
    floatingMenuButtonIcon: string;
    floatingMenuButtonIconColor: string;
    floatingMenuButtonBackgroundColor: string;
    floatingMenuLabel: string;
    floatingMenuLinksColor: string;
    floatingMenuBackgroundColor: string;
    floatingMenuLinkTextDecoration: string;
    floatingMenuLinkIcon?: string;
    floatingMenuBorderColor: string;

    customStyles: string;
}
