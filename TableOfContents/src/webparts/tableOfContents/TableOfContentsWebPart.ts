import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneLabel,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from "@pnp/spfx-property-controls/lib/PropertyFieldColorPicker";
import { PropertyFieldIconPicker } from "@pnp/spfx-property-controls/lib/PropertyFieldIconPicker";
import { PropertyPaneWebPartInformation } from '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation';
import { IReadonlyTheme, ThemeChangedEventArgs, ThemeProvider } from '@microsoft/sp-component-base';

import * as strings from 'TableOfContentsWebPartStrings';
import TableOfContents from './components/TableOfContents';
import { ITableOfContentsProps } from './components/ITableOfContentsProps';

export interface ITableOfContentsWebPartProps {
  title: string;
  showSectionTitles: boolean;
  showHeading1: boolean;
  showHeading2: boolean;
  showHeading3: boolean;

  floatingMenu: boolean;
  floatingMenuButtonIcon: string;
  floatingMenuLabel: string;
  floatingMenuOpenOnLoad: boolean;
  floatingMenuButtonIconColor: string;
  floatingMenuButtonBackgroundColor: string;

  hideInMobileView: boolean;
  customStyles: string;
}

export default class TableOfContentsWebPart extends BaseClientSideWebPart<ITableOfContentsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  public render(): void {
    const element: React.ReactElement<ITableOfContentsProps> = React.createElement(
      TableOfContents,
      {
        // description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,

        theme: this._themeVariant,
        title: this.properties.title,
        displayMode: this.displayMode,
        updateProperty: this.handleUpdateProperty,

        showHeading2: this.properties.showHeading1,
        showHeading3: this.properties.showHeading2,
        showHeading4: this.properties.showHeading3,

        floatingMenu: this.properties.floatingMenu,
        floatingMenuButtonIcon: this.properties.floatingMenuButtonIcon,
        floatingMenuOpenOnLoad: this.properties.floatingMenuOpenOnLoad,
        floatingMenuButtonIconColor: this.properties.floatingMenuButtonIconColor,
        floatingMenuButtonBackgroundColor: this.properties.floatingMenuButtonBackgroundColor,
        floatingMenuLabel: this.properties.floatingMenuLabel,

        hideInMobileView: this.properties.hideInMobileView,
        customStyles: this.properties.customStyles,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    
    const floatingMenu = PropertyPaneToggle("floatingMenu", {label:strings.MenuType, onText:strings.MenuTypeFloating, offText:strings.MenuTypeStandard });

    const floatingMenuButtonIcon = PropertyFieldIconPicker('floatingMenuButtonIcon', {
      currentIcon: this.properties.floatingMenuButtonIcon ? this.properties.floatingMenuButtonIcon : "BulletedList2",
      key: "iconPickerId",
      onSave: (icon: string) => { this.properties.floatingMenuButtonIcon = icon; },
      onChanged:(icon: string) => { this.properties.floatingMenuButtonIcon = icon; },
      renderOption: "dialog",
      properties: this.properties,
      onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
      label: strings.floatingMenuButtonIcon,          
    });

    const floatingMenuButtonIconColor = PropertyFieldColorPicker("floatingMenuButtonIconColor", {
      key: "floatingMenuButtonIconColor",
      label: strings.floatingMenuButtonIconColor,
      selectedColor: this.properties.floatingMenuButtonIconColor ? this.properties.floatingMenuButtonIconColor : this._themeVariant?.palette?.white,
      onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
      properties: this.properties,
      alphaSliderHidden: false,
      style: PropertyFieldColorPickerStyle.Inline,
      iconName: "Color"
    });

    const floatingMenuButtonBackgroundColor = PropertyFieldColorPicker("floatingMenuButtonBackgroundColor", {
      key: "floatingMenuButtonBackgroundColor",
      label: strings.floatingMenuButtonBackgroundColor,
      selectedColor: this.properties.floatingMenuButtonBackgroundColor ? this.properties.floatingMenuButtonBackgroundColor : this._themeVariant?.palette?.themePrimary,
      onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
      properties: this.properties,
      alphaSliderHidden: false,
      style: PropertyFieldColorPickerStyle.Inline,
      iconName: "Color"
    });

    const displayConfiguration = [floatingMenu];

    if (this.properties.floatingMenu) {
      displayConfiguration.push(floatingMenuButtonIcon);
      displayConfiguration.push(floatingMenuButtonIconColor);
      displayConfiguration.push(floatingMenuButtonBackgroundColor);
    }

    
    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          header: {
            description: ""
          },
          groups: [
            {
              groupName: strings.GroupFieldGeneralSettingsTitle,
              isCollapsed:false,
              groupFields: [
                PropertyPaneLabel("", {text:strings.PropertyPaneDescription}),
                PropertyPaneCheckbox('showHeading1', {
                  text: strings.ShowHeading1FieldLabel
                }),
                PropertyPaneCheckbox('showHeading2', {
                  text: strings.ShowHeading2FieldLabel
                }),
                PropertyPaneCheckbox('showHeading3', {
                  text: strings.ShowHeading3FieldLabel
                })
              ]
            },
            {
              groupName: strings.GroupFieldDisplayTitle,
              isCollapsed:false,
              groupFields: displayConfiguration
            },
            {
              groupName: strings.GroupFieldCustomStylesTitle,
              isCollapsed:true,
              groupFields: [
                PropertyFieldCodeEditor('customStyles', {
                  label: 'Edit CSS Code',
                  panelTitle: 'Edit CSS Code',
                  initialValue: this.properties.customStyles,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  key: 'codeEditorFieldId',
                  language: PropertyFieldCodeEditorLanguages.css,
                  options: {
                    wrap: true,
                    fontSize: 14,
                    // more options
                  }
                }),
                PropertyPaneWebPartInformation({
                  description: `<p>use following CSS structure to customize styles:</p>/* level 1: */<br/>&nbsp;&nbsp;&nbsp;li.toc_level_0 {}<br/>&nbsp;&nbsp;&nbsp;li.toc_level_0 a {}<br/>/* level 2: */<br/>&nbsp;&nbsp;&nbsp;li.toc_level_1 {}<br/>&nbsp;&nbsp;&nbsp;li.toc_level_1 a {}<br/>/* level 3: */<br/>&nbsp;&nbsp;&nbsp;li.toc_level_2 {}<br/>li.toc_level_2 a{}<br/>`,
                  key: 'webPartInfoId'
                })    
              ]
            },
            {
              groupName: strings.GroupFieldMobileViewTitle,
              isCollapsed:true,
              groupFields: [
                PropertyPaneToggle('hideInMobileView', {
                  label: strings.HideInMobileViewLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  /**
   * Saves new value for the title property.
   */
  private handleUpdateProperty = (newValue: string): void => {
    this.properties.title = newValue;
  }

  /**
   * Update the current theme variant reference and re-render.
   * @param args The new theme
   */
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }
}
