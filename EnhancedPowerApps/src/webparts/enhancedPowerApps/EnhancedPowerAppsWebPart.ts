import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneLabel,
  PropertyPaneChoiceGroup,
  // IPropertyPaneField
} from '@microsoft/sp-property-pane';

import * as strings from 'EnhancedPowerAppsWebPartStrings';
import EnhancedPowerApps from './components/EnhancedPowerApps';
import { IEnhancedPowerAppsProps } from './components/IEnhancedPowerAppsProps';

/**
 * Use this for dynamic properties
 */
import { DynamicProperty } from '@microsoft/sp-component-base';

/**
 * Plain old boring web part thingies
 */
import {
  BaseClientSideWebPart,
  IWebPartPropertiesMetadata,
} from '@microsoft/sp-webpart-base';

import {
  PropertyPaneDynamicFieldSet,
  PropertyPaneDynamicField
} from '@microsoft/sp-property-pane';


/**
 * Use this for theme awareness
 */
import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme
} from '@microsoft/sp-component-base';

/**
 * Use the multi-select for large checklists
 */
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
// import { PropertyFieldCollectionData, CustomCollectionFieldType, IPropertyFieldCollectionDataPropsInternal } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { ThemeVariantSlots } from './ThemeVariantSlots';

/**
 * Super-cool text functions included in SPFx that people don't use often enough
 */
import { Text } from '@microsoft/sp-core-library';
import { PropertyPaneWebPartInformation } from '@pnp/spfx-property-controls';

export interface IEnhancedPowerAppsWebPartProps {
  displayMode: string;
  appWebLink: string;
  dynamicProp: DynamicProperty<string>;
  useDynamicProp: boolean;

  dynamicPropertiesConfig: any[];

  /** BEGIN - Old implementation kept for backward compatibility*/
  dynamicPropName: string;
  useStaticProp: boolean;
  staticPropName: string;
  staticPropValue: string;
  /** END - Old implementation kept for backward compatibility*/

  border: boolean;
  layout: 'FixedHeight'|'AspectRatio';
  height: number;
  width: number;
  aspectratio: '16:9'|'3:2'|'16:10'|'4:3'|'Custom';
  themeValues: string[];
}

export default class EnhancedPowerAppsWebPart extends BaseClientSideWebPart<IEnhancedPowerAppsWebPartProps> {
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  protected onInit(): Promise<void> {
    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    return super.onInit();
  }

  public render(): void {
    // Context variables and dynamic properties
    const dynamicProp: string | undefined = this.properties.dynamicProp.tryGetValue();
    const locale: string = this.context.pageContext.cultureInfo.currentCultureName;

    // Get the client width. This is how we'll calculate the aspect ratio and resize the iframe
    const { clientWidth } = this.domElement;

    // Get the aspect width and height based on aspect ratio for the web part
    let aspectWidth: number;
    let aspectHeight: number;
    switch(this.properties.aspectratio) {
      case "16:10":
        aspectWidth = 16;
        aspectHeight = 10;
        break;
      case "16:9":
        aspectWidth = 16;
        aspectHeight = 9;
        break;
      case "3:2":
        aspectWidth = 3;
        aspectHeight = 2;
        break;
      case "4:3":
        aspectWidth = 4;
        aspectHeight = 3;
        break;
      case "Custom":
        // Custom aspects just use the width and height properties
        aspectWidth = this.properties.width;
        aspectHeight = this.properties.height;
    }

    // If we're using fixed height, we pass the height and don't resize, otherwise we
    // calculate the height based on the web part's width and selected aspect ratio
    const clientHeight: number = this.properties.layout === 'FixedHeight' ?
      this.properties.height :
      clientWidth * (aspectHeight/aspectWidth);

    const element: React.ReactElement<IEnhancedPowerAppsProps> = React.createElement(
      EnhancedPowerApps,
      {
        locale: locale,
        dynamicProp: dynamicProp,
        useDynamicProp: this.properties.useDynamicProp,
        dynamicPropName: this.properties.dynamicPropName,
        useStaticProp: this.properties.useStaticProp,
        staticPropName: this.properties.staticPropName,
        staticPropValue: this.properties.staticPropValue,
        onConfigure: this._onConfigure,
        appWebLink: this.properties.appWebLink,
        width: clientWidth,
        height: clientHeight,
        themeVariant: this._themeVariant,
        border: this.properties.border,
        themeValues: this.properties.themeValues
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    
    // #region Appearance property fields
    const _appearanceGroupFields:any[] = [
      PropertyPaneToggle('border', {
        label: strings.BorderFieldLabel
      }),
      PropertyPaneChoiceGroup('layout', {
        label: strings.LayoutFieldLabel,
        options: [
          {
            key: 'FixedHeight',
            text: strings.LayoutFixedHeightOption,
            iconProps: {
              officeFabricIconFontName: 'FullWidth'
            }
          },
          {
            key: 'AspectRatio',
            text: strings.LayoutAspectRatioOption,
            iconProps: {
              officeFabricIconFontName: 'AspectRatio'
            }
          }
        ]
      })
    ];

    if (this.properties.layout === "FixedHeight") {
      _appearanceGroupFields.push(
        PropertyPaneTextField('height', {
          label: strings.HeightFieldLabel
        })
      );
    }

    if (this.properties.layout === "AspectRatio") {
      _appearanceGroupFields.push(
        PropertyPaneChoiceGroup('aspectratio', {
          label: strings.AspectRatioFieldLabel,
          options: [
            {
              key: '16:9',
              text: '16:9',
            },
            {
              key: '3:2',
              text: '3:2',
            },
            {
              key: '16:10',
              text: '16:10',
            },
            {
              key: '4:3',
              text: '4:3',
            },
            {
              key: 'Custom',
              text: strings.AspectRatioCustomOption,
            }
          ]
        })
      );

      if (this.properties.aspectratio === "Custom") {
        _appearanceGroupFields.push(
          PropertyPaneTextField('width', {
            label: strings.WidthFieldLabel,
          })
        );

        _appearanceGroupFields.push(
          PropertyPaneTextField('height', {
            label: strings.HeightFieldLabel,
          })
        );
      }
    }
    // #endregion

    // #region Dynamic property fields
    const _dynamicPropertiesGroupFields:any[] = [
      PropertyPaneWebPartInformation({
        description: Text.format(strings.DynamicsPropsGroupDescription1, this.properties.dynamicPropName!== undefined ?this.properties.dynamicPropName:'parametername'),
        moreInfoLink: null,
        videoProperties: null,
        key: 'dynamicPropertiesId1'
      }),
      PropertyPaneWebPartInformation({
        description: strings.DynamicsPropsGroupDescription2,
        moreInfoLink: null,
        videoProperties: null,
        key: 'dynamicPropertiesId2'
      }),


      PropertyPaneToggle('useDynamicProp', {
        checked: this.properties.useDynamicProp === true,
        label: strings.UseDynamicPropsFieldLabel
      })
    ];

    if (this.properties.useDynamicProp === true) {
      
      // _dynamicPropertiesGroupFields.push(this._buildPropertyFieldCollectionData());
      
      
      _dynamicPropertiesGroupFields.push(PropertyPaneDynamicFieldSet({
        label: strings.SelectDynamicSource,
        fields: [
          PropertyPaneDynamicField('dynamicProp', {
            label: strings.DynamicPropFieldLabel
          })
        ]
      }));
      _dynamicPropertiesGroupFields.push(PropertyPaneTextField('dynamicPropName', {
        label: strings.DynamicPropsNameFieldLabel,
        description: strings.DynamicsPropNameDescriptionLabel,
        value: this.properties.dynamicPropName
      }));
    }
    // #endregion

    // #region Static property fields
    const _staticPropertiesGroupFields:any[] = [
      PropertyPaneWebPartInformation({
        description: Text.format(strings.StaticPropertiesGroupDescription, this.properties.staticPropName !== undefined ? this.properties.staticPropName:'parametername'),
        moreInfoLink: null,
        videoProperties: null,
        key: 'useStaticProp'
      }),
      PropertyPaneToggle('useStaticProp', {
        checked: this.properties.useStaticProp === true,
        label: strings.UseStaticPropsFieldLabel
      })
    ];

    if(this.properties.useStaticProp === true) {
      _staticPropertiesGroupFields.push(PropertyPaneTextField('staticPropName', {
        label: strings.StaticPropertyNameFieldLabel,
        description: "",
        value: this.properties.staticPropName
      }));
      _staticPropertiesGroupFields.push(PropertyPaneTextField('staticPropValue', {
        label: strings.StaticPropertyValueLabel,
        description: "",
        value: this.properties.staticPropValue
      }));
    }
    // #endregion

    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              isCollapsed: false,
              groupFields: [
                PropertyPaneTextField('appWebLink', {
                  label: strings.AppWebLinkFieldLabel
                })
              ]
            },
            {
              groupName: strings.AppearanceGroupName,
              isCollapsed: true,
              groupFields: _appearanceGroupFields
            },
            {
              groupName: strings.DynamicPropertiesGroupLabel,
              isCollapsed: true,
              groupFields: _dynamicPropertiesGroupFields
            },
            {
              groupName: strings.StaticPropertiesGroupLabel,
              
              isCollapsed: true,
              groupFields: _staticPropertiesGroupFields
            },
            {
              groupName: strings.ThemeGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneLabel('themeValuesPre',{
                  text: strings.ThemeValuePreLabel
                }),
                PropertyFieldMultiSelect('themeValues', {
                  key: 'multithemeValuesSelect',
                  label: strings.ThemeValueFieldLabel,
                  options: ThemeVariantSlots,
                  selectedKeys: this.properties.themeValues
                }),
                PropertyPaneWebPartInformation({
                  description: strings.ThemeValuePostLabel,
                  moreInfoLink: null,
                  videoProperties: null,
                  key: 'themeValuesPost'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  
  // private _buildPropertyFieldCollectionData(): IPropertyPaneField<IPropertyFieldCollectionDataPropsInternal> {
  //   return PropertyFieldCollectionData("dynamicPropertiesConfig", {
  //     key: "collectionData",
  //     label: "Collection data",
  //     panelHeader: "Collection data panel header",
  //     manageBtnLabel: "Manage collection data",
  //     value: this.properties.dynamicPropertiesConfig,
  //     fields: [
  //       {
  //         id: "type",
  //         title: strings.PropertyPane.DynamicFieldCollection.ParameterType.Title,
  //         type: CustomCollectionFieldType.dropdown,
  //         options: [
  //           {
  //             key: "dynamic",
  //             text: strings.PropertyPane.DynamicFieldCollection.ParameterType.OptionLabelDynamic
  //           },
  //           {
  //             key: "static",
  //             text: strings.PropertyPane.DynamicFieldCollection.ParameterType.OptionLabelStatic
  //           }
  //         ],
  //         required: true
  //       },
  //       {
  //         id: "test",
  //         title: "test",
  //         type: CustomCollectionFieldType.custom
  //       },
  //       {
  //         id: "Lastname",
  //         title: "Lastname",
  //         type: CustomCollectionFieldType.string,
  //       },
  //       {
  //         id: "Age",
  //         title: "Age",
  //         type: CustomCollectionFieldType.number,
  //         required: true
  //       },
  //       {
  //         id: "City",
  //         title: "Favorite city",
  //         type: CustomCollectionFieldType.dropdown,
  //         options: [
  //           {
  //             key: "antwerp",
  //             text: "Antwerp"
  //           },
  //           {
  //             key: "helsinki",
  //             text: "Helsinki"
  //           },
  //           {
  //             key: "montreal",
  //             text: "Montreal"
  //           }
  //         ],
  //         required: true
  //       },
  //       {
  //         id: "Sign",
  //         title: "Signed",
  //         type: CustomCollectionFieldType.boolean
  //       }
  //     ],
  //     disabled: false
  //   })
  // }
  
  
  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      // Specify the web part properties data type to allow the address
      // information to be serialized by the SharePoint Framework.
      'dynamicProp': {
        dynamicPropertyType: 'string'
      }
    };
  }

  private _onConfigure = (): void => {
    this.context.propertyPane.open();
  }

  /**
 * Update the current theme variant reference and re-render.
 *
 * @param args The new theme
 */
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }
  
  /**
   * Redraws the web part when resized
   * @param _newWidth
   */
  protected onAfterResize(_newWidth: number): void {
    // redraw the web part
    this.render();
  }
}
