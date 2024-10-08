declare interface IEnhancedPowerAppsWebPartStrings {
  ThemeValuePostLabel: string;
  ThemeValueFieldLabel: string;
  ThemeValuePreLabel: string;
  ThemeGroupName: string;
  DynamicsPropNameDescriptionLabel: string;
  DynamicPropsNameFieldLabel: string;
  DynamicPropFieldLabel: string;
  SelectDynamicSource: string;
  UseDynamicPropsFieldLabel: string;
  DynamicsPropsGroupDescription2: string;
  DynamicsPropsGroupDescription1: string;
  DynamicPropertiesGroupLabel: string;
  WidthFieldLabel: string;
  AspectRatioCustomOption: string;
  AspectRatioFieldLabel: string;
  HeightFieldLabel: string;
  LayoutAspectRatioOption: string;
  LayoutFixedHeightOption: string;
  LayoutFieldLabel: string;
  BorderFieldLabel: string;
  AppearanceGroupName: string;
  PlaceholderButtonLabel: string;
  PlaceholderDescription: string;
  PlaceholderIconText: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  AppWebLinkFieldLabel: string;
  StaticPropertiesGroupLabel:string;
  StaticPropertiesGroupDescription:string;
  StaticPropertyNameFieldLabel: string;
  StaticPropertyValueLabel: string;
  UseStaticPropsFieldLabel: string;
  PropertyPane: {
    DynamicFieldCollection: {
      ParameterType: {
        Title: string;
        OptionLabelDynamic: string;
        OptionLabelStatic: string;
      }
    }
  }
}

declare module 'EnhancedPowerAppsWebPartStrings' {
  const strings: IEnhancedPowerAppsWebPartStrings;
  export = strings;
}
