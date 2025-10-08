import * as React from 'react';
import * as ReactDom from 'react-dom';
import PnPTelemetry from "@pnp/telemetry-js";
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  IPropertyPaneField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'HtmlEditorWebPartStrings';
import { IHtmlEditorProps } from './components/IHtmlEditorProps';
import { PropertyFieldMessage} from '@pnp/spfx-property-controls/lib/PropertyFieldMessage';
import { MessageBarType } from '@fluentui/react';
import { HtmlMarkupHelper } from '../../utils/htmlMarkupHelper';

export interface IHtmlEditorWebPartProps {
  title: string;
  HtmlContent: string;
  removePadding: boolean;
  hideTitle: boolean;
  removeIframeBorders: boolean;
  openAllLinksInNewTab: boolean;
}

export default class HtmlEditorWebPart extends BaseClientSideWebPart<IHtmlEditorWebPartProps> {
  private _propertyPaneCodeEditorLoader: any; // eslint-disable-line @typescript-eslint/no-explicit-any
  private _isDarkTheme: boolean = false;
  private _showErrorMessage: boolean = false;

  constructor() {
    super();
    this.htmlContentUpdate = this.htmlContentUpdate.bind(this);
    
    const telemetry = PnPTelemetry.getInstance();
    telemetry.optOut()
  }

  private htmlContentUpdate(_property: string, _oldVal: string, newVal: string): void {
    this.properties.HtmlContent = newVal;
    this._propertyPaneCodeEditorLoader.initialValue = newVal;

    this._showErrorMessage = HtmlMarkupHelper.hasForbiddenCode(newVal);
  }

  public async render(): Promise<void> {
    if (this.properties.removePadding 
        && this.domElement.parentElement
        && this.displayMode === DisplayMode.Read) {
      let element:HTMLElement = this.domElement.parentElement;
      // check up to 5 levels up for padding and exit once found
      for (let i = 0; i < 5; i++) {
          const style = window.getComputedStyle(element);
          const hasPadding = style.paddingTop !== "0px";
          if (hasPadding) {
              element.style.paddingTop = "0px";
              element.style.paddingBottom = "0px";
              element.style.marginTop = "0px";
              element.style.marginBottom = "0px";
          }

          if (element.parentElement) {
            element = element.parentElement;
          }
      }
    }

    await this.renderCompnentContent();
  }

  private async renderCompnentContent(): Promise<void> {
    // Dynamically load the editor pane to reduce overall bundle size
    const compnent = await import(
      /* webpackChunkName: 'htmlEditor' */
      './components/HtmlEditor'
    );
    
    const element: React.ReactElement<IHtmlEditorProps> = React.createElement(
      compnent.default,
      {
        title: this.properties.title,
        content: this.properties.HtmlContent,
        hideTitle: this.properties.hideTitle,
        removeIframeBorders: this.properties.removeIframeBorders,
        displayMode: this.displayMode,
        isDarkTheme: this._isDarkTheme,
        openAllLinksInNewTab : this.properties.openAllLinksInNewTab,
        openPropertyPane: () => {
          this.context.propertyPane.open();
        },
        updateTitle: (value: string) => {
          this.properties.title = value;
          // BEGIN - Workaround for bug : https://github.com/pnp/sp-dev-fx-controls-react/issues/1877
          this.render(); // eslint-disable-line @typescript-eslint/no-floating-promises
          // END 
        }
      }
    );

    ReactDom.render(element, this.domElement);
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

  protected async loadPropertyPaneResources(): Promise<void> {
    const editorProp = await import(
        /* webpackChunkName: 'scripteditor' */
        '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor'
    );

    this._propertyPaneCodeEditorLoader = editorProp.PropertyFieldCodeEditor('scriptCode', {
        label: strings.PropertyPane.CodeEditor.Label,
        panelTitle: strings.PropertyPane.CodeEditor.PanelTitle,
        initialValue: this.properties.HtmlContent,
        onPropertyChange: this.htmlContentUpdate,
        properties: this.properties,
        disabled: false,
        key: 'codeEditorFieldId',
        language: editorProp.PropertyFieldCodeEditorLanguages.HTML
    });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    
    this._showErrorMessage = HtmlMarkupHelper.hasForbiddenCode(this.properties.HtmlContent);
    
    /* eslint-disable @typescript-eslint/no-explicit-any */
    const webPartOptions: IPropertyPaneField<any>[] = [
      PropertyPaneToggle("removePadding", {
          label: strings.PropertyPane.RemovePadding.Label,
          checked: this.properties.removePadding,
          onText: strings.PropertyPane.RemovePadding.OnText,
          offText: strings.PropertyPane.RemovePadding.OffText
      }),
      PropertyPaneToggle("hideTitle", {
        label: strings.PropertyPane.HideTitle.Label,
        checked: this.properties.hideTitle,
        onText: strings.PropertyPane.HideTitle.OnText,
        offText: strings.PropertyPane.HideTitle.OffText
      }),
      PropertyPaneToggle("openAllLinksInNewTab", {
        label: strings.PropertyPane.OpenAllLinksInNewTab.Label,
        checked: this.properties.openAllLinksInNewTab,
        onText: strings.PropertyPane.OpenAllLinksInNewTab.OnText,
        offText: strings.PropertyPane.OpenAllLinksInNewTab.OffText
    }),
      PropertyFieldMessage("", {
          key: "htmlMarkupMessageKey",
          text: strings.PropertyPane.CodeEditor.ErrorMessage,
          messageType: MessageBarType.error,
          isVisible: this._showErrorMessage
      }),
      this._propertyPaneCodeEditorLoader
    ];
    /* eslint-enable @typescript-eslint/no-explicit-any */

    if (HtmlMarkupHelper.hasIframe(this.properties.HtmlContent)) {
      webPartOptions.push(
        PropertyPaneToggle("removeIframeBorders", {
          label: strings.PropertyPane.RemoveIframeBorder.Label,
          checked: this.properties.removeIframeBorders,
          onText: strings.PropertyPane.RemoveIframeBorder.OnText,
          offText: strings.PropertyPane.RemoveIframeBorder.OffText
        })
      )
    } else {
      this.properties.removeIframeBorders = false;
    }
    
    return {
      pages: [
        {
          groups: [
            {
              groupFields: webPartOptions
            }
          ]
        }
      ]
    };
  }
}
