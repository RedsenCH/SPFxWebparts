declare interface IHtmlEditorWebPartStrings {
  PropertyPane: {
    RemovePadding: {
      Label: string;
      OnText: string;
      OffText: string;
    },
    HideTitle: {
      Label: string;
      OnText: string;
      OffText: string;
    },
    OpenAllLinksInNewTab : {
      Label : string;
      OnText: string;
      OffText: string
    },
    CodeEditor: {
      Label: string;
      PanelTitle: string;
      ErrorMessage: string;
    },
    RemoveIframeBorder: {
      Label: string;
      OnText: string;
      OffText: string;
    }
  },
  Placeholder: {
    Title: string;
    Description: string;
    ButtonLabel: string;
  }
}

declare module 'HtmlEditorWebPartStrings' {
  const strings: IHtmlEditorWebPartStrings;
  export = strings;
}
