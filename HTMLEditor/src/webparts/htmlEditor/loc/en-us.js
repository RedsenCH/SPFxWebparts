define([], function() {
  return {
    PropertyPane: {
      RemovePadding: {
        Label: "Remove top/bottom padding of web part container",
        OnText: "Remove padding",
        OffText: "Keep padding"
      },
      HideTitle: {
        Label: "Show/Hide title",
        OnText: "Hide title",
        OffText: "Show title"
      },
      CodeEditor: {
        Label: "Edit HTML Code",
        PanelTitle: "Edit HTML Code",
        ErrorMessage: "JavaScript code or iframe found in code. Due to security restriction these parts of code won't be rendered or executed."
      }
    },
    Placeholder: {
      Title: "HTML Editor",
      Description: "It this you have no code yet here. Please click on the button below to start configuring your component.",
      ButtonLabel: "Edit markup"
    }
  }
});