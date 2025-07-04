import { DisplayMode } from '@microsoft/sp-core-library';

export interface IHtmlEditorProps {
  title: string;
  content: string;
  displayMode: DisplayMode;
  hideTitle: boolean;
  removeIframeBorders: boolean;
  isDarkTheme: boolean;
  openAllLinksInNewTab: boolean;
  openPropertyPane: () => void;
  updateTitle: (value: string) => void;
}
