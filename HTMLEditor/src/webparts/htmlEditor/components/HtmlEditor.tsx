import * as React from 'react';
// import styles from './HtmlEditor.module.scss';
import type { IHtmlEditorProps } from './IHtmlEditorProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import * as DOMPurify from 'dompurify';
import * as strings from 'HtmlEditorWebPartStrings';

export default class HtmlEditor extends React.Component<IHtmlEditorProps, {}> {

  /* eslint-disable @typescript-eslint/no-explicit-any */
  constructor(props: IHtmlEditorProps, state: any) {
    super(props);
  }
   /* eslint-enable @typescript-eslint/no-explicit-any */

  public render(): React.ReactElement<IHtmlEditorProps> {

    const cleanHTML = DOMPurify.sanitize(this.props.content, {FORBID_TAGS: ['script', 'iframe'], ADD_TAGS: ['style'], FORCE_BODY: true});
    const cleanContent = <span dangerouslySetInnerHTML={{ __html: cleanHTML }} />;

    return (
      <div>
        {!this.props.hideTitle && <WebPartTitle 
              displayMode={this.props.displayMode}
              title={this.props.title}
              updateProperty={this.props.updateTitle} />}

        {(this.props.displayMode === DisplayMode.Edit && (!this.props.content || this.props.content.length === 0)) && 
          <Placeholder iconName='PasteAsCode'
                     iconText={strings.Placeholder.Title}
                     description={strings.Placeholder.Description}
                     buttonLabel={strings.Placeholder.ButtonLabel}
                     onConfigure={this.props.openPropertyPane} />
         || 
          <div>
            {cleanContent}
          </div>
        }
      </div>
    );
  }

  public componentDidUpdate(prevProps:IHtmlEditorProps): void {
    if (prevProps.content !== this.props.content) {
      this.render();
    }
  }
}
