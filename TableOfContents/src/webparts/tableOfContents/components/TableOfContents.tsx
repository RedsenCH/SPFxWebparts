import * as React from 'react';
import styles from './TableOfContents.module.scss';
import type { ITableOfContentsProps } from './ITableOfContentsProps';
import { IButtonStyles, IconButton } from '@fluentui/react/lib/Button';
import { IIconProps, Stack } from '@fluentui/react';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { DisplayMode } from '@microsoft/sp-core-library';

/**
 * Describes a link for a header
 */
interface ITOCLink {
  /**
   * The Source html element.
   */
  element: HTMLElement | undefined;
  /**
   * Child nodes for the link.
   */
  childNodes: ITOCLink[];
  /**
   * Parent link. Undefined for the root link.
   */
  parent: ITOCLink | undefined;
}

export interface ITableOfContentsState {
  showMenu:boolean;
}

export default class TableOfContents extends React.Component<ITableOfContentsProps, ITableOfContentsState> {
  
  private static timeout = 500;

  private static h2Tag = "h2";
  private static h3Tag = "h3";
  private static h4Tag = "h4";

  constructor(props: ITableOfContentsProps) {
    super(props);

    this.state = {
      showMenu:false,
    };

    this.toggleShowHideFloatingMenu = this.toggleShowHideFloatingMenu.bind(this);
  }

  /**
   * Force the component to re-render with a specified interval.
   * This is needed to get valid id values for headers to use in links. Right after the rendering headers won't have valid ids, they are assigned later once the whole page got rendered.
   * The component will display the correct list of headers on the first render and will be able to process clicks (as a link to an HTMLElement is stored by the component).
   * Once valid ids got assigned to headers by SharePoint code, the component will get valid ids for headers. This way a link from ToC can be copied by a user and it will be a valid link to a header.
   */
  public componentDidMount(): void {
    setInterval(() => {
      this.setState({});
    }, TableOfContents.timeout);

    if (this.props.customStyles && this.props.customStyles.length > 0 ) {
      const styleBlock = document.createElement("style");
      styleBlock.innerHTML = this.props.customStyles;
      document.getElementsByTagName('head')[0].append(styleBlock);
    }
  }
  
  public render(): React.ReactElement<ITableOfContentsProps> {
    // get headers, then filter out empty and headers from <aside> tags
    const querySelector = this.getQuerySelector(this.props);
    const headers = this.getHtmlElements(querySelector).filter(this.filterEmpty).filter(this.filterAside);
    // create a list of links from headers
    const links = this.getLinks(headers);
    // create components from a list of links
    const toc = (<ul>{this.renderLinks(links, 0)}</ul>);
    // add CSS class to hide in mobile view if needed
    const hideInMobileViewClass = this.props.hideInMobileView ? (styles.hideInMobileView) : '';


    const floatingButtonStyles:IButtonStyles = {
      root:{
        borderRadius: "50%",
        width:40,
        height:40,
        backgroundColor: this.props.floatingMenuButtonBackgroundColor ? this.props.floatingMenuButtonBackgroundColor : this.props.theme?.palette?.themePrimary,
        borderColor: this.props.theme?.palette?.white,
        borderWidth:1,
        borderStyle: "solid",
        boxShadow: "rgb(0 0 0 / 13%) 0px 1.6px 3.6px 0px, rgb(0 0 0 / 11%) 0px 0.3px 0.9px 0px",
        transition: "box-shadow 0.5s ease 0s",
      },
      rootHovered: {
        backgroundColor: this.props.floatingMenuButtonBackgroundColor ? this.props.floatingMenuButtonBackgroundColor : this.props.theme?.palette?.themePrimary
      },
      rootPressed: {
        backgroundColor: this.props.floatingMenuButtonBackgroundColor ? this.props.floatingMenuButtonBackgroundColor : this.props.theme?.palette?.themePrimary
      },
      icon:{
        fontSize:20,
        color: this.props.floatingMenuButtonIconColor ? this.props.floatingMenuButtonIconColor : this.props.theme?.palette?.white
      },
      iconHovered:{
        fontSize:20,
        color: this.props.floatingMenuButtonIconColor ? this.props.floatingMenuButtonIconColor : this.props.theme?.palette?.white
      }
    };

    const floatingButtonIcon:IIconProps = {
      iconName : this.props.floatingMenuButtonIcon ? this.props.floatingMenuButtonIcon : 'BulletedList2'
    };



    return (
      <>
      { (this.props.floatingMenu && this.props.displayMode === DisplayMode.Read) &&
        <div className={styles.floatingTableOfContent}>
          <Stack className={styles.menuContainer} styles={this.state.showMenu ? {root:{display:"block"}} : {root:{display:"none"}}}>
            <section className={styles.tableOfContents}>
              <WebPartTitle displayMode={this.props.displayMode} title={this.props.title} updateProperty={this.props.updateProperty} className={styles.webpartTitle} />
              <nav className='CustomSPToc'>
                {toc}
              </nav>
            </section>
          </Stack>
          <IconButton 
            className={styles.iconButton} 
            iconProps={floatingButtonIcon}
            styles={floatingButtonStyles}
            onClick={() => this.toggleShowHideFloatingMenu()} />
        </div>
      }
      { (!this.props.floatingMenu || this.props.displayMode === DisplayMode.Edit) && 
        <section className={styles.tableOfContents}>
          <div className={hideInMobileViewClass}>
            <WebPartTitle displayMode={this.props.displayMode}
              title={this.props.title}
              updateProperty={this.props.updateProperty} />
            <nav className='CustomSPToc'>
              {toc}
            </nav>
          </div>
        </section>}
      </>
    );
  }

  /**
   * Gets a nested list of links based on the list of headers specified.
   * @param headers List of HtmlElements for H2, H3, and H4 headers.
   */
  private getLinks(headers: HTMLElement[]): ITOCLink[] {
    // create a root link that will be a root for links' tree
    const root: ITOCLink = { childNodes: [], parent: undefined, element: undefined };

    let prevLink: ITOCLink | null = null;

    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      const link: ITOCLink = { childNodes: [], parent: undefined, element: header };

      if (i === 0) {
        // the first header is always added as a child of the root
        link.parent = root;
        root.childNodes.push(link);
      } else {
        const prevHeader = headers[i - 1];

        // compare the current header and the previous one to define where to add new link
        const compare = this.compareHeaders(header.tagName, prevHeader.tagName);

        if (compare === 0) {
          // if headers are on the same level, add header to the same parent
          link.parent = prevLink?.parent;
          prevLink?.parent?.childNodes.push(link);
        } else if (compare < 0) {

          let targetParent = prevLink?.parent;
          // if current header bigger than the previous one, go up in the hierarchy to find a place to add link
          // go up in the hierarchy of links until a link with bigger tag is found or until the root link found
          // i.e. for H4 look for H3 or H2, for H3 look for H2, for H2 look for the root.
          while (targetParent && (targetParent !== root) && (this.compareHeaders(header.tagName, targetParent.element?.tagName) <= 0)) {
            targetParent = targetParent?.parent;
          }

          link.parent = targetParent;
          targetParent?.childNodes.push(link);
        } else {
          // if current header is smaller than the previous one, add link for it as a child of the previous link
          if (prevLink)  {
            link.parent = prevLink;
            prevLink.childNodes.push(link);
          }
        }
      }

      prevLink = link;
    }

    // return list of links for top-level headers
    return root.childNodes;
  }

  /**
   * Compares two header tags by their weights.
   * The function is used to compare the size of headers (e.g. should H3 go under H2?)
   * @param header1
   * @param header2
   */
  private compareHeaders(header1: string | undefined, header2: string | undefined): number {
    return this.getHeaderWeight(header1) - this.getHeaderWeight(header2);
  }

  /**
   * Returns a digital weight of a tag. Used for comparing header tags.
   * @param header
   */
  private getHeaderWeight(header: string | undefined): number {
    switch (header?.toLowerCase()) {
      case (TableOfContents.h2Tag):
        return 2;
      case (TableOfContents.h3Tag):
        return 3;
      case (TableOfContents.h4Tag):
        return 4;
      default:
        throw new Error('Unknown header: ' + header);
    }
  }

  /**
   * Returns html elements in the current page specified by the query selector.
   */
  private getHtmlElements(querySelector: string): HTMLElement[] {
    if (querySelector.length === 0) {
      return [];
    } else {
      const elements = document.querySelectorAll(querySelector);
      const htmlElements: HTMLElement[] = [];

      elements.forEach((element) => {
        htmlElements.push(element as HTMLElement);
      });

      // for (let i = 0; i < elements.length; i++) {
      //   htmlElements.push(elements[i] as HTMLElement);
      // }

      return htmlElements;
    }
  }

  /**
   * Returns a query selector based on the specified props
   * @param props
   */
  private getQuerySelector(props: ITableOfContentsProps): any {
    const queryParts = [];

    if (props.showHeading2) {
      queryParts.push(TableOfContents.h2Tag);
    }

    if (props.showHeading3) {
      queryParts.push(TableOfContents.h3Tag);
    }

    if (props.showHeading4) {
      queryParts.push(TableOfContents.h4Tag);
    }

    return queryParts.join(',');
  }

  /**
   * Filters elements with empty text.
   * @param element
   */
  private filterEmpty(element: HTMLElement): boolean {
    return element.innerText.trim() !== '';
  }

  /**
   * Filters elements that are inside <aside> tag and thus not related to a page.
   * @param element
   */
  private filterAside(element: HTMLElement): boolean {
    let inAsideTag = false;

    let parentElement = element.parentElement;

    while (parentElement) {
      if (parentElement.tagName.toLocaleLowerCase() === 'aside') {
        inAsideTag = true;
        break;
      }

      parentElement = parentElement.parentElement;
    }

    return !inAsideTag;
  }

  /**
   * Returns a click handler that scrolls a page to the specified element.
   */
  private scrollToHeader = (target: HTMLElement): any => {
    return (event: React.SyntheticEvent) => {
      event.preventDefault();
      document.location.hash = target.id;
      target.scrollIntoView({ behavior: 'smooth', block: 'start', inline: 'nearest' });
    };
  }

  /**
   * Creates a list of components to display from a list of links.
   * @param links
   */
  private renderLinks(links: ITOCLink[], currentLevel:number = 0): JSX.Element[] {
    // for each link render a <li> element with a link. If the link has got childNodes, additionaly render <ul> with child links.
    const nestedLevel = currentLevel + 1;
    
    let elements : any = []
    elements = links.map((link, index) => {
      if (link && link.element) {
        return (
          <li key={index} className={`toc_level_${currentLevel}`}>
            <a onClick={this.scrollToHeader(link.element)} href={'#' + link.element.id}>{link.element.innerText}</a>
            {link.childNodes.length > 0 ? (<ul>{this.renderLinks(link.childNodes, nestedLevel)}</ul>) : ''}
          </li>
        );
      }
    });

    if (elements) {
      return elements;
    } else {
      return [];
    }
    // return elements ? elements : [];
  }

  private toggleShowHideFloatingMenu(): void {
    const showMenu = !this.state.showMenu;

    this.setState({
      ...this.state,
      showMenu
    });
  }
}
