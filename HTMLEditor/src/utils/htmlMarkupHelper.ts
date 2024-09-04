export class HtmlMarkupHelper {

    /**
     * Gets the template HTML markup in the full template content
     * @param templateContent the full template content
     */
    public static getDocumentFromString(htmlMarkup: string): Document {
        const domParser = new DOMParser();
        const htmlContent: Document = domParser.parseFromString(htmlMarkup, 'text/html');
        return htmlContent;
    }

    /**
     * Check if content has any javascript code
     * @param htmlContent the content
     * @returns true if javascript code is found, false otherwise
     */
    public static hasJavascript(htmlContent: string): boolean {
        if (htmlContent && htmlContent.length > 0 && htmlContent.trim().length > 0) {
            const regex:RegExp = /<script/g;
            const result = regex.test(htmlContent.toLowerCase());
            return result;
        } else {
            return false;
        }
    }

    /**
     * Check if content has any iframe code
     * @param htmlcontent the content
     * @returns true if iframe code is found, false otherwise
     */
    public static hasIframe(htmlcontent: string): boolean {
        if (htmlcontent && htmlcontent.length > 0 && htmlcontent.trim().length > 0) {
            const regex:RegExp = /<iframe/g;
            return regex.test(htmlcontent.toLowerCase());
        } else {
            return false;
        }
    }

    /**
     * Check if content has any forbidden code
     * @param htmlcontent the content
     * @returns true if forbidden code is found, false otherwise
     */
    public static hasForbiddenCode(htmlcontent: string): boolean {
        const hasJS = this.hasJavascript(htmlcontent);
        const hasIframe = this.hasIframe(htmlcontent);
        
        return (hasJS || hasIframe);
    }
}