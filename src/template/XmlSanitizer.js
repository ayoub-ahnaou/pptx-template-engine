/**
 * Normalizes fragmented OpenXML text runs (<a:r> / <a:t>) inside paragraphs (<a:p>).
 * Recombines text split across multiple runs so template tags like {{project.name}} 
 * or {{bgColor:...}} exist as unbroken strings.
 */
class XmlSanitizer {
    static sanitizeParagraphs(xmlString) {
        if (!xmlString || typeof xmlString !== "string") return xmlString;

        return xmlString.replace(/<a:p[^>]*?>[\s\S]*?<\/a:p>/gi, (pXml) => {
            // Strip XML tags to inspect plain text content
            const plainText = pXml.replace(/<[^>]+>/g, "");
            
            // If paragraph does not contain mustache syntax {{...}}, leave it untouched
            if (!plainText.includes("{{") || !plainText.includes("}}")) {
                return pXml;
            }

            // Extract all text nodes inside <a:t>
            const textMatches = [...pXml.matchAll(/<a:t[^>]*?>([\s\S]*?)<\/a:t>/gi)];
            if (textMatches.length === 0) return pXml;

            // Combine fragmented text strings
            const fullText = textMatches.map(m => m[1]).join("");

            // Rebuild the paragraph with full text in the first run and empty remaining runs
            let replacedFirst = false;
            return pXml.replace(/<a:t([^>]*?)>[\s\S]*?<\/a:t>/gi, (match, attrs) => {
                if (!replacedFirst) {
                    replacedFirst = true;
                    return `<a:t${attrs}>${fullText}</a:t>`;
                }
                return `<a:t${attrs}></a:t>`;
            });
        });
    }
}

module.exports = XmlSanitizer;