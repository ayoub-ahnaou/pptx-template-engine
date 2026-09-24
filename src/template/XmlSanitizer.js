/**
 * Normalizes fragmented OpenXML text runs (<a:r> / <a:t>) inside paragraphs (<a:p>).
 * Recombines text split across multiple runs so template tags like {{project.name}} 
 * or {{bgColor:...}} exist as unbroken strings.
 */
class XmlSanitizer {
    static sanitizeParagraphs(xmlString) {
        if (!xmlString || typeof xmlString !== "string") return xmlString;

        return xmlString.replace(/<a:p[^>]*?>[\s\S]*?<\/a:p>/gi, (pXml) => {
            const plainText = pXml.replace(/<[^>]+>/g, "");
            if (!plainText.includes("{{")) return pXml;

            const textMatches = [...pXml.matchAll(/<a:t[^>]*?>([\s\S]*?)<\/a:t>/gi)];
            if (textMatches.length <= 1) return pXml;

            const fullText = textMatches.map(m => m[1]).join("");

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