class PptxSlide {

    constructor(xml) {
        this.xml = xml;
    }

    getTextNodes() {

        const nodes = [];

        const regex =
            /<a:t>([\s\S]*?)<\/a:t>/g;

        let match;

        while (
            (match = regex.exec(this.xml)) !== null
        ) {

            nodes.push({
                text: match[1],
                start: match.index,
                end: regex.lastIndex
            });
        }

        return nodes;
    }

    getTexts() {

        return this
            .getTextNodes()
            .map(node => node.text);
    }

    setText(oldText, newText) {

        const escapedOldText =
            this.escapeRegExp(oldText);

        const regex =
            new RegExp(
                `(<a:t>)${escapedOldText}(</a:t>)`
            );

        this.xml =
            this.xml.replace(
                regex,
                `$1${this.escapeXml(newText)}$2`
            );
    }

    getXml() {
        return this.xml;
    }

    escapeRegExp(value) {

        return value.replace(
            /[.*+?^${}()|[\]\\]/g,
            "\\$&"
        );
    }

    escapeXml(value) {

        return value
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;");
    }
}

module.exports = PptxSlide;