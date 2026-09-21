class PptxShape {

    constructor(xml) {
        this.xml = xml;

        this.id = this.extractAttribute("<p:cNvPr", "id");
        this.name = this.extractAttribute("<p:cNvPr", "name");
    }

    extractAttribute(element, attribute) {
        const regex = new RegExp(`${element}\\s+[^>]*?${attribute}="([^"]*)"`);
        const match = this.xml.match(regex);
        return match ? match[1] : null;
    }

    getText() {
        const regex = /<a:t>([\s\S]*?)<\/a:t>/g;
        const texts = [];
        let match;

        while ((match = regex.exec(this.xml)) !== null) {
            texts.push(match[1]);
        }

        return texts.join("");
    }

    setText(text) {
        const escapedText = this.escapeXml(text);

        // Remove all existing <a:t> contents across paragraphs
        // Put the new full text into the first <a:t> and clear the rest
        let firstReplaced = false;

        this.xml = this.xml.replace(/<a:t>[\s\S]*?<\/a:t>/g, () => {
            if (!firstReplaced) {
                firstReplaced = true;
                return `<a:t>${escapedText}</a:t>`;
            }
            return `<a:t></a:t>`;
        });
    }

    getPosition() {
        const match = this.xml.match(/<a:off\s+x="([^"]+)"\s+y="([^"]+)"/);
        if (!match) return null;
        return {
            x: Number(match[1]),
            y: Number(match[2])
        };
    }

    getSize() {
        const match = this.xml.match(/<a:ext\s+cx="([^"]+)"\s+cy="([^"]+)"/);
        if (!match) return null;
        return {
            width: Number(match[1]),
            height: Number(match[2])
        };
    }

    setPosition(x, y) {
        this.xml = this.xml.replace(
            /(<a:off\s+x=")[^"]+(")\s+(y=")[^"]+(")/,
            `$1${x}$2 $3${y}$4`
        );
    }

    clone(newId) {
        const clonedXml = this.xml.replace(
            /(<p:cNvPr\s+id=")[^"]+/,
            `$1${newId}`
        );
        return new PptxShape(clonedXml);
    }

    getXml() {
        return this.xml;
    }

    escapeXml(value) {
        return String(value)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;");
    }
}

module.exports = PptxShape;