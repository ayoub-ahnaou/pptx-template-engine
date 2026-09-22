const PptxShape = require("./PptxShape");

class PptxSlide {

    constructor(xml) {
        this.xml = xml;
    }

    getShapes() {
        const shapes = [];
        const regex = /<p:sp>[\s\S]*?<\/p:sp>/g;
        let match;

        while ((match = regex.exec(this.xml)) !== null) {
            shapes.push(new PptxShape(match[0]));
        }

        return shapes;
    }

    getTextShapes() {
        return this.getShapes().filter(
            shape => shape.getText().length > 0
        );
    }

    getNextShapeId() {
        const shapes = this.getShapes();
        const ids = shapes
            .map(shape => Number(shape.id))
            .filter(id => Number.isInteger(id));

        if (ids.length === 0) return 1;
        return Math.max(...ids) + 1;
    }

    insertShape(shape) {
        const newXml = shape.getXml();
        const position = this.xml.lastIndexOf("</p:spTree>");

        if (position === -1) {
            throw new Error("Could not find <p:spTree>");
        }

        this.xml =
            this.xml.slice(0, position) +
            newXml +
            this.xml.slice(position);
    }

    /**
     * Removes a shape from the slide XML reliably using its unique shape ID or XML content.
     */
    removeShape(shape) {
        if (!shape) return;

        // Try removing by shape ID first (most reliable)
        if (shape.id) {
            const shapeIdRegex = new RegExp(`<p:sp>(?:(?!<\/p:sp>)[\s\S])*?<p:cNvPr[^>]*?id="${shape.id}"[\s\S]*?<\/p:sp>`, 'g');
            if (shapeIdRegex.test(this.xml)) {
                this.xml = this.xml.replace(shapeIdRegex, "");
                return;
            }
        }

        // Fallback: try direct XML replace
        const shapeXml = shape.getXml();
        if (this.xml.includes(shapeXml)) {
            this.xml = this.xml.replace(shapeXml, "");
        }
    }

    getXml() {
        return this.xml;
    }
}

module.exports = PptxSlide;