const ExpressionEvaluator = require("./ExpressionEvaluator");

class PptxBlockRenderer {

    constructor(data) {
        this.data = data;
        this.evaluator = new ExpressionEvaluator(data);
    }

    render(slide, tree) {
        // 1. First, check if slide XML contains any table loops (e.g. {{#risks}} inside <a:tbl>)
        this.renderAllTableSections(slide);

        // 2. Then, process standard shape loops from the AST tree
        for (const node of tree) {
            if (node.type === "section" && !this.isTableLoop(slide, node)) {
                this.renderSection(slide, node);
            }
        }
    }

    /**
     * Scans raw slide XML for any table loops {{#expression}}...{{/expression}} inside <a:tbl>
     * and processes them even if PptxTemplateParser skipped graphic frames.
     */
    renderAllTableSections(slide) {
        const slideXml = slide.getXml();
        
        // Find all loop start tags inside the slide XML
        const loopStartMatches = [...slideXml.matchAll(/\{\{#([^{}]+)\}\}/g)];

        for (const match of loopStartMatches) {
            const expression = match[1].trim();

            // Skip 'if' and 'eachSlide' directives
            if (expression.startsWith("if ") || expression.startsWith("eachSlide ")) {
                continue;
            }

            const section = { expression };

            // Check if this specific loop tag resides inside an <a:tbl> element
            if (this.isTableLoop(slide, section)) {
                this.renderTableSection(slide, section);
            }
        }
    }

    /**
     * Determines if the section loop exists inside a PowerPoint table (<a:tbl>).
     */
    isTableLoop(slide, section) {
        const slideXml = slide.getXml();
        const startTag = `{{#${section.expression}}}`;
        
        // Check if the start tag sits inside a table element
        const tableRegex = new RegExp(`<a:tbl[^>]*?>[\\s\\S]*?${this.escapeRegex(startTag)}[\\s\\S]*?<\\/a:tbl>`, 'i');
        return tableRegex.test(slideXml);
    }

    /**
     * Renders array items by duplicating data rows containing embedded loop tags inside <a:tbl>.
     * Guarantees the header row is NEVER included in the cloned row buffer.
     */
    renderTableSection(slide, section) {
        let slideXml = slide.getXml();
        const items = this.resolve(this.data, section.expression);

        if (!Array.isArray(items)) {
            throw new Error(`Section '${section.expression}' must be an array`);
        }

        const startTag = `{{#${section.expression}}}`;
        const endTag = `{{/${section.expression}}}`;

        // 1. Extract all individual <a:tr>...</a:tr> rows inside the table
        const trRegex = /<a:tr[^>]*?>[\s\S]*?<\/a:tr>/gi;
        const allRows = slideXml.match(trRegex);

        if (!allRows) return;

        // 2. Find the exact row index that contains the start tag
        let targetRowIndex = -1;
        for (let i = 0; i < allRows.length; i++) {
            if (allRows[i].includes(startTag)) {
                targetRowIndex = i;
                break;
            }
        }

        if (targetRowIndex === -1) return;

        // The isolated data row XML
        const templateRowXml = allRows[targetRowIndex];
        let renderedRowsXml = "";

        // 3. Duplicate ONLY the target data row for each item
        for (const itemContext of items) {
            let rowInstanceXml = templateRowXml;

            // Remove control loop tags from row XML
            rowInstanceXml = rowInstanceXml.replace(startTag, "").replace(endTag, "");

            // Replace variables ({{name}}, {{level}}, etc.)
            rowInstanceXml = this.replaceVariables(rowInstanceXml, itemContext);

            renderedRowsXml += rowInstanceXml;
        }

        // 4. Replace ONLY the single target template row in slideXml
        slideXml = slideXml.replace(templateRowXml, renderedRowsXml);
        slide.setXml(slideXml);
    }

    renderSection(slide, section) {
        const items = this.resolve(this.data, section.expression);

        if (!Array.isArray(items)) {
            throw new Error(`Section '${section.expression}' must be an array`);
        }

        const allShapes = this.collectShapes(section);

        for (const shape of allShapes) {
            slide.removeShape(shape);
        }

        if (section.startShape) slide.removeShape(section.startShape);
        if (section.endShape) slide.removeShape(section.endShape);

        const spacing = 900000;

        for (let index = 0; index < items.length; index++) {
            const context = items[index];
            const shapes = this.renderNodes(section.children, context);

            for (const templateShape of shapes) {
                const clone = templateShape.clone(slide.getNextShapeId());
                let text = clone.getText();

                text = this.replaceVariables(text, context);
                clone.setText(text);

                const position = clone.getPosition();
                if (position) {
                    clone.setPosition(position.x, position.y + (index * spacing));
                }

                slide.insertShape(clone);
            }
        }
    }

    renderSingleItem(slide, tree, context) {
        for (const node of tree) {
            if (node.type === "section") {
                const selectedShapes = this.renderNodes(node.children, context);
                const allShapes = this.collectShapes(node);

                for (const shape of allShapes) {
                    if (!selectedShapes.includes(shape)) {
                        slide.removeShape(shape);
                    }
                }

                for (const shape of selectedShapes) {
                    const originalXml = shape.getXml();
                    const currentText = shape.getText();
                    const updatedText = this.replaceVariables(currentText, context);

                    shape.setText(updatedText);
                    slide.xml = slide.xml.replace(originalXml, shape.getXml());
                }

                if (node.startShape) slide.removeShape(node.startShape);
                if (node.endShape) slide.removeShape(node.endShape);
                if (node.elseShape) slide.removeShape(node.elseShape);
            }
        }
    }

    renderNodes(nodes, context) {
        const result = [];

        for (const node of nodes) {
            if (node.type === "shape") {
                result.push(node.shape);
                continue;
            }

            if (node.type === "condition") {
                const value = this.evaluator.evaluate(node.expression, context);
                const selected = value ? node.children : node.elseChildren;
                result.push(...this.renderNodes(selected, context));
            }
        }

        return result;
    }

    collectShapes(section) {
        const result = [];

        const collect = nodes => {
            for (const node of nodes) {
                if (node.type === "shape") {
                    result.push(node.shape);
                } else if (node.type === "condition") {
                    collect(node.children);
                    collect(node.elseChildren);

                    if (node.startShape) result.push(node.startShape);
                    if (node.endShape) result.push(node.endShape);
                    if (node.elseShape) result.push(node.elseShape);
                }
            }
        };

        collect(section.children);
        collect(section.elseChildren);

        return result;
    }

    replaceVariables(text, context) {
        return text.replace(
            /\{\{([^{}]+)\}\}/g,
            (match, expression) => {
                const trimmed = expression.trim();
                if (trimmed.startsWith("#") || trimmed.startsWith("/")) {
                    return match;
                }

                const value = this.resolve(context, trimmed) ?? this.resolve(this.data, trimmed);

                if (value === undefined || value === null) {
                    return match;
                }

                return this.escapeXml(String(value));
            }
        );
    }

    resolve(context, expression) {
        const parts = expression.split(".");
        let current = context;

        for (const part of parts) {
            if (current === null || current === undefined) {
                return undefined;
            }

            if (typeof current !== "object" || !(part in current)) {
                return undefined;
            }

            current = current[part];
        }

        return current;
    }

    escapeXml(value) {
        return String(value)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;");
    }

    escapeRegex(string) {
        return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
}

module.exports = PptxBlockRenderer;