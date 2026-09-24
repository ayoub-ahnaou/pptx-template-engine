const { XMLParser, XMLBuilder } = require("fast-xml-parser");
const ExpressionEvaluator = require("./ExpressionEvaluator");
const XmlSanitizer = require("./XmlSanitizer");

class PptxBlockRenderer {

    constructor(data) {
        this.data = data;
        this.evaluator = new ExpressionEvaluator(data);

        // Options for fast-xml-parser to preserve attributes and XML structure
        this.parserOptions = {
            ignoreAttributes: false,
            attributeNamePrefix: "@_",
            preserveOrder: true,
            commentPropName: "#comment"
        };

        // Instantiate parser and builder instance
        this.parser = new XMLParser(this.parserOptions);
        this.builder = new XMLBuilder(this.parserOptions);
    }

    render(slide, tree) {
        console.log(`Processing Slide ${slide.number || slide.id || '...'}`);

        // 1. Sanitize raw slide XML to repair broken <a:t> text runs across boundaries
        let slideXml = slide.getXml();
        slideXml = XmlSanitizer.sanitizeParagraphs(slideXml);
        slide.setXml(slideXml);

        // 2. Process all table loops first (e.g. {{#risks}} inside <a:tbl>)
        this.renderAllTableSections(slide);

        // 3. Process static slide placeholders (Titles, Subtitles, standalone shapes)
        this.renderGlobalPlaceholders(slide);

        // 4. Process standard AST shape section loops
        for (const node of tree) {
            if (node.type === "section" && !this.isTableLoop(slide, node)) {
                this.renderSection(slide, node);
            }
        }
    }

    /**
     * Replaces variables and handles bgColor/textColor on static shapes (titles, subtitles)
     * across the entire slide XML.
     */
    renderGlobalPlaceholders(slide) {
        let slideXml = slide.getXml();

        if (slideXml.includes("bgColor:")) {
            console.log("-> Found bgColor tag on slide! Processing...");
        }

        // 1. Sanitize text run splits across the entire slide
        slideXml = XmlSanitizer.sanitizeParagraphs(slideXml);

        // 2. Process background fills & text colors across slide XML
        slideXml = this.applyBgColors(slideXml, this.data);
        slideXml = this.applyTextColors(slideXml, this.data);

        // 3. Replace regular variables
        slideXml = this.replaceVariables(slideXml, this.data);

        // 4. Force-erase any lingering color tags in raw text runs (<a:t>...</a:t>)
        slideXml = slideXml.replace(/\{\{(bgColor|textColor):\s*[^{}\s]+\}\}\s*/g, "");

        // 5. Update slide XML
        slide.setXml(slideXml);
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

        const tableRegex = new RegExp(`(<a:tbl[^>]*?>[\\s\\S]*?${this.escapeRegex(startTag)}[\\s\\S]*?${this.escapeRegex(endTag)}[\\s\\S]*?<\\/a:tbl>)`, 'i');
        const match = slideXml.match(tableRegex);

        if (!match) return;

        const rawMatchedTableXml = match[1];
        let originalTableXml = rawMatchedTableXml;

        // 1. Clean table properties
        originalTableXml = originalTableXml.replace(/<a:tblPr([^>]*?)>/i, (tblPrMatch, attrs) => {
            let cleanedAttrs = attrs
                .replace(/firstRow="1"/g, 'firstRow="0"')
                .replace(/bandRow="1"/g, 'bandRow="0"')
                .replace(/firstCol="1"/g, 'firstCol="0"')
                .replace(/lastCol="1"/g, 'lastCol="0"');
            return `<a:tblPr${cleanedAttrs}>`;
        });

        originalTableXml = originalTableXml.replace(/<a:tableStyleId>[^<]*<\/a:tableStyleId>/gi, "");

        // 2. Parse AST
        const tableAst = this.parser.parse(originalTableXml);
        const tblNode = tableAst.find(node => node["a:tbl"]);

        if (!tblNode || !tblNode["a:tbl"]) return;

        const tblChildren = tblNode["a:tbl"];
        const rows = tblChildren.filter(child => child["a:tr"]);

        let startRowIdx = -1;
        let endRowIdx = -1;

        for (let i = 0; i < rows.length; i++) {
            const rowXml = this.builder.build([rows[i]]);
            const rowText = rowXml.replace(/<[^>]+>/g, "");

            if (rowText.includes(startTag)) startRowIdx = i;
            if (rowText.includes(endTag)) endRowIdx = i;
        }

        if (startRowIdx === -1 || endRowIdx === -1) return;

        const headerRows = rows.slice(0, startRowIdx);
        let contentRows = [];

        if (startRowIdx === endRowIdx) {
            contentRows = [rows[startRowIdx]];
        } else if (startRowIdx < endRowIdx) {
            if (startRowIdx + 1 < endRowIdx) {
                contentRows = rows.slice(startRowIdx + 1, endRowIdx);
            } else {
                contentRows = [rows[startRowIdx]];
            }
        }

        const footerRows = rows.slice(endRowIdx + 1);
        const generatedRows = [];

        for (const itemContext of items) {
            for (const templateRowAst of contentRows) {
                let rowXml = this.builder.build([templateRowAst]);

                // Strip loop control tags
                rowXml = rowXml.replace(new RegExp(this.escapeRegex(startTag), 'g'), "");
                rowXml = rowXml.replace(new RegExp(this.escapeRegex(endTag), 'g'), "");

                // Replace variables and apply {{bgColor:...}} / {{textColor:...}}
                rowXml = this.replaceVariables(rowXml, itemContext);

                const newRowAst = this.parser.parse(rowXml);
                generatedRows.push(...newRowAst);
            }
        }

        const firstRowIndex = tblChildren.findIndex(child => child["a:tr"]);
        const updatedTblChildren = [
            ...tblChildren.slice(0, firstRowIndex),
            ...headerRows,
            ...generatedRows,
            ...footerRows
        ];

        tblNode["a:tbl"] = updatedTblChildren;

        const updatedTableXml = this.builder.build(tableAst);
        slideXml = slideXml.replace(rawMatchedTableXml, updatedTableXml);
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

    /**
     * Replaces {{textColor:...}} tags inside text paragraphs and updates <a:rPr> color.
     * Supports dot-notation paths (e.g. report.textColor) and flexible spacing.
     */
    applyTextColors(xmlSnippet, context) {
        return xmlSnippet.replace(/<a:p[^>]*?>[\s\S]*?<\/a:p>/gi, (paraXml) => {
            const plainText = paraXml.replace(/<[^>]+>/g, "");
            const colorTagMatch = plainText.match(/\{\{textColor:\s*([^{}\s]+)\}\}/);

            if (!colorTagMatch) return paraXml;

            const colorKey = colorTagMatch[1].trim();
            const hexRaw = this.resolve(context, colorKey) ?? this.resolve(this.data, colorKey);

            let updatedPara = paraXml.replace(new RegExp(`\\{\\{textColor:\\s*${this.escapeRegex(colorKey)}\\}\\}\\s*`, 'g'), "");

            if (!hexRaw) return updatedPara;

            const hex = String(hexRaw).replace('#', '').trim();
            const colorXml = `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;

            return updatedPara.replace(/<a:rPr([^>]*?)>([\s\S]*?)<\/a:rPr>/gi, (match, attrs, innerProps) => {
                let newInnerProps = innerProps;
                if (/<a:solidFill[^>]*?>[\s\S]*?<\/a:solidFill>/i.test(newInnerProps)) {
                    newInnerProps = newInnerProps.replace(/<a:solidFill[^>]*?>[\s\S]*?<\/a:solidFill>/i, colorXml);
                } else {
                    newInnerProps = colorXml + newInnerProps;
                }
                return `<a:rPr${attrs}>${newInnerProps}</a:rPr>`;
            });
        });
    }

    /**
     * Replaces {{bgColor:...}} tags inside shapes/cells and updates <a:spPr> / <a:tcPr> fill.
     */
    applyBgColors(xmlSnippet, context) {
        // Target full PowerPoint shape XML blocks (<p:sp>...</p:sp>) or table cells (<a:tc>...<\/a:tc>)
        const containerRegex = /(<p:sp[^>]*?>[\s\S]*?<\/p:sp>|<a:tc[^>]*?>[\s\S]*?<\/a:tc>)/gi;

        let processedXml = xmlSnippet.replace(containerRegex, (containerXml) => {
            const plainText = containerXml.replace(/<[^>]+>/g, "");
            const colorTagMatch = plainText.match(/\{\{bgColor:\s*([^{}\s]+)\}\}/);

            if (!colorTagMatch) return containerXml;

            const colorKey = colorTagMatch[1].trim();
            const hexRaw = this.resolve(context, colorKey) ?? this.resolve(this.data, colorKey);

            // 1. Strip the {{bgColor:...}} tag string from internal text runs
            let updatedXml = containerXml.replace(/\{\{bgColor:\s*[^{}\s]+\}\}\s*/g, "");

            if (!hexRaw) return updatedXml;

            const hex = String(hexRaw).replace('#', '').trim();
            const fillXml = `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;

            // 2. Target Shape Properties (<p:spPr>) — Handles Titles, Subtitles, Text Boxes
            if (/<p:spPr[^>]*?>/i.test(updatedXml)) {
                updatedXml = updatedXml.replace(/<p:spPr([^>]*?)>([\s\S]*?)<\/p:spPr>/i, (match, attrs, innerProps) => {
                    let cleanedProps = innerProps.replace(/<a:(solidFill|gradFill|blipFill|pattFill|noFill)[^>]*?(\/>|>[\s\S]*?<\/a:\1>)/gi, "");
                    return `<p:spPr${attrs}>${fillXml}${cleanedProps}</p:spPr>`;
                });
            }
            // 3. Target Table Cell Properties (<a:tcPr>) — Table Cells
            else if (/<a:tcPr[^>]*?>/i.test(updatedXml)) {
                updatedXml = updatedXml.replace(/<a:tcPr([^>]*?)>([\s\S]*?)<\/a:tcPr>/i, (match, attrs, innerProps) => {
                    let cleanedProps = innerProps.replace(/<a:(solidFill|gradFill|blipFill|pattFill|noFill)[^>]*?(\/>|>[\s\S]*?<\/a:\1>)/gi, "");
                    return `<a:tcPr${attrs}>${fillXml}${cleanedProps}</a:tcPr>`;
                });
            } else if (/<a:tc[^>]*?>/i.test(updatedXml)) {
                updatedXml = updatedXml.replace(/<a:tc([^>]*?)>/i, `<a:tc$1><a:tcPr>${fillXml}</a:tcPr>`);
            }

            return updatedXml;
        });

        // Fail-safe cleanup: Erase any leftover {{bgColor:...}} tags in raw text runs across the slide XML
        return processedXml.replace(/\{\{bgColor:\s*[^{}\s]+\}\}\s*/g, "");
    }

    renderSingleItem(slide, tree, context) {
        // 1. Process global colors & variables on the entire cloned slide XML
        let slideXml = slide.getXml();
        slideXml = XmlSanitizer.sanitizeParagraphs(slideXml);
        slideXml = this.replaceVariables(slideXml, context);
        slide.setXml(slideXml);

        // 2. Process AST shape nodes (conditions and section shapes)
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
                    
                    // Run replaceVariables to handle item-level colors & variables
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
        // 1. Sanitize XML paragraph runs to eliminate split <a:t> boundaries
        let sanitizedText = XmlSanitizer.sanitizeParagraphs(text);

        // 2. Apply dynamic background fills (shapes & table cells)
        sanitizedText = this.applyBgColors(sanitizedText, context);

        // 3. Apply dynamic text colors
        sanitizedText = this.applyTextColors(sanitizedText, context);

        // 4. Process standard variable substitutions
        let processed = sanitizedText.replace(
            /\{\{([^{}]+)\}\}/g,
            (match, expression) => {
                const trimmed = expression.trim();
                
                // If it's a loop boundary or color directive, leave it for its specific renderer or erase it
                if (trimmed.startsWith("#") || trimmed.startsWith("/")) {
                    return match;
                }
                
                if (trimmed.startsWith("textColor:") || trimmed.startsWith("bgColor:")) {
                    return ""; // Always strip color tags from visible text
                }

                const value = this.resolve(context, trimmed) ?? this.resolve(this.data, trimmed);

                if (value === undefined || value === null) {
                    return match;
                }

                return this.escapeXml(String(value));
            }
        );

        // 5. Final safety pass: remove any leftover {{bgColor:...}} or {{textColor:...}} tags
        return processed.replace(/\{\{(bgColor|textColor):\s*[^{}\s]+\}\}\s*/g, "");
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