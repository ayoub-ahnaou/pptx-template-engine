const PptxDocument = require("../pptx/PptxDocument");
const PptxSlide = require("../pptx/PptxSlide");
const PptxTemplateParser = require("../pptx/PptxTemplateParser");
const PptxBlockRenderer = require("./PptxBlockRenderer");
const PptxConditionRenderer = require("./PptxConditionRenderer");

class TemplateEngine {

    constructor(data) {
        this.data = data;
        this.parser = new PptxTemplateParser();
        this.blockRenderer = new PptxBlockRenderer(data);
        this.conditionRenderer = new PptxConditionRenderer(data);
    }

    /**
     * Process an entire PPTX document and save to destination path.
     */
    async renderToFile(inputPath, outputPath) {
        const doc = await PptxDocument.load(inputPath);
        await this.renderDocument(doc);
        await doc.save(outputPath);
    }

    /**
     * Main pipeline running across all slides in presentation order.
     */
    async renderDocument(doc) {
        const initialSlideFiles = doc.getSlideFiles();

        for (let idx = 0; idx < initialSlideFiles.length; idx++) {
            const slideFile = initialSlideFiles[idx];
            const rawSlideXml = await doc.getSlideXml(slideFile);
            const slide = new PptxSlide(rawSlideXml);
            let shapes = slide.getTextShapes();

            let tree = [];
            try {
                tree = this.parser.parse(shapes);
            } catch (e) {
                tree = shapes.map(s => ({ type: "shape", shape: s }));
            }

            // 1. DYNAMIC SLIDE DUPLICATION LOOP ({{#eachSlide array}})
            const eachSlideNode = tree.find(node => node.type === "eachSlide");

            if (eachSlideNode) {
                const items = this.blockRenderer.resolve(this.data, eachSlideNode.expression);

                if (Array.isArray(items) && items.length > 0) {
                    let previousSlideFile = slideFile;
                    const slideTargetFiles = [slideFile];

                    for (let itemIdx = 1; itemIdx < items.length; itemIdx++) {
                        const targetFile = await doc.duplicateSlide(previousSlideFile);
                        slideTargetFiles.push(targetFile);
                        previousSlideFile = targetFile;
                    }

                    for (let itemIdx = 0; itemIdx < items.length; itemIdx++) {
                        const itemContext = items[itemIdx];
                        const targetFile = slideTargetFiles[itemIdx];

                        const freshXml = await doc.getSlideXml(targetFile);
                        const targetSlide = new PptxSlide(freshXml);
                        const targetShapes = targetSlide.getTextShapes();
                        const targetTree = this.parser.parse(targetShapes);

                        const adaptedTree = targetTree.map(node => {
                            if (node.type === "eachSlide") {
                                return { ...node, type: "section" };
                            }
                            return node;
                        });

                        this.blockRenderer.renderSingleItem(targetSlide, adaptedTree, itemContext);
                        this.processGlobalVariables(targetSlide, itemContext);

                        doc.setSlide(targetFile, targetSlide);
                    }

                    continue;
                }
            }

            // 2. RENDER ALL SLIDES (TABLES, SHAPES, TITLES, BG & TEXT COLORS)
            // Always run blockRenderer so global placeholders, titles, and color tags are processed
            this.blockRenderer.render(slide, tree);
            
            // Refresh shapes and AST tree from updated slide XML after block rendering
            shapes = slide.getTextShapes();
            try {
                tree = this.parser.parse(shapes);
            } catch (e) {
                tree = shapes.map(s => ({ type: "shape", shape: s }));
            }

            // 3. STANDALONE CONDITIONS ({{#if condition}})
            const conditionNodes = tree.filter(node => node.type === "condition");
            if (conditionNodes.length > 0) {
                for (const conditionNode of conditionNodes) {
                    this.conditionRenderer.render(slide, conditionNode, this.data);
                }

                // Refresh shapes again after condition removals
                shapes = slide.getTextShapes();
            }

            // 4. GLOBAL VARIABLE REPLACEMENT (Fallback for remaining project.name, etc.)
            this.processGlobalVariables(slide, this.data);

            // Commit final slide XML back to presentation ZIP
            doc.setSlide(slideFile, slide);
        }
    }

    /**
     * Replace simple variables like {{project.name}} across all shapes in a slide.
     */
    processGlobalVariables(slide, context) {
        const shapes = slide.getTextShapes();

        for (const shape of shapes) {
            const originalXml = shape.getXml();
            const currentText = shape.getText();

            if (!currentText.includes("{{")) continue;

            const updatedText = currentText.replace(
                /\{\{([^{}]+)\}\}/g,
                (match, expression) => {
                    const trimmed = expression.trim();

                    if (trimmed.startsWith("#") || trimmed.startsWith("/")) {
                        return match;
                    }

                    const val = this.blockRenderer.resolve(context, trimmed) ?? 
                                this.blockRenderer.resolve(this.data, trimmed);

                    return val !== undefined && val !== null ? String(val) : match;
                }
            );

            if (updatedText !== currentText) {
                shape.setText(updatedText);
                slide.xml = slide.xml.replace(originalXml, shape.getXml());
            }
        }
    }
}

module.exports = TemplateEngine;