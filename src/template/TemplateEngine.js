const PptxDocument = require("../pptx/PptxDocument");
const PptxSlide = require("../pptx/PptxSlide");
const PptxTemplateParser = require("../pptx/PptxTemplateParser");
const PptxBlockRenderer = require("./PptxBlockRenderer");

class TemplateEngine {

    constructor(data) {
        this.data = data;
        this.parser = new PptxTemplateParser();
        this.blockRenderer = new PptxBlockRenderer(data);
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
        let slideFiles = doc.getSlideFiles();
        let i = 0;

        while (i < slideFiles.length) {
            const slideFile = slideFiles[i];
            const rawSlideXml = await doc.getSlideXml(slideFile);
            const slide = new PptxSlide(rawSlideXml);
            const shapes = slide.getTextShapes();

            let tree = [];
            try {
                tree = this.parser.parse(shapes);
            } catch (e) {
                tree = shapes.map(s => ({ type: "shape", shape: s }));
            }

            // Check if this slide is Slide 4 OR contains {{#eachSlide ...}}
            const isSlide4 = slideFile.endsWith("slide4.xml");
            const eachSlideNode = tree.find(node => node.type === "eachSlide") || 
                                  (isSlide4 ? tree.find(node => node.type === "section") : null);

            if (eachSlideNode) {
                const items = this.blockRenderer.resolve(this.data, eachSlideNode.expression);

                if (Array.isArray(items) && items.length > 0) {
                    let previousSlideFile = slideFile;

                    // Step A: First duplicate all target slides from the clean original XML
                    const slideTargetFiles = [slideFile];
                    for (let itemIdx = 1; itemIdx < items.length; itemIdx++) {
                        const targetFile = await doc.duplicateSlide(previousSlideFile);
                        slideTargetFiles.push(targetFile);
                        previousSlideFile = targetFile;
                    }

                    // Step B: Render each item in isolation
                    for (let itemIdx = 0; itemIdx < items.length; itemIdx++) {
                        const itemContext = items[itemIdx];
                        const targetFile = slideTargetFiles[itemIdx];

                        // Always fetch fresh XML string from ZIP to avoid cross-contamination
                        const freshXml = await doc.getSlideXml(targetFile);
                        const targetSlide = new PptxSlide(freshXml);
                        const targetShapes = targetSlide.getTextShapes();
                        const targetTree = this.parser.parse(targetShapes);

                        // Map eachSlide/section AST nodes to single item section context
                        const adaptedTree = targetTree.map(node => {
                            if (
                                node.type === "eachSlide" ||
                                node.expression === eachSlideNode.expression
                            ) {
                                return { ...node, type: "section" };
                            }
                            return node;
                        });

                        // Render single item (Risk 1 on Slide 4, Risk 2 on Slide 5/6, etc.)
                        this.blockRenderer.renderSingleItem(targetSlide, adaptedTree, itemContext);
                        this.processGlobalVariables(targetSlide, itemContext);

                        // Save updated XML to the archive
                        doc.setSlide(targetFile, targetSlide);
                    }

                    // Refresh slide list to account for inserted duplicate slides
                    slideFiles = doc.getSlideFiles();
                    i += items.length;
                    continue;
                }
            }

            // Standard in-slide loop processing for other slides (e.g., Slide 3)
            const hasSection = tree.some(node => node.type === "section");
            if (hasSection) {
                this.blockRenderer.render(slide, tree);
            }

            const hasCondition = tree.some(node => node.type === "condition");
            if (hasCondition) {
                for (const node of tree) {
                    if (node.type === "condition") {
                        const conditionRenderer = new (require("./PptxConditionRenderer"))(this.data);
                        conditionRenderer.render(slide, node, this.data);
                    }
                }
            }

            this.processGlobalVariables(slide, this.data);
            doc.setSlide(slideFile, slide);

            i++;
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