const ExpressionEvaluator =
    require("./ExpressionEvaluator");

class PptxBlockRenderer {

    constructor(data) {
        this.data = data;

        this.evaluator =
            new ExpressionEvaluator(data);
    }

    render(slide, tree) {

        for (const node of tree) {

            if (node.type === "section") {
                this.renderSection(
                    slide,
                    node
                );
            }
        }
    }

    renderSection(slide, section) {

        const items =
            this.resolve(
                this.data,
                section.expression
            );

        if (!Array.isArray(items)) {
            throw new Error(
                `Section '${section.expression}' must be an array`
            );
        }

        const allShapes =
            this.collectShapes(section);

        // Remove original template shapes
        for (const shape of allShapes) {
            slide.removeShape(shape);
        }

        slide.removeShape(
            section.startShape
        );

        slide.removeShape(
            section.endShape
        );

        const spacing = 900000;

        for (
            let index = 0;
            index < items.length;
            index++
        ) {

            const context =
                items[index];

            const shapes =
                this.renderNodes(
                    section.children,
                    context
                );

            for (const templateShape of shapes) {

                const clone =
                    templateShape.clone(
                        slide.getNextShapeId()
                    );

                let text =
                    clone.getText();

                text =
                    this.replaceVariables(
                        text,
                        context
                    );

                clone.setText(text);

                const position =
                    clone.getPosition();

                if (position) {
                    clone.setPosition(
                        position.x,
                        position.y +
                        (index * spacing)
                    );
                }

                slide.insertShape(
                    clone
                );
            }
        }
    }

    /**
     * Renders a block tree using a single item context (for slide duplication).
     * Replaces placeholders, evaluates nested conditions, cleans up unselected shapes,
     * and maintains original shape positioning.
     */
    renderSingleItem(slide, tree, context) {
        for (const node of tree) {
            if (node.type === "section") {

                // Get selected branch shapes (handling conditions if present)
                const selectedShapes = this.renderNodes(node.children, context);
                const allShapes = this.collectShapes(node);

                // 1. Remove shapes belonging to unselected condition branches
                for (const shape of allShapes) {
                    if (!selectedShapes.includes(shape)) {
                        slide.removeShape(shape);
                    }
                }

                // 2. Perform variable replacement on active shapes & update slide XML
                for (const shape of selectedShapes) {
                    const originalXml = shape.getXml();
                    const currentText = shape.getText();
                    const updatedText = this.replaceVariables(currentText, context);

                    shape.setText(updatedText);

                    // Update slide XML string by replacing the old shape XML with updated shape XML
                    slide.xml = slide.xml.replace(originalXml, shape.getXml());
                }

                // 3. Remove loop control markers ({{#risks}}, {{/risks}}, {{else}})
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

                result.push(
                    node.shape
                );

                continue;
            }

            if (node.type === "condition") {

                const value =
                    this.evaluator.evaluate(
                        node.expression,
                        context
                    );

                const selected =
                    value
                        ? node.children
                        : node.elseChildren;

                result.push(
                    ...this.renderNodes(
                        selected,
                        context
                    )
                );
            }
        }

        return result;
    }

    collectShapes(section) {

        const result = [];

        const collect =
            nodes => {

                for (const node of nodes) {

                    if (node.type === "shape") {
                        result.push(node.shape);
                    }

                    else if (
                        node.type === "condition"
                    ) {
                        collect(node.children);
                        collect(node.elseChildren);

                        if (node.startShape) {
                            result.push(
                                node.startShape
                            );
                        }

                        if (node.endShape) {
                            result.push(
                                node.endShape
                            );
                        }

                        if (node.elseShape) {
                            result.push(
                                node.elseShape
                            );
                        }
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

                const value =
                    this.resolve(
                        context,
                        expression.trim()
                    );

                if (
                    value === undefined ||
                    value === null
                ) {
                    return match;
                }

                return String(value);
            }
        );
    }

    resolve(context, expression) {

        const parts =
            expression.split(".");

        let current = context;

        for (const part of parts) {

            if (
                current === null ||
                current === undefined
            ) {
                return undefined;
            }

            if (
                typeof current !== "object" ||
                !(part in current)
            ) {
                return undefined;
            }

            current =
                current[part];
        }

        return current;
    }
}

module.exports =
    PptxBlockRenderer;