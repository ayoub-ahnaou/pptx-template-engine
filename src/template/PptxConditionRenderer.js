const ExpressionEvaluator = require("./ExpressionEvaluator");

class PptxConditionRenderer {

    constructor(data) {
        this.data = data;
        this.evaluator = new ExpressionEvaluator(data);
    }

    render(slide, condition, context = this.data) {
        // 1. Evaluate condition
        const isTrue = this.evaluator.evaluate(condition.expression, context);

        // 2. Identify shapes to keep vs remove
        const selectedNodes = isTrue ? condition.children : condition.elseChildren;
        const unselectedNodes = isTrue ? condition.elseChildren : condition.children;

        const selectedShapes = selectedNodes
            .filter(node => node.type === "shape")
            .map(node => node.shape);

        const unselectedShapes = unselectedNodes
            .filter(node => node.type === "shape")
            .map(node => node.shape);

        // 3. Remove shapes belonging to the unselected branch
        for (const shape of unselectedShapes) {
            slide.removeShape(shape);
        }

        // 4. Remove block tag shapes ({{#if ...}}, {{else}}, {{/if}})
        if (condition.startShape) slide.removeShape(condition.startShape);
        if (condition.elseShape) slide.removeShape(condition.elseShape);
        if (condition.endShape) slide.removeShape(condition.endShape);

        // 5. Replace any remaining placeholders in selected shapes and update slide XML
        for (const shape of selectedShapes) {
            const currentText = shape.getText();

            if (currentText.includes("{{")) {
                const updatedText = currentText.replace(
                    /\{\{([^{}]+)\}\}/g,
                    (match, expr) => {
                        const trimmed = expr.trim();
                        if (trimmed.startsWith("#") || trimmed.startsWith("/")) return match;
                        const val = this.evaluator.resolve(context, trimmed) ?? 
                                    this.evaluator.resolve(this.data, trimmed);
                        return val !== undefined && val !== null ? String(val) : match;
                    }
                );

                const oldShapeXml = shape.getXml();
                shape.setText(updatedText);
                slide.xml = slide.xml.replace(oldShapeXml, shape.getXml());
            }
        }
    }
}

module.exports = PptxConditionRenderer;