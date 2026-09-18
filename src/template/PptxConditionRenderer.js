const ExpressionEvaluator =
    require("./ExpressionEvaluator");

class PptxConditionRenderer {

    constructor(data) {
        this.data = data;

        this.evaluator =
            new ExpressionEvaluator(data);
    }

    render(slide, condition, context) {

        const result =
            this.evaluator.evaluate(
                condition.expression,
                context
            );

        const shapes =
            result
                ? condition.children
                : condition.elseChildren;

        const selectedShapes =
            shapes
                .filter(
                    node =>
                        node.type === "shape"
                )
                .map(
                    node =>
                        node.shape
                );

        const allConditionShapes = [
            ...condition.children,
            ...condition.elseChildren
        ]
            .filter(
                node =>
                    node.type === "shape"
            )
            .map(
                node =>
                    node.shape
            );

        // Remove condition content
        for (
            const shape
            of allConditionShapes
        ) {
            slide.removeShape(shape);
        }

        // Remove markers
        slide.removeShape(
            condition.startShape
        );

        slide.removeShape(
            condition.endShape
        );

        if (condition.elseShape) {
            slide.removeShape(
                condition.elseShape
            );
        }

        // Reinsert selected shapes
        for (
            const shape
            of selectedShapes
        ) {
            slide.insertShape(shape);
        }
    }
}

module.exports =
    PptxConditionRenderer;