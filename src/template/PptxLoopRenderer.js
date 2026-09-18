const ExpressionEvaluator =
    require("./ExpressionEvaluator");

class PptxLoopRenderer {

    constructor(data) {
        this.data = data;

        this.evaluator =
            new ExpressionEvaluator(data);
    }

    render(slide, block) {

        const shapes =
            slide.getTextShapes();

        const startIndex =
            shapes.findIndex(
                shape =>
                    shape === block.startShape
            );

        const endIndex =
            shapes.findIndex(
                shape =>
                    shape === block.endShape
            );

        if (
            startIndex === -1 ||
            endIndex === -1
        ) {
            return;
        }

        const templateShapes =
            shapes.slice(
                startIndex + 1,
                endIndex
            );

        const items =
            this.resolve(
                this.data,
                block.expression
            );

        if (!Array.isArray(items)) {
            throw new Error(
                `Loop '${block.expression}' must resolve to an array`
            );
        }

        // Remove the original template block
        slide.removeShape(
            block.startShape
        );

        slide.removeShape(
            block.endShape
        );

        for (
            const shape
            of templateShapes
        ) {
            slide.removeShape(shape);
        }

        const spacing = 900000;

        for (
            let i = 0;
            i < items.length;
            i++
        ) {

            const item =
                items[i];

            for (
                const templateShape
                of templateShapes
            ) {

                const clone =
                    templateShape.clone(
                        slide.getNextShapeId()
                    );

                let text =
                    clone.getText();

                // Ignore template control tags
                if (
                    text.includes("{{#if") ||
                    text === "{{else}}" ||
                    text === "{{/if}}"
                ) {
                    continue;
                }

                text =
                    this.replaceVariables(
                        text,
                        item
                    );

                clone.setText(text);

                const position =
                    clone.getPosition();

                clone.setPosition(
                    position.x,
                    position.y +
                    (i * spacing)
                );

                slide.insertShape(clone);
            }
        }
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

        for (
            const part
            of parts
        ) {

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
    PptxLoopRenderer;