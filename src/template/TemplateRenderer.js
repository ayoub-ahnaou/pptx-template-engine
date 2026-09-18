class PptxLoopRenderer {

    constructor(data) {
        this.data = data;
    }

    render(slide) {

        const shapes =
            slide.getTextShapes();

        const startIndex =
            shapes.findIndex(
                shape =>
                    shape.getText().trim() ===
                    "{{#risks}}"
            );

        const endIndex =
            shapes.findIndex(
                shape =>
                    shape.getText().trim() ===
                    "{{/risks}}"
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

        const risks =
            this.data.risks || [];

        slide.removeShape(
            shapes[startIndex]
        );

        slide.removeShape(
            shapes[endIndex]
        );

        for (const shape of templateShapes) {
            slide.removeShape(shape);
        }

        const spacing = 500000;

        for (
            let i = 0;
            i < risks.length;
            i++
        ) {

            const risk =
                risks[i];

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

                text =
                    text.replace(
                        /\{\{name\}\}/g,
                        risk.name
                    );

                text =
                    text.replace(
                        /\{\{level\}\}/g,
                        risk.level
                    );

                text =
                    text.replace(
                        /\{\{status\}\}/g,
                        risk.status
                    );

                text =
                    text.replace(
                        /\{\{owner\}\}/g,
                        risk.owner
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
}

module.exports = PptxLoopRenderer;