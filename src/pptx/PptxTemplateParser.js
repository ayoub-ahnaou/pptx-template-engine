class PptxTemplateParser {

    parse(shapes) {

        const root = [];

        const stack = [
            {
                type: "root",
                children: root,
                elseChildren: [],
                inElse: false
            }
        ];

        for (const shape of shapes) {

            const text =
                shape.getText().trim();

            const current =
                stack[stack.length - 1];

            const target =
                current.inElse
                    ? current.elseChildren
                    : current.children;

            const sectionStart =
                text.match(
                    /^\{\{#(.+)\}\}$/
                );

            if (sectionStart) {

                const expression =
                    sectionStart[1].trim();

                if (
                    expression.startsWith("if ")
                ) {

                    const condition = {

                        type: "condition",

                        expression:
                            expression
                                .substring(3)
                                .trim(),

                        startShape:
                            shape,

                        children: [],

                        elseChildren: [],

                        inElse: false
                    };

                    target.push(
                        condition
                    );

                    stack.push(
                        condition
                    );

                } else if (expression.startsWith("eachSlide ")) {

                    const slideSection = {

                        type: "eachSlide",

                        expression:
                            expression
                                .substring(10)
                                .trim(),

                        startShape:
                            shape,

                        children: [],

                        elseChildren: [],

                        inElse: false
                    };

                    target.push(
                        slideSection
                    );

                    stack.push(
                        slideSection
                    );

                } else {

                    const section = {

                        type: "section",

                        expression,

                        startShape:
                            shape,

                        children: [],

                        elseChildren: [],

                        inElse: false
                    };

                    target.push(
                        section
                    );

                    stack.push(
                        section
                    );
                }

                continue;
            }

            const sectionEnd =
                text.match(
                    /^\{\{\/(.+)\}\}$/
                );

            if (sectionEnd) {

                const closingName =
                    sectionEnd[1].trim();

                const block =
                    stack.pop();

                if (!block) {
                    throw new Error(
                        `Unexpected closing block: ${closingName}`
                    );
                }

                let expected = block.expression;
                if (block.type === "condition") expected = "if";
                if (block.type === "eachSlide") expected = `eachSlide ${block.expression}`;

                if (
                    closingName !== expected && closingName !== block.expression
                ) {
                    throw new Error(
                        `Mismatched block: ${expected} != ${closingName}`
                    );
                }

                block.endShape =
                    shape;

                continue;
            }

            if (
                text === "{{else}}"
            ) {

                if (
                    current.type !== "section" &&
                    current.type !== "condition" &&
                    current.type !== "eachSlide"
                ) {
                    throw new Error(
                        "{{else}} must be inside a block"
                    );
                }

                current.inElse = true;

                current.elseShape =
                    shape;

                continue;
            }

            target.push({
                type: "shape",
                shape
            });
        }

        if (stack.length !== 1) {
            throw new Error(
                "Unclosed PowerPoint template block"
            );
        }

        return root;
    }
}

module.exports = PptxTemplateParser;