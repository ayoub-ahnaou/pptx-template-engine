const ExpressionEvaluator = require("./ExpressionEvaluator");

class TemplateRenderer {

    constructor(data) {
        this.data = data;
        this.evaluator = new ExpressionEvaluator(data);
    }


    render(nodes, context = this.data) {

        let result = "";


        for (const node of nodes) {


            // -------------------------
            // TEXT
            // -------------------------

            if (node.type === "text") {

                result += node.value;

            }


            // -------------------------
            // VARIABLE
            // -------------------------

            else if (
                node.type === "variable"
            ) {

                const value =
                    this.resolve(
                        context,
                        node.value
                    );


                if (
                    value === undefined ||
                    value === null
                ) {

                    console.warn(
                        `Variable not found: ${node.value}`
                    );

                    result +=
                        `{{${node.value}}}`;

                } else {

                    result +=
                        String(value);
                }
            }


            // -------------------------
            // SECTION
            // -------------------------

            else if (
                node.type === "section"
            ) {

                result +=
                    this.renderSection(
                        node,
                        context
                    );
            }

            else if (
                node.type === "condition"
            ) {

                result +=
                    this.renderCondition(
                        node,
                        context
                    );
            }
        }


        return result;
    }


    renderSection(section, context) {

        const value =
            this.resolve(
                context,
                section.expression
            );


        // Array → loop
        if (Array.isArray(value)) {

            return value
                .map(item =>
                    this.render(
                        section.children,
                        item
                    )
                )
                .join("");
        }


        // Truthy → render children
        if (value) {

            return this.render(
                section.children,
                context
            );
        }


        // False → render else
        return this.render(
            section.elseChildren,
            context
        );
    }

    
    renderCondition(condition, context) {

        const result =
            this.evaluator.evaluate(
                condition.expression,
                context
            );


        if (result) {

            return this.render(
                condition.children,
                context
            );
        }


        return this.render(
            condition.elseChildren,
            context
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


module.exports = TemplateRenderer;