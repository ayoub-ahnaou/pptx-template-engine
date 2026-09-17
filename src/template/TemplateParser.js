class TemplateParser {

    parse(template) {

        const tokens = this.tokenize(template);

        return this.buildTree(tokens);
    }


    /**
     * Convert the template into tokens.
     */
    tokenize(template) {

        const regex = /\{\{([\s\S]*?)\}\}/g;

        const tokens = [];

        let lastIndex = 0;
        let match;


        while ((match = regex.exec(template)) !== null) {

            // Text before {{...}}
            if (match.index > lastIndex) {

                tokens.push({
                    type: "text",
                    value: template.substring(
                        lastIndex,
                        match.index
                    )
                });
            }


            const expression =
                match[1].trim();


            // Condition start
            if (expression.startsWith("#")) {

                const value = expression.substring(1).trim();

                if (value.startsWith("if ")) {

                    tokens.push({
                        type: "condition_start",
                        expression:
                            value.substring(3).trim()
                    });

                } else {

                    tokens.push({
                        type: "section_start",
                        value
                    });
                }

            }

            // Condition / section end
            else if (expression.startsWith("/")) {

                tokens.push({
                    type: "section_end",
                    value: expression.substring(1).trim()
                });

            }

            // Else
            else if (expression === "else") {

                tokens.push({
                    type: "else"
                });

            }

            // Variable
            else {

                tokens.push({
                    type: "variable",
                    value: expression
                });
            }


            lastIndex =
                regex.lastIndex;
        }


        // Remaining text
        if (lastIndex < template.length) {

            tokens.push({
                type: "text",
                value: template.substring(
                    lastIndex
                )
            });
        }


        return tokens;
    }


    /**
     * Convert tokens into a tree.
     */
    buildTree(tokens) {

        const root = [];

        const stack = [
            {
                type: "root",
                children: root,
                elseChildren: [],
                inElse: false
            }
        ];


        for (const token of tokens) {

            const current =
                stack[stack.length - 1];


            // -----------------------------
            // TEXT
            // -----------------------------

            if (token.type === "text") {

                const target =
                    current.inElse
                        ? current.elseChildren
                        : current.children;

                target.push(token);
            }


            // -----------------------------
            // VARIABLE
            // -----------------------------

            else if (
                token.type === "variable"
            ) {

                const target =
                    current.inElse
                        ? current.elseChildren
                        : current.children;

                target.push(token);
            }


            // -----------------------------
            // LOOP
            // -----------------------------

            else if (
                token.type === "section_start"
            ) {

                const section = {

                    type: "section",

                    expression:
                        token.value,

                    children: [],

                    elseChildren: [],

                    inElse: false
                };


                const target =
                    current.inElse
                        ? current.elseChildren
                        : current.children;

                target.push(section);

                stack.push(section);
            }


            // -----------------------------
            // CONDITION
            // -----------------------------

            else if (
                token.type === "condition_start"
            ) {

                const condition = {

                    type: "condition",

                    expression:
                        token.expression,

                    children: [],

                    elseChildren: [],

                    inElse: false
                };


                const target =
                    current.inElse
                        ? current.elseChildren
                        : current.children;

                target.push(condition);

                stack.push(condition);
            }


            // -----------------------------
            // ELSE
            // -----------------------------

            else if (
                token.type === "else"
            ) {

                if (
                    current.type !== "section" &&
                    current.type !== "condition"
                ) {

                    throw new Error(
                        "{{else}} can only be used inside a section or condition"
                    );
                }


                current.inElse = true;
            }


            // -----------------------------
            // END
            // -----------------------------

            else if (
                token.type === "section_end"
            ) {

                const section =
                    stack.pop();


                if (
                    !section ||
                    (
                        section.type !== "section" &&
                        section.type !== "condition"
                    )
                ) {

                    throw new Error(
                        `Unexpected closing section: ${token.value}`
                    );
                }


                const expected =
                    section.type === "condition"
                        ? "if"
                        : section.expression;


                if (
                    token.value !== expected
                ) {

                    throw new Error(
                        `Mismatched section: ` +
                        `${expected} != ${token.value}`
                    );
                }
            }
        }


        if (stack.length !== 1) {

            throw new Error(
                "Unclosed template section"
            );
        }


        return root;
    }
}


module.exports = TemplateParser;