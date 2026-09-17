class TemplateEngine {

    constructor(data) {
        this.data = data;
    }


    /**
     * Process a slide XML.
     */
    process(xml) {

        return xml.replace(
            /\{\{([^{}]+)\}\}/g,
            (match, expression) => {

                expression = expression.trim();

                // Template instructions are not implemented yet.
                if (
                    expression.startsWith("#") ||
                    expression.startsWith("/")
                ) {
                    return match;
                }

                const value =
                    this.resolve(expression);

                if (value === undefined) {

                    console.warn(
                        `Variable not found: {{${expression}}}`
                    );

                    return match;
                }

                return this.escapeXml(
                    String(value)
                );
            }
        );
    }


    /**
     * Resolve:
     *
     * project.name
     *
     * project.manager
     *
     * company.name
     */
    resolve(expression) {

        const parts = expression.split(".");

        let current = this.data;

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

            current = current[part];
        }

        return current;
    }


    /**
     * Escape XML special characters.
     */
    escapeXml(value) {

        return value
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;");
    }
}


module.exports = TemplateEngine;