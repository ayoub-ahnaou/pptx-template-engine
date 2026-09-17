class ExpressionEvaluator {

    constructor(rootContext) {
        this.rootContext = rootContext;
    }


    evaluate(expression, context) {

        expression = expression.trim();


        // --------------------------------
        // Equality
        // --------------------------------

        let match = expression.match(
            /^(.+?)\s*===\s*(.+)$/
        );

        if (!match) {
            match = expression.match(
                /^(.+?)\s*==\s*(.+)$/
            );
        }

        if (match) {

            const left =
                this.resolve(
                    context,
                    match[1].trim()
                );

            const right =
                this.parseLiteral(
                    context,
                    match[2].trim()
                );

            return left === right;
        }


        // --------------------------------
        // Not equal
        // --------------------------------

        match = expression.match(
            /^(.+?)\s*!==\s*(.+)$/
        );

        if (!match) {
            match = expression.match(
                /^(.+?)\s*!=\s*(.+)$/
            );
        }

        if (match) {

            const left =
                this.resolve(
                    context,
                    match[1].trim()
                );

            const right =
                this.parseLiteral(
                    context,
                    match[2].trim()
                );

            return left !== right;
        }


        // --------------------------------
        // Greater than
        // --------------------------------

        match = expression.match(
            /^(.+?)\s*>\s*(.+)$/
        );

        if (match) {

            const left =
                this.resolve(
                    context,
                    match[1].trim()
                );

            const right =
                this.parseLiteral(
                    context,
                    match[2].trim()
                );

            return left > right;
        }


        // --------------------------------
        // Less than
        // --------------------------------

        match = expression.match(
            /^(.+?)\s*<\s*(.+)$/
        );

        if (match) {

            const left =
                this.resolve(
                    context,
                    match[1].trim()
                );

            const right =
                this.parseLiteral(
                    context,
                    match[2].trim()
                );

            return left < right;
        }


        // --------------------------------
        // Greater / equal
        // --------------------------------

        match = expression.match(
            /^(.+?)\s*>=\s*(.+)$/
        );

        if (match) {

            const left =
                this.resolve(
                    context,
                    match[1].trim()
                );

            const right =
                this.parseLiteral(
                    context,
                    match[2].trim()
                );

            return left >= right;
        }


        // --------------------------------
        // Less / equal
        // --------------------------------

        match = expression.match(
            /^(.+?)\s*<=\s*(.+)$/
        );

        if (match) {

            const left =
                this.resolve(
                    context,
                    match[1].trim()
                );

            const right =
                this.parseLiteral(
                    context,
                    match[2].trim()
                );

            return left <= right;
        }


        // --------------------------------
        // Simple value
        // --------------------------------

        return Boolean(
            this.resolve(
                context,
                expression
            )
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


    parseLiteral(context, value) {

        // String
        if (
            (value.startsWith('"') &&
             value.endsWith('"')) ||

            (value.startsWith("'") &&
             value.endsWith("'"))
        ) {

            return value.substring(
                1,
                value.length - 1
            );
        }


        // Number
        if (
            /^-?\d+(\.\d+)?$/.test(value)
        ) {

            return Number(value);
        }


        // Boolean
        if (value === "true") {
            return true;
        }

        if (value === "false") {
            return false;
        }


        // Null
        if (value === "null") {
            return null;
        }


        // Otherwise treat it as a variable
        return this.resolve(
            context,
            value
        );
    }
}


module.exports = ExpressionEvaluator;