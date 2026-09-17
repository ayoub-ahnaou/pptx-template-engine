const TemplateParser =
    require("./template/TemplateParser");

const TemplateRenderer =
    require("./template/TemplateRenderer");


const data = {

    project: {
        name: "GRC Platform",
        status: "Active"
    },

    risks: [
        {
            name: "Data Loss",
            level: "Critical"
        },
        {
            name: "Unauthorized Access",
            level: "High"
        }
    ]
};


const template = `

Project: {{project.name}}

{{#if project.status == "Active"}}

STATUS: The project is active.

{{else}}

STATUS: The project is not active.

{{/if}}


RISKS:

{{#risks}}

- {{name}}
  Level: {{level}}

{{/risks}}

`;


const parser =
    new TemplateParser();

const tree =
    parser.parse(template);


const renderer =
    new TemplateRenderer(data);


const result =
    renderer.render(tree);


console.log(result);