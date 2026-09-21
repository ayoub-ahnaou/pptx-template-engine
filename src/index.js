const path = require("path");
const fs = require("fs");
const TemplateEngine = require("./template/TemplateEngine");

async function run() {
    const templatePath = path.join(__dirname, "../template/template.pptx");
    const dataPath = path.join(__dirname, "../data/data.json");
    const outputPath = path.join(__dirname, "../output/final-report.pptx");

    const data = JSON.parse(fs.readFileSync(dataPath, "utf8"));
    const engine = new TemplateEngine(data);

    console.log("Generating presentation...");
    await engine.renderToFile(templatePath, outputPath);
    console.log(`Report generated successfully at: ${outputPath}`);
}

run().catch(console.error);