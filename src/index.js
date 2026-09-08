const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const { XMLParser, XMLBuilder } = require("fast-xml-parser");

const TEMPLATE_PATH = path.join(__dirname, "../template/template.pptx");
const DATA_PATH = path.join(__dirname, "../data/data.json");
const OUTPUT_PATH = path.join(__dirname, "../output/generated.pptx");

async function generatePptx() {
    console.log("Loading files...");

    // 1. Read JSON data
    const data = JSON.parse(
        fs.readFileSync(DATA_PATH, "utf-8")
    );

    // 2. Read PPTX
    const pptxBuffer = fs.readFileSync(TEMPLATE_PATH);

    // 3. Open PPTX as ZIP
    const zip = await JSZip.loadAsync(pptxBuffer);

    // 4. Find PowerPoint slide XML files
    const slideFiles = Object.keys(zip.files).filter(file =>
        file.startsWith("ppt/slides/slide") &&
        file.endsWith(".xml")
    );

    console.log(`Found ${slideFiles.length} slides.`);

    // 5. Process every slide
    for (const slideFile of slideFiles) {

        let xml = await zip.file(slideFile).async("string");

        xml = replaceVariables(xml, data);

        zip.file(slideFile, xml);
    }

    // 6. Generate new PPTX
    const outputBuffer = await zip.generateAsync({
        type: "nodebuffer"
    });

    // 7. Save result
    fs.writeFileSync(OUTPUT_PATH, outputBuffer);

    console.log(`Generated: ${OUTPUT_PATH}`);
}


/**
 * Replace {{variable}} placeholders
 */
function replaceVariables(xml, data) {

    return xml.replace(/\{\{([^}]+)\}\}/g, (match, expression) => {

        const value = getValue(data, expression.trim());

        if (value === undefined || value === null) {
            console.warn(`Variable not found: ${expression}`);
            return match;
        }

        return String(value);
    });
}


/**
 * Resolve nested properties.
 *
 * Example:
 *
 * getValue(data, "project.name")
 *
 * returns:
 *
 * data.project.name
 */
function getValue(data, expression) {

    const parts = expression.split(".");

    let value = data;

    for (const part of parts) {

        if (value === undefined || value === null) {
            return undefined;
        }

        value = value[part];
    }

    return value;
}


generatePptx().catch(error => {
    console.error("Generation failed:");
    console.error(error);
});
