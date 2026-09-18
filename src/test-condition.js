const fs = require("fs");
const path = require("path");

const PptxDocument =
    require("./pptx/PptxDocument");

const PptxTemplateParser =
    require("./pptx/PptxTemplateParser");

const PptxConditionRenderer =
    require("./template/PptxConditionRenderer");

const TEMPLATE_PATH =
    path.join(
        __dirname,
        "../template/template.pptx"
    );

const DATA_PATH =
    path.join(
        __dirname,
        "../data/data.json"
    );

const OUTPUT_PATH =
    path.join(
        __dirname,
        "../output/condition-test.pptx"
    );

async function main() {

    const data =
        JSON.parse(
            fs.readFileSync(
                DATA_PATH,
                "utf-8"
            )
        );

    const document =
        await PptxDocument.load(
            TEMPLATE_PATH
        );

    const slide =
        await document.getSlide(
            "ppt/slides/slide5.xml"
        );

    const parser =
        new PptxTemplateParser();

    const tree =
        parser.parse(
            slide.getTextShapes()
        );

    const condition =
        tree.find(
            node =>
                node.type === "condition"
        );

    if (!condition) {
        throw new Error(
            "No condition found"
        );
    }

    console.log(
        "Condition:",
        condition.expression
    );

    const renderer =
        new PptxConditionRenderer(data);

    renderer.render(
        slide,
        condition,
        data
    );

    document.setSlide(
        "ppt/slides/slide5.xml",
        slide
    );

    await document.save(
        OUTPUT_PATH
    );

    console.log(
        `Generated: ${OUTPUT_PATH}`
    );
}

main().catch(error => {
    console.error(error);
});