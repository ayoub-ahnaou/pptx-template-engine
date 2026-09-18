const PptxDocument =
    require("./pptx/PptxDocument");

const PptxTemplateParser =
    require("./pptx/PptxTemplateParser");

const path =
    require("path");

const TEMPLATE_PATH =
    path.join(
        __dirname,
        "../template/template.pptx"
    );

async function main() {

    const document =
        await PptxDocument.load(
            TEMPLATE_PATH
        );

    const slide =
        await document.getSlide(
            "ppt/slides/slide3.xml"
        );

    const shapes =
        slide.getTextShapes();

    const parser =
        new PptxTemplateParser();

    const tree =
        parser.parse(shapes);

    console.dir(
        tree,
        {
            depth: 10
        }
    );
}

main().catch(error => {
    console.error(error);
});