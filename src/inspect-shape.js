const PptxDocument =
    require("./pptx/PptxDocument");

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

    const slideFile =
        "ppt/slides/slide3.xml";

    const xml =
        await document.getSlideXml(
            slideFile
        );

    const shapeRegex =
        /<p:sp>([\s\S]*?)<\/p:sp>/g;

    let match;
    let shapeNumber = 0;

    while (
        (match = shapeRegex.exec(xml)) !== null
    ) {

        shapeNumber++;

        const shapeXml =
            match[0];

        if (
            shapeXml.includes(
                "{{name}}"
            )
        ) {

            console.log(
                `Found {{name}} in Shape ${shapeNumber}`
            );

            console.log(
                "\n========== SHAPE XML ==========\n"
            );

            console.log(shapeXml);

            console.log(
                "\n========== END ==========\n"
            );

            break;
        }
    }
}

main().catch(error => {
    console.error(error);
});