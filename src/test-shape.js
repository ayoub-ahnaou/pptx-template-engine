const PptxDocument =
    require("./pptx/PptxDocument");

const PptxShape =
    require("./pptx/PptxShape");

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

    const xml =
        await document.getSlideXml(
            "ppt/slides/slide3.xml"
        );

    const shapeRegex =
        /<p:sp>([\s\S]*?)<\/p:sp>/g;

    let match;

    while (
        (match = shapeRegex.exec(xml)) !== null
    ) {

        const shapeXml =
            match[0];

        if (
            shapeXml.includes(
                "{{name}}"
            )
        ) {

            const shape =
                new PptxShape(
                    shapeXml
                );

            console.log(
                "ID:",
                shape.id
            );

            console.log(
                "Name:",
                shape.name
            );

            console.log(
                "Text:",
                shape.getText()
            );

            console.log(
                "Position:",
                shape.getPosition()
            );

            console.log(
                "Size:",
                shape.getSize()
            );

            console.log(
                "\nChanging text...\n"
            );

            shape.setText(
                "Data Loss"
            );

            console.log(
                "New text:",
                shape.getText()
            );

            break;
        }
    }
}

main().catch(error => {
    console.error(error);
});