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

    const slide =
        await document.getSlide(
            "ppt/slides/slide3.xml"
        );

    const shapes =
        slide.getTextShapes();

    const nameShape =
        shapes.find(
            shape =>
                shape.getText() ===
                "{{name}}"
        );

    if (!nameShape) {
        throw new Error(
            "{{name}} shape not found"
        );
    }

    console.log(
        "Original:"
    );

    console.log(
        "ID:",
        nameShape.id
    );

    console.log(
        "Text:",
        nameShape.getText()
    );

    console.log(
        "Position:",
        nameShape.getPosition()
    );

    const newId =
        slide.getNextShapeId();

    const clone =
        nameShape.clone(
            newId
        );

    clone.setText(
        "Data Loss"
    );

    const position =
        clone.getPosition();

    clone.setPosition(
        position.x,
        position.y +
            position.y * 0.25
    );

    slide.insertShape(
        clone
    );

    console.log(
        "\nClone:"
    );

    console.log(
        "ID:",
        clone.id
    );

    console.log(
        "Text:",
        clone.getText()
    );

    console.log(
        "Position:",
        clone.getPosition()
    );

    console.log(
        "\nTotal shapes:",
        slide.getShapes().length
    );
}

main().catch(error => {
    console.error(error);
});