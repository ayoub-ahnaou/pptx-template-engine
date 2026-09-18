const fs = require("fs");
const path = require("path");

const PptxDocument =
    require("./pptx/PptxDocument");

const PptxLoopRenderer =
    require("./template/PptxLoopRenderer");

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
        "../output/loop-test.pptx"
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
            "ppt/slides/slide3.xml"
        );

    const renderer =
        new PptxLoopRenderer(data);

    renderer.render(slide);

    document.setSlide(
        "ppt/slides/slide3.xml",
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