const PptxDocument =
    require("./pptx/PptxDocument");

const PptxSlide =
    require("./pptx/PptxSlide");

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

    const slideFiles =
        document.getSlideFiles();

    for (const slideFile of slideFiles) {

        const xml =
            await document.getSlideXml(
                slideFile
            );

        const slide =
            new PptxSlide(xml);

        console.log(
            `\n${slideFile}`
        );

        console.log(
            slide.getTexts()
        );
    }
}

main().catch(error => {
    console.error(error);
});