const fs = require("fs");
const path = require("path");

const PptxDocument =
    require("./pptx/PptxDocument");

const TemplateEngine =
    require("./template/TemplateEngine");


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
        "../output/generated-report.pptx"
    );


async function main() {

    console.log("Loading JSON...");

    const data = JSON.parse(
        fs.readFileSync(
            DATA_PATH,
            "utf-8"
        )
    );


    console.log("Loading PowerPoint...");

    const document =
        await PptxDocument.load(
            TEMPLATE_PATH
        );


    console.log(
        `Found ${document.getSlideFiles().length} slides.`
    );


    const engine =
        new TemplateEngine(data);


    for (
        const slideFile
        of document.getSlideFiles()
    ) {

        console.log(
            `Processing ${slideFile}...`
        );

        const xml =
            await document.getSlideXml(
                slideFile
            );

        const processedXml =
            engine.process(xml);

        document.setSlideXml(
            slideFile,
            processedXml
        );
    }


    await document.save(
        OUTPUT_PATH
    );


    console.log(
        `Generated: ${OUTPUT_PATH}`
    );
}


main().catch(error => {

    console.error(
        "Generation failed:"
    );

    console.error(error);

    process.exit(1);
});