const PptxDocument =
    require("./pptx/PptxDocument");

const path = require("path");

const TEMPLATE_PATH =
    path.join(__dirname, "../template/template.pptx");

async function main() {

    const document =
        await PptxDocument.load(
            TEMPLATE_PATH
        );

    const slides =
        document.getSlideFiles();

    console.log("Slides:");
    console.log(slides);

    for (const slideFile of slides) {

        console.log("\n==============================");
        console.log(slideFile);
        console.log("==============================");

        const xml =
            await document.getSlideXml(
                slideFile
            );

        const textRegex =
            /<a:t>([\s\S]*?)<\/a:t>/g;

        let match;

        while (
            (match = textRegex.exec(xml)) !== null
        ) {
            console.log(
                "TEXT:",
                match[1]
            );
        }
    }
}

main().catch(error => {
    console.error(error);
});