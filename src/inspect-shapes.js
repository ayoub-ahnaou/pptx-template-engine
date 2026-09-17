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

    const slides =
        document.getSlideFiles();

    for (const slideFile of slides) {

        const xml =
            await document.getSlideXml(
                slideFile
            );

        console.log("\n==============================");
        console.log(slideFile);
        console.log("==============================");

        const shapeRegex =
            /<p:sp>([\s\S]*?)<\/p:sp>/g;

        let shapeNumber = 0;
        let match;

        while (
            (match = shapeRegex.exec(xml)) !== null
        ) {

            shapeNumber++;

            const shapeXml =
                match[1];

            const textRegex =
                /<a:t>([\s\S]*?)<\/a:t>/g;

            const texts = [];

            let textMatch;

            while (
                (textMatch =
                    textRegex.exec(shapeXml)) !== null
            ) {
                texts.push(
                    textMatch[1]
                );
            }

            const idMatch =
                shapeXml.match(
                    /<p:cNvPr\s+id="([^"]+)"\s+name="([^"]*)"/
                );

            const shapeId =
                idMatch
                    ? idMatch[1]
                    : "unknown";

            const shapeName =
                idMatch
                    ? idMatch[2]
                    : "unknown";

            console.log(
                `Shape ${shapeNumber}:`
            );

            console.log(
                `  ID: ${shapeId}`
            );

            console.log(
                `  Name: ${shapeName}`
            );

            console.log(
                `  Text:`,
                texts
            );
        }
    }
}

main().catch(error => {
    console.error(error);
});