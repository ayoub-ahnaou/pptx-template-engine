const fs = require("fs");
const JSZip = require("jszip");
const PptxSlide = require("./PptxSlide");

class PptxDocument {

    constructor(zip) {
        this.zip = zip;
    }

    static async load(filePath) {
        const buffer =
            fs.readFileSync(filePath);

        const zip =
            await JSZip.loadAsync(buffer);

        return new PptxDocument(zip);
    }

    getSlideFiles() {
        return Object.keys(this.zip.files)
            .filter(file =>
                /^ppt\/slides\/slide\d+\.xml$/.test(file)
            );
    }

    async getSlideXml(slideFile) {
        const file =
            this.zip.file(slideFile);

        if (!file) {
            throw new Error(
                `Slide not found: ${slideFile}`
            );
        }

        return file.async("string");
    }

    async getSlide(slideFile) {
        const xml =
            await this.getSlideXml(slideFile);

        return new PptxSlide(xml);
    }

    setSlideXml(slideFile, xml) {
        this.zip.file(
            slideFile,
            xml
        );
    }

    setSlide(slideFile, slide) {
        this.setSlideXml(
            slideFile,
            slide.getXml()
        );
    }

    async save(filePath) {
        const buffer =
            await this.zip.generateAsync({
                type: "nodebuffer"
            });

        fs.writeFileSync(
            filePath,
            buffer
        );
    }

    async duplicateSlide(sourceSlideFile, newSlideNumber) {

        const sourceXml =
            await this.getSlideXml(sourceSlideFile);

        const newSlideFile =
            `ppt/slides/slide${newSlideNumber}.xml`;

        // Copy slide XML
        this.zip.file(
            newSlideFile,
            sourceXml
        );

        // --------------------------------------------------
        // Find presentation.xml
        // --------------------------------------------------

        const presentationFile =
            this.zip.file(
                "ppt/presentation.xml"
            );

        const presentationXml =
            await presentationFile.async("string");

        // --------------------------------------------------
        // Find presentation.xml.rels
        // --------------------------------------------------

        const presentationRelsFile =
            this.zip.file(
                "ppt/_rels/presentation.xml.rels"
            );

        let presentationRelsXml =
            await presentationRelsFile.async("string");

        // Find the relationship used by the source slide
        const sourceSlideNumber =
            Number(
                sourceSlideFile.match(
                    /slide(\d+)\.xml/
                )[1]
            );

        const sourceRelationshipRegex =
            new RegExp(
                `<Relationship[^>]+Target="slides/slide${sourceSlideNumber}\\.xml"[^>]*\\/>`
            );

        const sourceRelationship =
            presentationRelsXml.match(
                sourceRelationshipRegex
            );

        if (!sourceRelationship) {
            throw new Error(
                `Relationship for ${sourceSlideFile} not found`
            );
        }

        // Get next relationship ID
        const relationshipIds = [
            ...presentationRelsXml.matchAll(
                /Id="rId(\d+)"/g
            )
        ].map(
            match => Number(match[1])
        );

        const nextRelationshipId =
            Math.max(...relationshipIds) + 1;

        const newRelationshipId =
            `rId${nextRelationshipId}`;

        const newRelationship =
            sourceRelationship[0]
                .replace(
                    /Id="[^"]+"/,
                    `Id="${newRelationshipId}"`
                )
                .replace(
                    `slide${sourceSlideNumber}.xml`,
                    `slide${newSlideNumber}.xml`
                );

        presentationRelsXml =
            presentationRelsXml.replace(
                "</Relationships>",
                `${newRelationship}</Relationships>`
            );

        this.zip.file(
            "ppt/_rels/presentation.xml.rels",
            presentationRelsXml
        );

        // --------------------------------------------------
        // Add slide to presentation.xml
        // --------------------------------------------------

        const slideIds = [
            ...presentationXml.matchAll(
                /<p:sldId\s+id="(\d+)"\s+r:id="([^"]+)"/g
            )
        ];

        const nextSlideId =
            Math.max(
                ...slideIds.map(
                    match => Number(match[1])
                )
            ) + 1;

        const sourceSlideId =
            slideIds.find(
                match =>
                    match[2] ===
                    this.findRelationshipId(
                        sourceSlideFile,
                        presentationRelsXml
                    )
            );

        const newSlideIdElement =
            `<p:sldId id="${nextSlideId}" r:id="${newRelationshipId}"/>`;

        const slideIdListEnd =
            "</p:sldIdLst>";

        const insertPosition =
            presentationXml.indexOf(
                slideIdListEnd
            );

        const updatedPresentationXml =
            presentationXml.slice(
                0,
                insertPosition
            ) +
            newSlideIdElement +
            presentationXml.slice(
                insertPosition
            );

        this.zip.file(
            "ppt/presentation.xml",
            updatedPresentationXml
        );

        // --------------------------------------------------
        // Add Content Type entry
        // --------------------------------------------------

        const contentTypesFile =
            this.zip.file(
                "[Content_Types].xml"
            );

        let contentTypesXml =
            await contentTypesFile.async("string");

        const hasOverride =
            contentTypesXml.includes(
                `/ppt/slides/slide${newSlideNumber}.xml`
            );

        if (!hasOverride) {

            const override =
                `<Override PartName="/ppt/slides/slide${newSlideNumber}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;

            contentTypesXml =
                contentTypesXml.replace(
                    "</Types>",
                    `${override}</Types>`
                );

            this.zip.file(
                "[Content_Types].xml",
                contentTypesXml
            );
        }

        // --------------------------------------------------
        // Copy slide relationships
        // --------------------------------------------------

        const sourceRelsPath =
            `ppt/slides/_rels/slide${sourceSlideNumber}.xml.rels`;

        const newRelsPath =
            `ppt/slides/_rels/slide${newSlideNumber}.xml.rels`;

        const sourceRels =
            this.zip.file(
                sourceRelsPath
            );

        if (sourceRels) {

            const relsXml =
                await sourceRels.async("string");

            this.zip.file(
                newRelsPath,
                relsXml
            );
        }

        return newSlideFile;
    }

    findRelationshipId(
        slideFile,
        presentationRelsXml
    ) {

        const slideNumber =
            Number(
                slideFile.match(
                    /slide(\d+)\.xml/
                )[1]
            );

        const regex =
            new RegExp(
                `<Relationship[^>]+Id="([^"]+)"[^>]+Target="slides/slide${slideNumber}\\.xml"`
            );

        const match =
            presentationRelsXml.match(
                regex
            );

        if (!match) {
            throw new Error(
                `Relationship for ${slideFile} not found`
            );
        }

        return match[1];
    }
}

module.exports = PptxDocument;