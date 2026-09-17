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
                /^ppt\/slides\/slide\d+\.xml$/
                    .test(file)
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
            await this.getSlideXml(
                slideFile
            );

        return new PptxSlide(xml);
    }

    setSlideXml(slideFile, xml) {

        this.zip.file(
            slideFile,
            xml
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
}

module.exports = PptxDocument;