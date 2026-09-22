const fs = require("fs");
const JSZip = require("jszip");
const PptxSlide = require("./PptxSlide");

class PptxDocument {
    constructor(zip) {
        this.zip = zip;
    }

    static async load(filePath) {
        const buffer = fs.readFileSync(filePath);
        const zip = await JSZip.loadAsync(buffer);
        return new PptxDocument(zip);
    }

    getSlideFiles() {
        return Object.keys(this.zip.files)
            .filter(file => /^ppt\/slides\/slide\d+\.xml$/.test(file))
            .sort((a, b) => {
                const numA = parseInt(a.match(/\d+/)[0], 10);
                const numB = parseInt(b.match(/\d+/)[0], 10);
                return numA - numB;
            });
    }

    async getSlideXml(slideFile) {
        const file = this.zip.file(slideFile);
        if (!file) {
            throw new Error(`Slide not found: ${slideFile}`);
        }
        return file.async("string");
    }

    async getSlide(slideFile) {
        const xml = await this.getSlideXml(slideFile);
        return new PptxSlide(xml);
    }

    setSlideXml(slideFile, xml) {
        this.zip.file(slideFile, xml);
    }

    setSlide(slideFile, slide) {
        this.setSlideXml(slideFile, slide.getXml());
    }

    /**
     * Duplicates a slide, generates new relationship IDs, and inserts the new slide
     * directly after the source slide in ppt/presentation.xml without corrupting slide references.
     */
    async duplicateSlide(sourceSlideFile) {
        const slideFiles = this.getSlideFiles();
        const existingNumbers = slideFiles.map(f => parseInt(f.match(/\d+/)[0], 10));
        const maxNumber = Math.max(...existingNumbers);
        const newSlideNumber = maxNumber + 1;
        const targetSlideFile = `ppt/slides/slide${newSlideNumber}.xml`;

        // 1. Copy Slide XML
        const sourceXml = await this.getSlideXml(sourceSlideFile);
        this.zip.file(targetSlideFile, sourceXml);

        // 2. Copy Slide Relationships (.rels)
        const sourceRelsFile = sourceSlideFile.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
        const targetRelsFile = targetSlideFile.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
        const sourceRels = this.zip.file(sourceRelsFile);
        if (sourceRels) {
            const relsXml = await sourceRels.async("string");
            this.zip.file(targetRelsFile, relsXml);
        }

        // 3. Register in [Content_Types].xml
        let contentTypesXml = await this.zip.file("[Content_Types].xml").async("string");
        const newOverride = `<Override PartName="/${targetSlideFile}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
        if (!contentTypesXml.includes(`PartName="/${targetSlideFile}"`)) {
            contentTypesXml = contentTypesXml.replace("</Types>", `${newOverride}</Types>`);
            this.zip.file("[Content_Types].xml", contentTypesXml);
        }

        // 4. Register in ppt/_rels/presentation.xml.rels
        let presRelsXml = await this.zip.file("ppt/_rels/presentation.xml.rels").async("string");
        const rIdMatches = [...presRelsXml.matchAll(/Id="rId(\d+)"/g)];
        const maxRId = rIdMatches.reduce((max, m) => Math.max(max, parseInt(m[1], 10)), 0);
        const newRId = `rId${maxRId + 1}`;

        const newRelationship = `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newSlideNumber}.xml"/>`;
        presRelsXml = presRelsXml.replace("</Relationships>", `${newRelationship}</Relationships>`);
        this.zip.file("ppt/_rels/presentation.xml.rels", presRelsXml);

        // 5. Register in ppt/presentation.xml (Insert immediately AFTER the source slide)
        let presXml = await this.zip.file("ppt/presentation.xml").async("string");
        
        const sourceFileName = sourceSlideFile.replace("ppt/", "");
        const sourceRelMatch = presRelsXml.match(new RegExp(`<Relationship\\s+[^>]*?Id="([^"]+)"[^>]*?Target="${sourceFileName}"`));
        const sourceRId = sourceRelMatch ? sourceRelMatch[1] : null;

        const sldIdMatches = [...presXml.matchAll(/id="(\d+)"/g)];
        const maxSldId = sldIdMatches.reduce((max, m) => Math.max(max, parseInt(m[1], 10)), 255);
        const newSldId = maxSldId + 1;

        const newSldIdEntry = `<p:sldId id="${newSldId}" r:id="${newRId}"/>`;

        if (sourceRId) {
            const sourceSldPattern = new RegExp(`(<p:sldId\\s+[^>]*?r:id="${sourceRId}"[^>]*?\\/>)`);
            if (sourceSldPattern.test(presXml)) {
                presXml = presXml.replace(sourceSldPattern, `$1${newSldIdEntry}`);
            } else {
                presXml = presXml.replace("</p:sldIdLst>", `${newSldIdEntry}</p:sldIdLst>`);
            }
        } else {
            presXml = presXml.replace("</p:sldIdLst>", `${newSldIdEntry}</p:sldIdLst>`);
        }

        this.zip.file("ppt/presentation.xml", presXml);

        return targetSlideFile;
    }

    async save(filePath) {
        const buffer = await this.zip.generateAsync({ type: "nodebuffer" });
        fs.writeFileSync(filePath, buffer);
    }
}

module.exports = PptxDocument;