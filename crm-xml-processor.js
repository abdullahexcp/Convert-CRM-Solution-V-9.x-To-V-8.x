#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const { DOMParser, XMLSerializer } = require('xmldom');
const AdmZip = require('adm-zip');

const TARGET_XML_ENTRY = 'customizations.xml';
const SOLUTION_XML_ENTRY = 'solution.xml';

class CRMXMLProcessor {
    constructor(configPath) {
        this.config = this.loadConfig(configPath);
        this.parser = new DOMParser();
        this.serializer = new XMLSerializer();
        const { target, source } = this.loadEntityTypeCodes();
        this.targetCodes = target;
        this.sourceCodes = source;
        this.summary = this.freshSummary();
    }

    freshSummary() {
        return {
            startTime: Date.now(),
            mode: null,                  // 'zip' | 'xml'
            inputFile: null,
            backupFile: null,
            xmlEntryName: null,          // path inside the zip
            originalSize: 0,
            finalSize: 0,
            originalXmlSize: 0,
            finalXmlSize: 0,
            elementsRemoved: {},         // { tagName: count }
            attributesRemoved: {},       // { 'tagName.attr': count }
            objectTypeCodes: {
                added: 0,
                fromTarget: 0,
                fromSource: 0,
                placeholders: 0,
                newEntities: [],          // in source only (using source OTC as seed)
                missingEntities: []       // in neither CSV (used '##')
            },
            solutionXml: {
                processed: false,
                notFound: false,
                changes: []               // [{ attr, from, to }]
            }
        };
    }

    loadConfig(configPath) {
        try {
            const configContent = fs.readFileSync(configPath, 'utf8');
            return JSON.parse(configContent);
        } catch (error) {
            console.error(`Error loading config file: ${error.message}`);
            process.exit(1);
        }
    }

    // Supports two config shapes:
    //   (new)    "entityTypeCodes": { "target": "path/to/target.csv", "source": "path/to/source.csv" }
    //   (legacy) "entityTypeCodesFile": "path/to/one.csv"   // treated as target
    loadEntityTypeCodes() {
        const paths = { target: null, source: null };
        if (this.config.entityTypeCodes) {
            paths.target = this.config.entityTypeCodes.target || null;
            paths.source = this.config.entityTypeCodes.source || null;
        } else if (this.config.entityTypeCodesFile) {
            paths.target = this.config.entityTypeCodesFile;
        }

        return {
            target: this.loadCsvMap(paths.target, 'target'),
            source: this.loadCsvMap(paths.source, 'source')
        };
    }

    loadCsvMap(csvPath, label) {
        if (!csvPath) {
            return new Map();
        }
        if (!fs.existsSync(csvPath)) {
            console.warn(`⚠  ${label} CSV not found at ${csvPath} — skipping.`);
            return new Map();
        }

        try {
            const csvContent = fs.readFileSync(csvPath, 'utf8');
            const lines = csvContent.split('\n').filter(line => line.trim());
            const typeCodes = new Map();

            // Skip header row if present
            const startIndex = lines[0].toLowerCase().includes('entity') || lines[0].toLowerCase().includes('name') ? 1 : 0;

            for (let i = startIndex; i < lines.length; i++) {
                const line = lines[i].trim();
                if (line) {
                    const [entityName, typeCode] = line.split(',').map(s => s.trim().replace(/"/g, ''));
                    if (entityName && typeCode) {
                        typeCodes.set(entityName, typeCode);
                    }
                }
            }

            console.log(`✓ Loaded ${typeCodes.size} ${label} entity type codes from ${path.basename(csvPath)}`);
            return typeCodes;
        } catch (error) {
            console.warn(`⚠  Could not load ${label} CSV (${csvPath}): ${error.message}`);
            return new Map();
        }
    }

    createBackup(filePath) {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const backupPath = `${filePath}.backup.${timestamp}`;
        fs.copyFileSync(filePath, backupPath);
        console.log(`✓ Backup created: ${path.basename(backupPath)}`);
        return backupPath;
    }

    // ---------- XML transformations (operate on a parsed DOM) ----------

    removeElements(doc, tagNames) {
        tagNames.forEach(tagName => {
            const elements = Array.from(doc.getElementsByTagName(tagName));
            elements.forEach(element => {
                if (element.parentNode) {
                    element.parentNode.removeChild(element);
                }
            });

            if (elements.length > 0) {
                this.summary.elementsRemoved[tagName] =
                    (this.summary.elementsRemoved[tagName] || 0) + elements.length;
            }
        });
    }

    removeAttributes(doc, tagName, attributes) {
        const elements = doc.getElementsByTagName(tagName);
        const perAttrCounts = {};

        for (let i = 0; i < elements.length; i++) {
            const element = elements[i];
            attributes.forEach(attr => {
                if (element.hasAttribute(attr)) {
                    element.removeAttribute(attr);
                    perAttrCounts[attr] = (perAttrCounts[attr] || 0) + 1;
                }
            });
        }

        Object.entries(perAttrCounts).forEach(([attr, count]) => {
            const key = `<${tagName}>.${attr}`;
            this.summary.attributesRemoved[key] = (this.summary.attributesRemoved[key] || 0) + count;
        });
    }

    // Two-tier lookup:
    //   1) target CSV — the code the v8 org actually uses (authoritative when both orgs know the entity)
    //   2) source CSV — seed value for entities that don't exist in v8 yet (v8 will reassign on import)
    //   3) '##' placeholder — neither org knows it, human must fix
    resolveObjectTypeCode(entityName) {
        const targetCode = this.targetCodes.get(entityName);
        if (targetCode) return { code: targetCode, origin: 'target' };

        const sourceCode = this.sourceCodes.get(entityName);
        if (sourceCode) return { code: sourceCode, origin: 'source' };

        return { code: '##', origin: 'placeholder' };
    }

    addObjectTypeCode(doc) {
        const entities = doc.getElementsByTagName('Entity');
        const otc = this.summary.objectTypeCodes;

        for (let i = 0; i < entities.length; i++) {
            const entity = entities[i];

            // Skip if ObjectTypeCode already present
            if (entity.getElementsByTagName('ObjectTypeCode').length > 0) continue;

            const nameElements = entity.getElementsByTagName('Name');
            if (nameElements.length === 0) continue;

            const entityName = nameElements[0].textContent || nameElements[0].innerText;
            const { code, origin } = this.resolveObjectTypeCode(entityName);

            if (origin === 'target') {
                otc.fromTarget++;
            } else if (origin === 'source') {
                otc.fromSource++;
                otc.newEntities.push({ name: entityName, sourceCode: code });
            } else {
                otc.placeholders++;
                otc.missingEntities.push(entityName);
            }

            const objectTypeCode = doc.createElement('ObjectTypeCode');
            objectTypeCode.textContent = code;

            if (entity.firstChild) {
                entity.insertBefore(objectTypeCode, entity.firstChild);
            } else {
                entity.appendChild(objectTypeCode);
            }
            otc.added++;
        }
    }

    // ---------- solution.xml version rewrite ----------

    // Rewrites the <ImportExportXml> root attributes from any 9.x to the configured 8.x target.
    // Matches:
    //   version="9..."                    -> version="<targetVersion>"
    //   SolutionPackageVersion="9..."     -> SolutionPackageVersion="<targetPackageVersion>"
    // Leaves <Version>...</Version> (the solution's own version) untouched.
    rewriteSolutionVersion(xmlContent) {
        const cfg = this.config.solutionVersion || {};
        const targetVersion = cfg.targetVersion || '8.2.0009.0019';
        const targetPackage = cfg.targetPackageVersion || '8.2';
        const changes = [];

        // Note: the attribute name `version` is lowercase, so \b differentiates it
        // from the camelCase `SolutionPackageVersion` that also ends in "Version".
        const rules = [
            { attr: 'version',                target: targetVersion, re: /(\bversion=")9(?:\.\d+)*(")/g },
            { attr: 'SolutionPackageVersion', target: targetPackage, re: /(\bSolutionPackageVersion=")9(?:\.\d+)*(")/g }
        ];

        let updated = xmlContent;
        rules.forEach(rule => {
            updated = updated.replace(rule.re, (full, open, close) => {
                const from = full.slice(open.length, -close.length);
                changes.push({ attr: rule.attr, from, to: rule.target });
                return `${open}${rule.target}${close}`;
            });
        });

        this.summary.solutionXml.processed = true;
        this.summary.solutionXml.changes = changes;
        return updated;
    }

    // ---------- Core: process an XML string in memory ----------

    processXmlString(xmlContent) {
        const doc = this.parser.parseFromString(xmlContent, 'text/xml');

        const parseErrors = doc.getElementsByTagName('parsererror');
        if (parseErrors.length > 0) {
            throw new Error('XML parsing failed');
        }

        if (this.config.removeElements && this.config.removeElements.length > 0) {
            this.removeElements(doc, this.config.removeElements);
        }

        if (this.config.removeAttributes) {
            this.config.removeAttributes.forEach(rule => {
                this.removeAttributes(doc, rule.tagName, rule.attributes);
            });
        }

        if (this.config.addObjectTypeCode) {
            this.addObjectTypeCode(doc);
        }

        return this.serializer.serializeToString(doc);
    }

    // ---------- Entry point: dispatch by extension ----------

    processFile(filePath) {
        const ext = path.extname(filePath).toLowerCase();
        this.summary.inputFile = filePath;

        try {
            if (ext === '.zip') {
                this.processZip(filePath);
            } else {
                this.processXml(filePath);
            }
            this.printSummary();
        } catch (error) {
            console.error(`Error processing file: ${error.message}`);
            process.exit(1);
        }
    }

    processXml(filePath) {
        this.summary.mode = 'xml';
        console.log(`Processing XML: ${filePath}`);

        this.summary.backupFile = this.createBackup(filePath);
        this.summary.originalSize = fs.statSync(filePath).size;
        this.summary.originalXmlSize = this.summary.originalSize;

        const xmlContent = fs.readFileSync(filePath, 'utf8');
        const updatedXml = this.processXmlString(xmlContent);
        fs.writeFileSync(filePath, updatedXml, 'utf8');

        this.summary.finalSize = fs.statSync(filePath).size;
        this.summary.finalXmlSize = this.summary.finalSize;
    }

    processZip(zipPath) {
        this.summary.mode = 'zip';
        console.log(`Processing ZIP: ${zipPath}`);

        this.summary.backupFile = this.createBackup(zipPath);
        this.summary.originalSize = fs.statSync(zipPath).size;

        const zip = new AdmZip(zipPath);
        const entries = zip.getEntries();

        // Find customizations.xml (case-insensitive, may be at root or nested)
        const targetEntry = entries.find(e =>
            !e.isDirectory &&
            path.basename(e.entryName).toLowerCase() === TARGET_XML_ENTRY
        );

        if (!targetEntry) {
            throw new Error(
                `'${TARGET_XML_ENTRY}' not found inside ${path.basename(zipPath)}. ` +
                `Entries: ${entries.map(e => e.entryName).join(', ')}`
            );
        }

        this.summary.xmlEntryName = targetEntry.entryName;
        const originalXmlBuffer = targetEntry.getData();
        this.summary.originalXmlSize = originalXmlBuffer.length;

        const xmlContent = originalXmlBuffer.toString('utf8');
        const updatedXml = this.processXmlString(xmlContent);
        const updatedBuffer = Buffer.from(updatedXml, 'utf8');
        this.summary.finalXmlSize = updatedBuffer.length;

        // Most efficient path: only swap this single entry's data, preserve everything else
        zip.updateFile(targetEntry.entryName, updatedBuffer);

        // Also rewrite solution.xml version attributes (v9 -> v8) if present and configured
        if (this.config.solutionVersion) {
            const solutionEntry = entries.find(e =>
                !e.isDirectory &&
                path.basename(e.entryName).toLowerCase() === SOLUTION_XML_ENTRY
            );
            if (!solutionEntry) {
                this.summary.solutionXml.notFound = true;
            } else {
                const originalSolutionXml = solutionEntry.getData().toString('utf8');
                const updatedSolutionXml = this.rewriteSolutionVersion(originalSolutionXml);
                if (updatedSolutionXml !== originalSolutionXml) {
                    zip.updateFile(solutionEntry.entryName, Buffer.from(updatedSolutionXml, 'utf8'));
                }
            }
        }

        zip.writeZip(zipPath);

        this.summary.finalSize = fs.statSync(zipPath).size;
    }

    // ---------- Summary report ----------

    printSummary() {
        const s = this.summary;
        const elapsed = ((Date.now() - s.startTime) / 1000).toFixed(2);
        const line = '═'.repeat(63);
        const sub = '─'.repeat(40);

        const fmtBytes = (n) => {
            if (n < 1024) return `${n} B`;
            if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
            return `${(n / 1024 / 1024).toFixed(2)} MB`;
        };
        const pctDelta = (before, after) => {
            if (!before) return '';
            const pct = ((after - before) / before * 100).toFixed(1);
            const sign = pct >= 0 ? '+' : '';
            return ` (${sign}${pct}%)`;
        };

        console.log('');
        console.log(line);
        console.log('  CRM Solution Downgrade Summary  (v9 → v8)');
        console.log(line);
        console.log(`  Mode          : ${s.mode.toUpperCase()}`);
        console.log(`  Input         : ${path.basename(s.inputFile)}`);
        console.log(`  Backup        : ${path.basename(s.backupFile)}`);
        if (s.mode === 'zip') {
            console.log(`  XML entry     : ${s.xmlEntryName}`);
            console.log(`  Zip size      : ${fmtBytes(s.originalSize)} → ${fmtBytes(s.finalSize)}${pctDelta(s.originalSize, s.finalSize)}`);
            console.log(`  XML size      : ${fmtBytes(s.originalXmlSize)} → ${fmtBytes(s.finalXmlSize)}${pctDelta(s.originalXmlSize, s.finalXmlSize)}`);
        } else {
            console.log(`  XML size      : ${fmtBytes(s.originalSize)} → ${fmtBytes(s.finalSize)}${pctDelta(s.originalSize, s.finalSize)}`);
        }
        console.log(`  Duration      : ${elapsed}s`);

        // Elements
        const elementEntries = Object.entries(s.elementsRemoved).sort((a, b) => b[1] - a[1]);
        const totalElements = elementEntries.reduce((acc, [, n]) => acc + n, 0);
        console.log('');
        console.log('  Elements Removed');
        console.log(`  ${sub}`);
        if (elementEntries.length === 0) {
            console.log('    (none)');
        } else {
            elementEntries.forEach(([tag, count]) => {
                console.log(`    ${tag.padEnd(34)} ${String(count).padStart(5)}`);
            });
            console.log(`    ${'─'.repeat(34)} ${'─'.repeat(5)}`);
            console.log(`    ${'TOTAL'.padEnd(34)} ${String(totalElements).padStart(5)}`);
        }

        // Attributes
        const attrEntries = Object.entries(s.attributesRemoved).sort((a, b) => b[1] - a[1]);
        const totalAttrs = attrEntries.reduce((acc, [, n]) => acc + n, 0);
        console.log('');
        console.log('  Attributes Removed');
        console.log(`  ${sub}`);
        if (attrEntries.length === 0) {
            console.log('    (none)');
        } else {
            attrEntries.forEach(([key, count]) => {
                console.log(`    ${key.padEnd(34)} ${String(count).padStart(5)}`);
            });
            console.log(`    ${'─'.repeat(34)} ${'─'.repeat(5)}`);
            console.log(`    ${'TOTAL'.padEnd(34)} ${String(totalAttrs).padStart(5)}`);
        }

        // ObjectTypeCode
        const otc = s.objectTypeCodes;
        console.log('');
        console.log('  ObjectTypeCode Elements Added');
        console.log(`  ${sub}`);
        console.log(`    ${'Mapped from target CSV'.padEnd(34)} ${String(otc.fromTarget).padStart(5)}`);
        console.log(`    ${'Seeded from source CSV (new)'.padEnd(34)} ${String(otc.fromSource).padStart(5)}`);
        console.log(`    ${"Used '##' placeholder".padEnd(34)} ${String(otc.placeholders).padStart(5)}`);
        console.log(`    ${'─'.repeat(34)} ${'─'.repeat(5)}`);
        console.log(`    ${'TOTAL'.padEnd(34)} ${String(otc.added).padStart(5)}`);

        if (otc.newEntities.length > 0) {
            console.log('');
            console.log(`  ℹ  New entities not yet in target v8 (seeded with source code):`);
            otc.newEntities.forEach(({ name, sourceCode }) =>
                console.log(`     - ${name.padEnd(40)} ${sourceCode}`));
            console.log(`  → v8 will assign a fresh ObjectTypeCode on import; the seed is only a hint.`);
        }

        if (otc.missingEntities.length > 0) {
            console.log('');
            console.log(`  ⚠  Entities missing from BOTH CSVs (using '##'):`);
            otc.missingEntities.forEach(name => console.log(`     - ${name}`));
            console.log(`  → Replace '##' in the XML with real type codes before importing.`);
        }

        // solution.xml version rewrite
        const sol = s.solutionXml;
        if (s.mode === 'zip' && this.config.solutionVersion) {
            console.log('');
            console.log('  solution.xml Version Rewrite');
            console.log(`  ${sub}`);
            if (sol.notFound) {
                console.log(`    ⚠  solution.xml not found in the zip — skipped.`);
            } else if (sol.changes.length === 0) {
                console.log(`    (no 9.x version attributes found — already downgraded or v8 export)`);
            } else {
                sol.changes.forEach(c => {
                    console.log(`    ${c.attr.padEnd(24)} ${c.from.padEnd(12)} →  ${c.to}`);
                });
            }
        }

        console.log('');
        console.log(line);
        console.log('  ✓ Done');
        console.log(line);
        console.log('');
    }
}

// ---------- CLI ----------

function main() {
    const args = process.argv.slice(2);

    if (args.length !== 2) {
        console.log('Usage: node crm-xml-processor.js <config.json> <solution.zip|customizations.xml>');
        console.log('Examples:');
        console.log('  node crm-xml-processor.js config.json data/MySolution_1_0_0_0.zip');
        console.log('  node crm-xml-processor.js config.json data/customizations.xml');
        process.exit(1);
    }

    const [configPath, inputPath] = args;

    if (!fs.existsSync(configPath)) {
        console.error(`Config file not found: ${configPath}`);
        process.exit(1);
    }
    if (!fs.existsSync(inputPath)) {
        console.error(`Input file not found: ${inputPath}`);
        process.exit(1);
    }

    const processor = new CRMXMLProcessor(configPath);
    processor.processFile(inputPath);
}

if (require.main === module) {
    main();
}

module.exports = CRMXMLProcessor;
