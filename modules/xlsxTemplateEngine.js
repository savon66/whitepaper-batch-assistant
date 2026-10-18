/**
 * Excel 白皮模板识别与生成模块
 *
 * 依赖：
 *   1. lib/jszip.min.js
 *   2. WPS Application.FileSystem
 *
 * 设计原则：
 *   - 不使用 SheetJS 重写整个工作簿；
 *   - 只修改 xlsx 压缩包中的 xl/sharedStrings.xml；
 *   - 保留合并单元格、行高、列宽、打印设置和其他版式。
 */
var XlsxTemplateEngine = (function () {
    "use strict";

    var MAIN_NS =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    var REL_NS =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    var PACKAGE_REL_NS =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    var XML_NS =
        "http://www.w3.org/XML/1998/namespace";

    function logInfo(message) {
        if (
            typeof Logger !== "undefined" &&
            Logger &&
            typeof Logger.info === "function"
        ) {
            Logger.info(message);
        }
    }

    function logWarn(message) {
        if (
            typeof Logger !== "undefined" &&
            Logger &&
            typeof Logger.warn === "function"
        ) {
            Logger.warn(message);
        }
    }

    function getFileSystem() {
        if (
            typeof Application === "undefined" ||
            !Application ||
            !Application.FileSystem
        ) {
            throw new Error(
                "当前 WPS 环境未提供 Application.FileSystem"
            );
        }

        return Application.FileSystem;
    }

    function ensureDependencies() {
        if (
            typeof WhitePaperFields ===
                "undefined" ||
            !WhitePaperFields
        ) {
            throw new Error(
                "字段注册表未加载，请检查 modules\\fieldRegistry.js"
            );
        }

        if (typeof JSZip === "undefined") {
            throw new Error(
                "JSZip 未加载，请检查 lib\\jszip.min.js"
            );
        }

        if (typeof DOMParser === "undefined") {
            throw new Error(
                "当前任务窗格环境不支持 DOMParser"
            );
        }

        if (typeof XMLSerializer === "undefined") {
            throw new Error(
                "当前任务窗格环境不支持 XMLSerializer"
            );
        }
    }

    function readBinaryFile(filePath) {
        var fs = getFileSystem();

        if (
            typeof fs.readAsBinaryString !== "function"
        ) {
            throw new Error(
                "当前 WPS FileSystem 不支持 readAsBinaryString"
            );
        }

        var binary = fs.readAsBinaryString(
            String(filePath || "")
        );

        if (typeof binary !== "string") {
            throw new Error(
                "读取 Excel 模板失败：未获得二进制字符串"
            );
        }

        return binary;
    }

    function writeBinaryFile(filePath, binary) {
        var fs = getFileSystem();

        if (
            typeof fs.writeAsBinaryString !== "function"
        ) {
            throw new Error(
                "当前 WPS FileSystem 不支持 writeAsBinaryString"
            );
        }

        fs.writeAsBinaryString(
            String(filePath || ""),
            binary
        );
    }

    function parseXml(xmlText, fileName) {
        var parser = new DOMParser();
        var doc = parser.parseFromString(
            String(xmlText || ""),
            "application/xml"
        );

        var errors =
            doc.getElementsByTagName("parsererror");

        if (errors && errors.length > 0) {
            throw new Error(
                "解析 " +
                fileName +
                " 失败：" +
                errors[0].textContent
            );
        }

        return doc;
    }

    function serializeXml(doc) {
        var serializer = new XMLSerializer();
        var xml = serializer.serializeToString(doc);

        if (xml.indexOf("<?xml") !== 0) {
            xml =
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                xml;
        }

        return xml;
    }

    function getElementsByLocalName(node, localName) {
        if (
            node &&
            typeof node.getElementsByTagNameNS ===
                "function"
        ) {
            return node.getElementsByTagNameNS(
                "*",
                localName
            );
        }

        return node.getElementsByTagName(localName);
    }

    function getDirectChildrenByLocalName(
        node,
        localName
    ) {
        var result = [];
        var children = node.childNodes || [];

        for (var i = 0; i < children.length; i++) {
            var child = children[i];

            if (
                child.nodeType === 1 &&
                (
                    child.localName === localName ||
                    child.nodeName === localName ||
                    child.nodeName ===
                        "x:" + localName
                )
            ) {
                result.push(child);
            }
        }

        return result;
    }

    function getFirstDirectChild(
        node,
        localName
    ) {
        var children =
            getDirectChildrenByLocalName(
                node,
                localName
            );

        return children.length > 0
            ? children[0]
            : null;
    }

    function getNodeText(node) {
        if (!node) {
            return "";
        }

        return node.textContent || "";
    }

    function normalizePath(path) {
        var value = String(path || "")
            .replace(/\\/g, "/");

        while (value.indexOf("./") === 0) {
            value = value.substring(2);
        }

        var parts = value.split("/");
        var result = [];

        for (var i = 0; i < parts.length; i++) {
            var part = parts[i];

            if (!part || part === ".") {
                continue;
            }

            if (part === "..") {
                result.pop();
            } else {
                result.push(part);
            }
        }

        return result.join("/");
    }

    function resolveWorkbookTarget(target) {
        var value = String(target || "");

        if (value.charAt(0) === "/") {
            value = value.substring(1);
        }

        if (value.indexOf("xl/") === 0) {
            return normalizePath(value);
        }

        return normalizePath("xl/" + value);
    }

    function getRelationshipId(sheetNode) {
        var relationshipId = "";

        try {
            relationshipId =
                sheetNode.getAttributeNS(
                    REL_NS,
                    "id"
                ) || "";
        } catch (e) {}

        if (!relationshipId) {
            relationshipId =
                sheetNode.getAttribute("r:id") ||
                sheetNode.getAttribute("id") ||
                "";
        }

        return relationshipId;
    }

    function buildRelationshipMap(relsDoc) {
        var result = {};
        var relationships =
            getElementsByLocalName(
                relsDoc,
                "Relationship"
            );

        for (
            var i = 0;
            i < relationships.length;
            i++
        ) {
            var relationship = relationships[i];
            var id =
                relationship.getAttribute("Id") || "";
            var target =
                relationship.getAttribute("Target") || "";

            if (id && target) {
                result[id] =
                    resolveWorkbookTarget(target);
            }
        }

        return result;
    }

    function extractSheetDefinitions(
        workbookDoc,
        relationshipMap
    ) {
        var result = [];
        var sheetNodes =
            getElementsByLocalName(
                workbookDoc,
                "sheet"
            );

        for (
            var i = 0;
            i < sheetNodes.length;
            i++
        ) {
            var sheetNode = sheetNodes[i];
            var relationshipId =
                getRelationshipId(sheetNode);

            if (!relationshipId) {
                continue;
            }

            var sheetPath =
                relationshipMap[relationshipId];

            if (!sheetPath) {
                continue;
            }

            result.push({
                name:
                    sheetNode.getAttribute("name") ||
                    "",
                relationshipId: relationshipId,
                path: sheetPath
            });
        }

        return result;
    }

    function hasUnderline(rPr) {
        if (!rPr) {
            return false;
        }

        return (
            getElementsByLocalName(
                rPr,
                "u"
            ).length > 0
        );
    }

    function getFirstAttributeValue(
        node,
        localName,
        attributeName
    ) {
        if (!node) {
            return "";
        }

        var nodes =
            getElementsByLocalName(
                node,
                localName
            );

        if (nodes.length === 0) {
            return "";
        }

        return String(
            nodes[0].getAttribute(
                attributeName
            ) || ""
        );
    }

    function extractRunFormat(rPr) {
        if (!rPr) {
            return {};
        }

        var sizeText =
            getFirstAttributeValue(
                rPr,
                "sz",
                "val"
            );

        var size = Number(sizeText);

        return {
            name: getFirstAttributeValue(
                rPr,
                "rFont",
                "val"
            ),
            nameFarEast: getFirstAttributeValue(
                rPr,
                "rFont",
                "val"
            ),
            size:
                isFinite(size) && size > 0
                    ? size
                    : null
        };
    }

    function extractCellFontFormat(fontNode) {
        if (!fontNode) {
            return {};
        }

        var sizeText =
            getFirstAttributeValue(
                fontNode,
                "sz",
                "val"
            );

        var size = Number(sizeText);

        return {
            name: getFirstAttributeValue(
                fontNode,
                "name",
                "val"
            ),
            nameFarEast: getFirstAttributeValue(
                fontNode,
                "name",
                "val"
            ),
            size:
                isFinite(size) && size > 0
                    ? size
                    : null
        };
    }

    function buildCellStyleFormats(stylesDoc) {
        var result = [];

        if (!stylesDoc) {
            return result;
        }

        var fontsNode =
            getElementsByLocalName(
                stylesDoc,
                "fonts"
            )[0];

        var cellXfsNode =
            getElementsByLocalName(
                stylesDoc,
                "cellXfs"
            )[0];

        if (!fontsNode || !cellXfsNode) {
            return result;
        }

        var fontNodes =
            getDirectChildrenByLocalName(
                fontsNode,
                "font"
            );

        var xfNodes =
            getDirectChildrenByLocalName(
                cellXfsNode,
                "xf"
            );

        for (var i = 0; i < xfNodes.length; i++) {
            var fontId = parseInt(
                xfNodes[i].getAttribute(
                    "fontId"
                ) || "0",
                10
            );

            result[i] =
                extractCellFontFormat(
                    fontNodes[fontId] || null
                );
        }

        return result;
    }

    function extractSharedStringInfo(siNode, index) {
        var runs =
            getDirectChildrenByLocalName(
                siNode,
                "r"
            );

        var plainText = "";
        var prefixText = "";
        var underlineText = "";
        var prefixRuns = [];
        var underlineRPr = null;
        var underlineStarted = false;

        if (runs.length === 0) {
            var directText =
                getFirstDirectChild(
                    siNode,
                    "t"
                );

            plainText = getNodeText(directText);

            return {
                index: index,
                node: siNode,
                plainText: plainText,
                prefixText: plainText,
                underlineText: "",
                underlineUnits: 0,
                prefixRuns: [],
                underlineRPr: null,
                valueFormat: {}
            };
        }

        for (var i = 0; i < runs.length; i++) {
            var run = runs[i];
            var rPr =
                getFirstDirectChild(
                    run,
                    "rPr"
                );
            var textNode =
                getFirstDirectChild(
                    run,
                    "t"
                );
            var text = getNodeText(textNode);
            var underlined = hasUnderline(rPr);

            plainText += text;

            if (underlined) {
                underlineStarted = true;
                underlineText += text;

                if (!underlineRPr && rPr) {
                    underlineRPr =
                        rPr.cloneNode(true);
                }
            } else if (!underlineStarted) {
                prefixText += text;
                prefixRuns.push(
                    run.cloneNode(true)
                );
            } else {
                /*
                 * 已进入下划线区域后出现的非下划线运行，
                 * 也计入可替换区域，防止旧样例存在异常分段。
                 */
                underlineText += text;
            }
        }

        return {
            index: index,
            node: siNode,
            plainText: plainText,
            prefixText: prefixText,
            underlineText: underlineText,
            underlineUnits:
                getVisualUnits(underlineText),
            prefixRuns: prefixRuns,
            underlineRPr: underlineRPr,
            valueFormat:
                extractRunFormat(
                    underlineRPr
                )
        };
    }

    function buildSharedStringInfos(sharedStringsDoc) {
        var result = [];
        var siNodes =
            getElementsByLocalName(
                sharedStringsDoc,
                "si"
            );

        for (var i = 0; i < siNodes.length; i++) {
            result.push(
                extractSharedStringInfo(
                    siNodes[i],
                    i
                )
            );
        }

        return result;
    }

    function detectFieldName(text) {
        var match =
            WhitePaperFields
                .matchTemplateText(text);

        return match
            ? match.fieldName
            : "";
    }

    function getCellValueIndex(cellNode) {
        var valueNode =
            getFirstDirectChild(
                cellNode,
                "v"
            );

        if (!valueNode) {
            return -1;
        }

        var index = parseInt(
            getNodeText(valueNode),
            10
        );

        return isNaN(index) ? -1 : index;
    }

    function detectFieldsInSheet(
        sheetDoc,
        sheetDefinition,
        sharedStringInfos,
        cellStyleFormats
    ) {
        var fields = {};
        var cells =
            getElementsByLocalName(
                sheetDoc,
                "c"
            );

        for (var i = 0; i < cells.length; i++) {
            var cell = cells[i];

            if (
                String(
                    cell.getAttribute("t") || ""
                ) !== "s"
            ) {
                continue;
            }

            var sharedStringIndex =
                getCellValueIndex(cell);

            if (
                sharedStringIndex < 0 ||
                !sharedStringInfos[
                    sharedStringIndex
                ]
            ) {
                continue;
            }

            var sharedInfo =
                sharedStringInfos[
                    sharedStringIndex
                ];

            var fieldName =
                detectFieldName(
                    sharedInfo.plainText
                );

            if (!fieldName || fields[fieldName]) {
                continue;
            }

            fields[fieldName] = {
                name: fieldName,
                sheetName:
                    sheetDefinition.name,
                sheetPath:
                    sheetDefinition.path,
                cellRef:
                    cell.getAttribute("r") || "",
                sharedStringIndex:
                    sharedStringIndex,
                prefixText:
                    sharedInfo.prefixText,
                underlineText:
                    sharedInfo.underlineText,
                underlineUnits:
                    sharedInfo.underlineUnits,
                format:
                    (
                        sharedInfo.valueFormat &&
                        (
                            sharedInfo.valueFormat.name ||
                            sharedInfo.valueFormat.nameFarEast ||
                            sharedInfo.valueFormat.size
                        )
                    )
                        ? sharedInfo.valueFormat
                        : (
                            cellStyleFormats[
                                parseInt(
                                    cell.getAttribute("s") || "0",
                                    10
                                )
                            ] || {}
                        )
            };
        }

        return fields;
    }

    function mergeFieldMaps(target, source) {
        for (var key in source) {
            if (
                source.hasOwnProperty(key) &&
                !target[key]
            ) {
                target[key] = source[key];
            }
        }
    }

    function getVisualUnits(text) {
        var value = String(
            text === undefined || text === null
                ? ""
                : text
        );

        var units = 0;

        for (var i = 0; i < value.length; i++) {
            var code = value.charCodeAt(i);

            /* UTF-16 代理对按一个全宽字符处理。 */
            if (
                code >= 0xd800 &&
                code <= 0xdbff &&
                i + 1 < value.length
            ) {
                var next = value.charCodeAt(i + 1);

                if (
                    next >= 0xdc00 &&
                    next <= 0xdfff
                ) {
                    units += 2;
                    i++;
                    continue;
                }
            }

            var isFullWidth =
                code >= 0x2e80 ||
                (
                    code >= 0xff01 &&
                    code <= 0xff60
                ) ||
                (
                    code >= 0xffe0 &&
                    code <= 0xffe6
                );

            units += isFullWidth ? 2 : 1;
        }

        return units;
    }

    function repeatSpaces(count) {
        var length = Math.max(
            0,
            Math.round(Number(count) || 0)
        );

        var result = "";

        for (var i = 0; i < length; i++) {
            result += " ";
        }

        return result;
    }

    function createElement(doc, localName) {
        return doc.createElementNS(
            MAIN_NS,
            localName
        );
    }

    function setRunFont(
        rPr,
        doc,
        fontName
    ) {
        var fontNodes =
            getElementsByLocalName(
                rPr,
                "rFont"
            );

        var fontNode =
            fontNodes.length > 0
                ? fontNodes[0]
                : null;

        if (!fontNode) {
            fontNode = createElement(
                doc,
                "rFont"
            );
            rPr.appendChild(fontNode);
        }

        fontNode.setAttribute(
            "val",
            fontName
        );
    }

    function setRunSize(rPr, doc, size) {
        var sizeNodes =
            getElementsByLocalName(
                rPr,
                "sz"
            );

        var sizeNode =
            sizeNodes.length > 0
                ? sizeNodes[0]
                : null;

        if (!sizeNode) {
            sizeNode = createElement(
                doc,
                "sz"
            );
            rPr.appendChild(sizeNode);
        }

        sizeNode.setAttribute(
            "val",
            String(size)
        );
    }

    function ensureUnderline(rPr, doc) {
        var underlineNodes =
            getElementsByLocalName(
                rPr,
                "u"
            );

        if (underlineNodes.length === 0) {
            rPr.appendChild(
                createElement(doc, "u")
            );
        }
    }

    function createDefaultUnderlineRPr(doc) {
        var rPr = createElement(doc, "rPr");

        ensureUnderline(rPr, doc);
        setRunSize(rPr, doc, 16);
        setRunFont(rPr, doc, "仿宋");

        var color = createElement(doc, "color");
        color.setAttribute("rgb", "FF000000");
        rPr.appendChild(color);

        var charset = createElement(
            doc,
            "charset"
        );
        charset.setAttribute("val", "134");
        rPr.appendChild(charset);

        return rPr;
    }

    function createTextRun(
        doc,
        text,
        rPrTemplate,
        isValue
    ) {
        var run = createElement(doc, "r");
        var rPr = rPrTemplate
            ? rPrTemplate.cloneNode(true)
            : createDefaultUnderlineRPr(doc);

        ensureUnderline(rPr, doc);
        setRunSize(rPr, doc, 16);

        if (isValue) {
            setRunFont(rPr, doc, "黑体");
        }

        run.appendChild(rPr);

        var t = createElement(doc, "t");

        if (
            text.length > 0 &&
            (
                /^\s/.test(text) ||
                /\s$/.test(text) ||
                text.indexOf("  ") !== -1
            )
        ) {
            t.setAttributeNS(
                XML_NS,
                "xml:space",
                "preserve"
            );
        }

        t.appendChild(
            doc.createTextNode(text)
        );

        run.appendChild(t);

        return run;
    }

    function findSharedStringInfo(
        sharedStringInfos,
        index
    ) {
        if (
            index < 0 ||
            index >= sharedStringInfos.length
        ) {
            return null;
        }

        return sharedStringInfos[index];
    }

    function calculateTargetCenter(fields) {
        var source = fields || {};
        var ordered =
            WhitePaperFields
                .getFillableFieldNames();
        var preferred = ["种类"]
            .concat(ordered);
        var baseField = null;

        for (
            var i = 0;
            i < preferred.length;
            i++
        ) {
            var candidate =
                source[preferred[i]];

            if (
                candidate &&
                candidate.underlineUnits > 0
            ) {
                baseField = candidate;
                break;
            }
        }

        if (!baseField) {
            return null;
        }

        return (
            getVisualUnits(baseField.prefixText) +
            baseField.underlineUnits / 2
        );
    }

    function calculateSpacing(
        fieldInfo,
        value,
        targetCenter
    ) {
        var prefixUnits =
            getVisualUnits(fieldInfo.prefixText);
        var underlineUnits =
            Math.max(
                0,
                fieldInfo.underlineUnits
            );
        var valueUnits = getVisualUnits(value);

        var availableUnits =
            Math.max(
                0,
                underlineUnits - valueUnits
            );

        var leftUnits;

        if (
            targetCenter === null ||
            targetCenter === undefined
        ) {
            leftUnits =
                Math.floor(availableUnits / 2);
        } else {
            leftUnits = Math.round(
                targetCenter -
                prefixUnits -
                valueUnits / 2
            );
        }

        var minimumLeft =
            fieldInfo.name ===
                "配套教材版本" &&
            availableUnits > 0
                ? 1
                : 0;

        leftUnits = Math.max(
            minimumLeft,
            Math.min(
                availableUnits,
                leftUnits
            )
        );

        /*
         * “必修 第一册”“选择性必修 第三册”等册次值含一个
         * 半角空格，总视觉宽度会变成奇数。Math.round 会把多出的
         * 一个空白单位放到左侧，导致文字比标准位置偏右一格。
         *
         * 按白皮排版要求，将这一格从左侧移到右侧：文字整体前移
         * 一个半角字符，横线总长度仍保持不变。
         */
        var shouldShiftVolumeLeft =
            fieldInfo.name === "册次" &&
            /[^\s]\s+[^\s]/.test(
                String(value || "")
            ) &&
            availableUnits > 0 &&
            leftUnits > minimumLeft;

        if (shouldShiftVolumeLeft) {
            leftUnits -= 1;
        }

        var rightUnits = Math.max(
            0,
            availableUnits - leftUnits
        );

        return {
            prefixUnits: prefixUnits,
            underlineUnits: underlineUnits,
            valueUnits: valueUnits,
            leftUnits: leftUnits,
            rightUnits: rightUnits,
            shiftedForInternalSpace:
                shouldShiftVolumeLeft,
            overflow: valueUnits > underlineUnits
        };
    }

    function replaceSharedString(
        sharedStringsDoc,
        sharedStringInfos,
        fieldInfo,
        value,
        targetCenter
    ) {
        var sharedInfo =
            findSharedStringInfo(
                sharedStringInfos,
                fieldInfo.sharedStringIndex
            );

        if (!sharedInfo || !sharedInfo.node) {
            throw new Error(
                "无法定位字段“" +
                fieldInfo.name +
                "”的共享字符串"
            );
        }

        var spacing = calculateSpacing(
            fieldInfo,
            value,
            targetCenter
        );

        if (spacing.overflow) {
            logWarn(
                "字段“" +
                fieldInfo.name +
                "”内容超过模板横线宽度：" +
                value
            );
        }

        var siNode = sharedInfo.node;

        while (siNode.firstChild) {
            siNode.removeChild(
                siNode.firstChild
            );
        }

        var prefixRuns =
            sharedInfo.prefixRuns || [];

        if (prefixRuns.length > 0) {
            for (
                var i = 0;
                i < prefixRuns.length;
                i++
            ) {
                siNode.appendChild(
                    sharedStringsDoc.importNode
                        ? sharedStringsDoc.importNode(
                            prefixRuns[i],
                            true
                        )
                        : prefixRuns[i].cloneNode(true)
                );
            }
        } else {
            /* 极端情况下保底重建标签运行。 */
            var labelRun = createElement(
                sharedStringsDoc,
                "r"
            );
            var labelText = createElement(
                sharedStringsDoc,
                "t"
            );
            labelText.setAttributeNS(
                XML_NS,
                "xml:space",
                "preserve"
            );
            labelText.appendChild(
                sharedStringsDoc.createTextNode(
                    fieldInfo.prefixText || ""
                )
            );
            labelRun.appendChild(labelText);
            siNode.appendChild(labelRun);
        }

        var blankRPr =
            sharedInfo.underlineRPr ||
            createDefaultUnderlineRPr(
                sharedStringsDoc
            );

        if (spacing.leftUnits > 0) {
            siNode.appendChild(
                createTextRun(
                    sharedStringsDoc,
                    repeatSpaces(
                        spacing.leftUnits
                    ),
                    blankRPr,
                    false
                )
            );
        }

        if (value) {
            siNode.appendChild(
                createTextRun(
                    sharedStringsDoc,
                    value,
                    blankRPr,
                    true
                )
            );
        }

        if (spacing.rightUnits > 0) {
            siNode.appendChild(
                createTextRun(
                    sharedStringsDoc,
                    repeatSpaces(
                        spacing.rightUnits
                    ),
                    blankRPr,
                    false
                )
            );
        }

        return spacing;
    }

    function normalizeValue(value) {
        return String(
            value === undefined || value === null
                ? ""
                : value
        )
            .replace(/[\r\n\t]+/g, " ")
            .replace(/\s+/g, " ")
            .trim();
    }

    function splitGradeAndVolume(value) {
        return WhitePaperFields
            .splitGradeAndVolume(value);
    }

    function prepareValues(values) {
        var source = values || {};
        var prepared = {};

        for (var key in source) {
            if (!source.hasOwnProperty(key)) {
                continue;
            }

            var canonical =
                WhitePaperFields
                    .resolveHeader(key) ||
                WhitePaperFields
                    .resolve(key);

            if (!canonical) {
                continue;
            }

            prepared[canonical] =
                WhitePaperFields
                    .shouldClearInOutput(
                        canonical
                    )
                    ? ""
                    : WhitePaperFields
                        .formatValue(
                            canonical,
                            normalizeValue(
                                source[key]
                            )
                        );
        }

        var combined =
            prepared["使用年级及册次"] ||
            "";

        if (combined) {
            var split =
                splitGradeAndVolume(
                    combined
                );

            if (
                !prepared.hasOwnProperty(
                    "使用年级"
                )
            ) {
                prepared["使用年级"] =
                    split.grade;
            }

            if (
                !prepared.hasOwnProperty(
                    "册次"
                )
            ) {
                prepared["册次"] =
                    split.volume;
            }
        }

        return prepared;
    }

    function getMissingFields(fields) {
        /*
         * v1.0 不再把某一套固定字段当成所有模板的必填项。
         * 数据与模板之间的逐字段兼容检查由 taskpane 统一完成。
         */
        return [];
    }

    function loadPackage(templatePath) {
        ensureDependencies();

        var binary = readBinaryFile(templatePath);

        return JSZip.loadAsync(
            binary,
            {
                binary: true,
                checkCRC32: false
            }
        ).then(function (zip) {
            var workbookFile =
                zip.file("xl/workbook.xml");
            var relsFile =
                zip.file(
                    "xl/_rels/workbook.xml.rels"
                );
            var sharedStringsFile =
                zip.file(
                    "xl/sharedStrings.xml"
                );

            var stylesFile =
                zip.file(
                    "xl/styles.xml"
                );

            if (!workbookFile) {
                throw new Error(
                    "模板缺少 xl/workbook.xml"
                );
            }

            if (!relsFile) {
                throw new Error(
                    "模板缺少 workbook.xml.rels"
                );
            }

            if (!sharedStringsFile) {
                throw new Error(
                    "模板缺少 xl/sharedStrings.xml"
                );
            }

            return Promise.all([
                workbookFile.async("string"),
                relsFile.async("string"),
                sharedStringsFile.async("string"),
                stylesFile
                    ? stylesFile.async("string")
                    : Promise.resolve("")
            ]).then(function (xmlParts) {
                var workbookDoc = parseXml(
                    xmlParts[0],
                    "xl/workbook.xml"
                );
                var relsDoc = parseXml(
                    xmlParts[1],
                    "workbook.xml.rels"
                );
                var sharedStringsDoc = parseXml(
                    xmlParts[2],
                    "xl/sharedStrings.xml"
                );

                var stylesDoc =
                    xmlParts[3]
                        ? parseXml(
                            xmlParts[3],
                            "xl/styles.xml"
                        )
                        : null;

                var cellStyleFormats =
                    buildCellStyleFormats(
                        stylesDoc
                    );

                var relationshipMap =
                    buildRelationshipMap(relsDoc);
                var sheets =
                    extractSheetDefinitions(
                        workbookDoc,
                        relationshipMap
                    );
                var sharedStringInfos =
                    buildSharedStringInfos(
                        sharedStringsDoc
                    );

                if (sheets.length === 0) {
                    throw new Error(
                        "模板中没有可识别的工作表"
                    );
                }

                var sheetPromises = [];

                for (
                    var i = 0;
                    i < sheets.length;
                    i++
                ) {
                    (function (sheetDefinition) {
                        var sheetFile =
                            zip.file(
                                sheetDefinition.path
                            );

                        if (!sheetFile) {
                            return;
                        }

                        sheetPromises.push(
                            sheetFile
                                .async("string")
                                .then(function (sheetXml) {
                                    return {
                                        definition:
                                            sheetDefinition,
                                        document: parseXml(
                                            sheetXml,
                                            sheetDefinition.path
                                        )
                                    };
                                })
                        );
                    })(sheets[i]);
                }

                return Promise.all(
                    sheetPromises
                ).then(function (sheetParts) {
                    var fields = {};

                    for (
                        var j = 0;
                        j < sheetParts.length;
                        j++
                    ) {
                        mergeFieldMaps(
                            fields,
                            detectFieldsInSheet(
                                sheetParts[j].document,
                                sheetParts[j].definition,
                                sharedStringInfos,
                                cellStyleFormats
                            )
                        );
                    }

                    var missingFields =
                        getMissingFields(fields);

                    var primarySheetName = "";

                    for (var fieldName in fields) {
                        if (
                            fields.hasOwnProperty(
                                fieldName
                            )
                        ) {
                            primarySheetName =
                                fields[fieldName]
                                    .sheetName;
                            break;
                        }
                    }

                    return {
                        zip: zip,
                        workbookDoc: workbookDoc,
                        relsDoc: relsDoc,
                        sharedStringsDoc:
                            sharedStringsDoc,
                        sharedStringInfos:
                            sharedStringInfos,
                        sheets: sheets,
                        fields: fields,
                        missingFields: missingFields,
                        sheetName: primarySheetName
                    };
                });
            });
        });
    }

    /**
     * 识别 Excel 模板。
     *
     * callback(result, errorMessage)
     */
    function analyze(templatePath, callback) {
        loadPackage(templatePath)
            .then(function (packageInfo) {
                var publicFields = {};

                for (
                    var fieldName in packageInfo.fields
                ) {
                    if (
                        packageInfo.fields.hasOwnProperty(
                            fieldName
                        )
                    ) {
                        var field =
                            packageInfo.fields[fieldName];

                        publicFields[fieldName] = {
                            name: field.name,
                            sheetName: field.sheetName,
                            cellRef: field.cellRef,
                            sharedStringIndex:
                                field.sharedStringIndex,
                            prefixText: field.prefixText,
                            underlineUnits:
                                field.underlineUnits,
                            format: field.format || {}
                        };
                    }
                }

                var result = {
                    success:
                        Object.keys(publicFields)
                            .length > 0,
                    sheetName:
                        packageInfo.sheetName,
                    fields: publicFields,
                    missingFields:
                        packageInfo.missingFields,
                    fieldCount:
                        Object.keys(publicFields)
                            .length
                };

                logInfo(
                    "Excel 模板分析完成，识别到 " +
                    result.fieldCount +
                    " 个字段"
                );

                if (
                    typeof callback === "function"
                ) {
                    callback(result, "");
                }
            })
            .catch(function (error) {
                var message =
                    error && error.message
                        ? error.message
                        : String(error);

                if (
                    typeof callback === "function"
                ) {
                    callback(null, message);
                }
            });
    }

    /**
     * 根据 Excel 模板生成单个 xlsx 文件。
     *
     * values 支持字段注册表中的全部标准字段；仅替换模板实际存在的字段。
     *
     * callback(result, errorMessage)
     */
    function generate(
        templatePath,
        outputPath,
        values,
        callback
    ) {
        loadPackage(templatePath)
            .then(function (packageInfo) {
                if (
                    Object.keys(
                        packageInfo.fields || {}
                    ).length === 0
                ) {
                    throw new Error(
                        "模板中没有识别到可填写字段"
                    );
                }

                var preparedValues =
                    prepareValues(values);

                var targetCenter =
                    calculateTargetCenter(
                        packageInfo.fields
                    );

                var spacingResults = {};

                for (
                    var fieldName in preparedValues
                ) {
                    if (
                        !preparedValues.hasOwnProperty(
                            fieldName
                        ) ||
                        !packageInfo.fields[fieldName]
                    ) {
                        continue;
                    }

                    spacingResults[fieldName] =
                        replaceSharedString(
                            packageInfo.sharedStringsDoc,
                            packageInfo.sharedStringInfos,
                            packageInfo.fields[fieldName],
                            preparedValues[fieldName],
                            targetCenter
                        );
                }

                packageInfo.zip.file(
                    "xl/sharedStrings.xml",
                    serializeXml(
                        packageInfo.sharedStringsDoc
                    )
                );

                return packageInfo.zip
                    .generateAsync({
                        type: "binarystring",
                        compression: "DEFLATE",
                        compressionOptions: {
                            level: 6
                        }
                    })
                    .then(function (binary) {
                        writeBinaryFile(
                            outputPath,
                            binary
                        );

                        var result = {
                            success: true,
                            outputPath: outputPath,
                            sheetName:
                                packageInfo.sheetName,
                            values: preparedValues,
                            spacing:
                                spacingResults
                        };

                        logInfo(
                            "Excel 白皮生成成功：" +
                            outputPath
                        );

                        if (
                            typeof callback ===
                            "function"
                        ) {
                            callback(result, "");
                        }
                    });
            })
            .catch(function (error) {
                var message =
                    error && error.message
                        ? error.message
                        : String(error);

                if (
                    typeof callback === "function"
                ) {
                    callback(null, message);
                }
            });
    }

    return {
        analyze: analyze,
        generate: generate,
        prepareValues: prepareValues,
        splitGradeAndVolume:
            splitGradeAndVolume,
        getVisualUnits: getVisualUnits,
        _test: {
            detectFieldName:
                detectFieldName,
            getMissingFields:
                getMissingFields
        }
    };
})();
