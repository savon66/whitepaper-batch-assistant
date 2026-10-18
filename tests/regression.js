/* eslint-disable no-console */
"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const projectRoot = path.resolve(__dirname, "..");
const XLSX = require(
    path.join(projectRoot, "lib", "xlsx.full.min.js")
);

const messages = [];
const context = {
    console: console,
    XLSX: XLSX,
    Logger: {
        info: function (message) {
            messages.push("INFO " + message);
        },
        warn: function (message) {
            messages.push("WARN " + message);
        },
        error: function (message) {
            messages.push("ERROR " + message);
        },
        clear: function () {}
    },
    setTimeout: setTimeout,
    clearTimeout: clearTimeout
};

context.global = context;
vm.createContext(context);

function loadScript(relativePath) {
    const fullPath = path.join(projectRoot, relativePath);
    vm.runInContext(
        fs.readFileSync(fullPath, "utf8"),
        context,
        { filename: fullPath }
    );
}

loadScript("modules/fieldRegistry.js");

class FileReaderMock {
    readAsArrayBuffer(file) {
        try {
            const source = Buffer.from(file.buffer);
            const arrayBuffer = source.buffer.slice(
                source.byteOffset,
                source.byteOffset + source.byteLength
            );

            this.onload({
                target: {
                    result: arrayBuffer
                }
            });
        } catch (error) {
            this.error = error;

            if (this.onerror) {
                this.onerror();
            }
        }
    }
}

context.FileReader = FileReaderMock;
loadScript("modules/excelReader.js");
loadScript("modules/templateAnalyzer.js");
loadScript("modules/wordReplacer.js");
loadScript("modules/xlsxTemplateEngine.js");

function openWorkbook(buffer) {
    return new Promise(function (resolve, reject) {
        context.ExcelReader.openFile(
            { buffer: buffer },
            function (success, errorMessage) {
                if (success) {
                    resolve();
                } else {
                    reject(new Error(errorMessage));
                }
            }
        );
    });
}

function createFont() {
    return {
        Name: "方正黑体_GBK",
        NameFarEast: "方正黑体_GBK",
        Size: 16,
        Bold: 0,
        Italic: 0,
        Underline: 0,
        UnderlineColor: 0,
        Color: 0,
        Spacing: 0,
        Scaling: 100,
        Position: 0,
        StrikeThrough: 0,
        Subscript: 0,
        Superscript: 0
    };
}

class MockRange {
    constructor(paragraph, start, end) {
        this.paragraph = paragraph;
        this.Start = start;
        this.End = end;
        this.Font = paragraph.font;
    }

    get Duplicate() {
        return new MockRange(
            this.paragraph,
            this.Start,
            this.End
        );
    }

    get Text() {
        return this.paragraph.text.substring(
            this.Start,
            this.End
        );
    }

    set Text(value) {
        const text = String(value);
        this.paragraph.text =
            this.paragraph.text.substring(0, this.Start) +
            text +
            this.paragraph.text.substring(this.End);
        this.End = this.Start + text.length;
    }

    Information() {
        return this.Start * 8;
    }
}

class MockParagraph {
    constructor(text) {
        this.text = text + "\r";
        this.font = createFont();
    }

    get Range() {
        return new MockRange(
            this,
            0,
            this.text.length
        );
    }
}

function createMockDocument(paragraphTexts) {
    const paragraphs = paragraphTexts.map(function (text) {
        return new MockParagraph(text);
    });

    return {
        _paragraphs: paragraphs,
        ComputeStatistics: function () {
            return 1;
        },
        Paragraphs: {
            Count: paragraphs.length,
            Item: function (index) {
                return paragraphs[index - 1];
            }
        },
        PageSetup: {
            PageWidth: 595,
            PageHeight: 842,
            TopMargin: 72,
            BottomMargin: 72,
            LeftMargin: 72,
            RightMargin: 72,
            Orientation: 0
        },
        Content: {
            Font: createFont()
        }
    };
}

function makeWorkbook(sheets) {
    const workbook = XLSX.utils.book_new();

    sheets.forEach(function (sheetInfo) {
        XLSX.utils.book_append_sheet(
            workbook,
            XLSX.utils.aoa_to_sheet(sheetInfo.rows),
            sheetInfo.name
        );
    });

    return XLSX.write(workbook, {
        type: "buffer",
        bookType: "xlsx"
    });
}

async function run() {
    const excelBuffer = makeWorkbook([
            {
                name: "学考、寒暑假",
                rows: [
                    [
                        "教辅名称",
                        "类型",
                        "学科",
                        "使用年级",
                        "开本印张",
                        "教材印张",
                        "日期"
                    ],
                    [
                        "示例教辅 语文",
                        "考试辅导类",
                        "语文",
                        "九年级",
                        "大16开 20印张",
                        "",
                        "2026-08-10"
                    ],
                    ["测试书2", "考试辅导类", "数学", "九年级", "大16开 18印张", "", "2026-08-10"],
                    ["测试书3", "考试辅导类", "英语", "九年级", "大16开 19印张", "", "2026-08-10"],
                    ["测试书4", "考试辅导类", "物理", "九年级", "大16开 17印张", "", "2026-08-10"],
                    ["测试书5", "考试辅导类", "化学", "九年级", "大16开 16印张", "", "2026-08-10"]
                ]
            }
        ]);
    await openWorkbook(excelBuffer);

    const dataSheet = context.ExcelReader.detectSheet(
        ["06白皮制作单"],
        []
    );
    assert(dataSheet, "样例 Excel 应识别到业务工作表");
    assert.strictEqual(
        dataSheet.Name,
        "学考、寒暑假"
    );

    const headerMap =
        context.ExcelReader.getHeaderColumnMap(dataSheet);
    const mapping =
        context.WhitePaperFields.buildMapping(headerMap);
    const rows =
        context.ExcelReader.readDataRows(
            dataSheet,
            headerMap
        );

    assert.strictEqual(rows.length, 5);
    [
        "书名",
        "种类",
        "学科",
        "使用年级",
        "开本印张",
        "教材印张",
        "日期"
    ].forEach(function (fieldName) {
        assert(
            mapping[fieldName],
            "样例 Excel 未映射字段：" + fieldName
        );
    });
    assert.strictEqual(rows[0][mapping["日期"]], "2026-08-10");

    const paragraphs = [
            "示例省义务教育教辅材料评议",
            "送 评 样 书",
            "类    型：考试辅导类",
            "学    科：语文",
            "              使用年级：九年级",
            "开本印张：大16开 20印张",
            "教材印张：",
            "     2026年 8 月 10 日"
        ];
    const matchedFields = {};

    paragraphs.forEach(function (text, index) {
        const match =
            context.TemplateAnalyzer._test
                .matchParagraphText(text);

        if (match) {
            matchedFields[match.fieldName] = index + 1;
        }
    });

    [
        "种类",
        "学科",
        "使用年级",
        "开本印张",
        "教材印张",
        "日期"
    ].forEach(function (fieldName) {
        assert(
            matchedFields[fieldName],
            "样例 Word 未识别字段：" + fieldName
        );
    });
    assert.strictEqual(Object.keys(matchedFields).length, 6);

    assert.strictEqual(
        context.WhitePaperFields.resolveHeader("类    型："),
        "种类"
    );
    assert.strictEqual(
        context.WhitePaperFields.resolveHeader("开本印张"),
        "开本印张"
    );
    assert.strictEqual(
        context.WhitePaperFields.resolveHeader("教材印张"),
        "教材印张"
    );
    assert.strictEqual(
        context.WhitePaperFields.resolveHeader("印张"),
        "印张数"
    );
    assert.notStrictEqual(
        context.WhitePaperFields.resolveHeader("印张"),
        "开本印张"
    );

    [
        ["  种    类：同步练习类", "种类"],
        ["学 科:语文", "学科"],
        ["所配版本：人教版", "配套教材版本"],
        ["使用年级及册次：七年级 上册", "使用年级及册次"],
        ["印张数：12.5", "印张数"]
    ].forEach(function (testCase) {
        const match =
            context.WhitePaperFields
                .matchTemplateText(testCase[0]);
        assert(match, "旧模板字段应继续识别：" + testCase[0]);
        assert.strictEqual(match.fieldName, testCase[1]);
    });

    assert.strictEqual(
        context.WhitePaperFields.resolveHeader("编号"),
        "序号"
    );
    assert(
        !context.WhitePaperFields
            .getFillableFieldNames()
            .includes("序号")
    );
    assert(
        context.WhitePaperFields
            .shouldClearInOutput("序号")
    );

    const emptyValueFormat =
        context.WordReplacer._test
            .applyEmptyTemplateValueDefaults(
                {
                    name: "仿宋",
                    nameFarEast: "仿宋",
                    size: 12
                },
                { originalValue: "   " }
            );
    assert.strictEqual(emptyValueFormat.name, "黑体");
    assert.strictEqual(emptyValueFormat.nameFarEast, "黑体");
    assert.strictEqual(emptyValueFormat.size, 16);

    const existingValueFormat =
        context.WordReplacer._test
            .applyEmptyTemplateValueDefaults(
                {
                    name: "楷体",
                    nameFarEast: "楷体",
                    size: 14
                },
                { originalValue: "原值" }
            );
    assert.strictEqual(existingValueFormat.nameFarEast, "楷体");
    assert.strictEqual(existingValueFormat.size, 14);

    assert.strictEqual(
        context.XlsxTemplateEngine._test.detectFieldName(
            "  开 本 印 张：大16开 20印张"
        ),
        "开本印张"
    );
    assert.strictEqual(
        context.XlsxTemplateEngine._test.detectFieldName(
            "教材印张："
        ),
        "教材印张"
    );
    const preparedXlsxValues =
        context.XlsxTemplateEngine.prepareValues({
            "序号": 99,
            "类 型": "考试辅导类",
            "使用年级及册次": "七年级 上册",
            "日期": "2026-08-10",
            "开本印张": "大16开 20印张"
        });
    assert.strictEqual(preparedXlsxValues["序号"], "");
    assert.strictEqual(preparedXlsxValues["种类"], "考试辅导类");
    assert.strictEqual(preparedXlsxValues["使用年级"], "七年级");
    assert.strictEqual(preparedXlsxValues["册次"], "上册");
    assert.strictEqual(
        preparedXlsxValues["日期"],
        "2026年 8 月 10 日"
    );

    const mockDoc = createMockDocument(paragraphs);
    const analysis =
        context.TemplateAnalyzer.analyzeTemplate(mockDoc);
    assert.deepStrictEqual(
        Object.keys(analysis.fields).sort(),
        [
            "种类",
            "学科",
            "使用年级",
            "开本印张",
            "教材印张",
            "日期"
        ].sort()
    );

    const replacementValues = {
        "种类": "同步练习类",
        "学科": "数学",
        "使用年级": "八年级",
        "开本印张": "大16开 18印张",
        "教材印张": "15印张",
        "日期": "2026年 8 月 11 日"
    };

    Object.keys(replacementValues).forEach(function (fieldName) {
        const result = context.WordReplacer.replaceField(
            mockDoc,
            analysis.fields[fieldName],
            replacementValues[fieldName]
        );
        assert(result.success, "替换失败：" + fieldName);
    });

    const outputParagraphs = mockDoc._paragraphs.map(function (paragraph) {
        return paragraph.text.replace(/\r$/, "");
    });
    assert(outputParagraphs.includes("类    型：同步练习类"));
    assert(outputParagraphs.includes("学    科：数学"));
    assert(outputParagraphs.includes("              使用年级：八年级"));
    assert(outputParagraphs.includes("开本印张：大16开 18印张"));
    assert(outputParagraphs.includes("教材印张：15印张"));
    assert(outputParagraphs.includes("     2026年 8 月 11 日"));

    const numberedDoc = createMockDocument([
        "编号：3"
    ]);
    const numberedAnalysis =
        context.TemplateAnalyzer
            .analyzeTemplate(numberedDoc);
    assert(numberedAnalysis.fields["序号"]);
    const numberedTextLength =
        numberedDoc._paragraphs[0]
            .text.length;
    const clearNumberResult =
        context.WordReplacer
            .clearValueKeepUnderline(
            numberedDoc,
            numberedAnalysis.fields["序号"]
        );
    assert(clearNumberResult.success);
    assert(!clearNumberResult.skipped);
    assert.strictEqual(
        numberedDoc._paragraphs[0]
            .text.replace(/\r$/, ""),
        "编号： "
    );
    assert.strictEqual(
        numberedDoc._paragraphs[0]
            .text.length,
        numberedTextLength,
        "清除已有编号时不得改变字段字符总长度"
    );

    const blankNumberText =
        "编号：" + " ".repeat(18);
    const blankNumberDoc = createMockDocument([
        blankNumberText
    ]);
    const blankNumberAnalysis =
        context.TemplateAnalyzer
            .analyzeTemplate(blankNumberDoc);
    const blankNumberField =
        blankNumberAnalysis.fields["序号"];
    assert(blankNumberField);
    assert.strictEqual(
        blankNumberField.originalValue,
        ""
    );
    assert.strictEqual(
        blankNumberField.underlineTotalLength,
        18
    );
    const chongqingBlankNumberField =
        Object.assign(
            {},
            blankNumberField,
            {
                hasUnderline: true,
                underlineStartChar: 3,
                underlineEndChar: 21,
                valueStartChar: 21,
                valueEndChar: 21,
                leftPaddingCount: 18,
                rightPaddingCount: 0
            }
        );
    const blankNumberBefore =
        blankNumberDoc._paragraphs[0].text;
    const preserveBlankNumberResult =
        context.WordReplacer
            .clearValueKeepUnderline(
                blankNumberDoc,
                chongqingBlankNumberField
            );
    assert(preserveBlankNumberResult.success);
    assert(preserveBlankNumberResult.skipped);
    assert.strictEqual(
        blankNumberDoc._paragraphs[0].text,
        blankNumberBefore,
        "空编号横线必须逐字符保持原样，不能重建为固定 24 空格"
    );
    assert.strictEqual(
        preserveBlankNumberResult
            .fieldInfo
            .underlineTotalLength,
        18
    );

    const chongqingWorkbook = makeWorkbook([
                {
                    name: "06白皮制作单",
                    rows: [
                        ["白皮制作单"],
                        [
                            "序号*",
                            "书名*",
                            "种类",
                            "学科",
                            "使用年级及册次",
                            "印张数",
                            "配套教材版本"
                        ]
                    ].concat(
                        Array.from(
                            { length: 17 },
                            function (_, index) {
                                return [
                                    index + 1,
                                    index === 8
                                        ? "示例教辅 化学 大一轮总复习"
                                        : "测试书" + (index + 1),
                                    "考试辅导类",
                                    index === 8 ? "化学" : "物理",
                                    "高中三年级一轮",
                                    20,
                                    ""
                                ];
                            }
                        )
                    )
                }
            ]);
    await openWorkbook(chongqingWorkbook);
    const chongqingSheet =
        context.ExcelReader.detectSheet(
            ["06白皮制作单"],
            []
        );
    assert(chongqingSheet);
    const chongqingHeaderMap =
        context.ExcelReader
            .getHeaderColumnMap(
                chongqingSheet
            );
    const chongqingMapping =
        context.WhitePaperFields
            .buildMapping(
                chongqingHeaderMap
            );
    const chongqingRows =
        context.ExcelReader.readDataRows(
            chongqingSheet,
            chongqingHeaderMap
        );
    assert.strictEqual(chongqingRows.length, 17);
    assert(chongqingMapping["序号"]);
    assert.strictEqual(
        String(
            chongqingRows[8]["序号*"]
        ),
        "9"
    );
    assert.strictEqual(
        chongqingRows[8]["书名*"],
        "示例教辅 化学 大一轮总复习"
    );

    context.window = {
        confirm: function () {
            return true;
        }
    };
    context.document = {
        getElementById: function () {
            return null;
        }
    };
    context.navigator = {};
    context.Alignment = {};
    context.FileManager = {};
    context.Validator = {};
    loadScript("taskpane.js");

    context.state.templateType = "word";
    context.state.templateReady = true;
    context.state.templateInfo = {
        fields: analysis.fields
    };
    context.state.fieldMapping = mapping;
    context.state.dataRows = rows;

    const compatibility =
        context.buildTemplateCompatibilityReport();
    assert(compatibility.compatible);
    assert.strictEqual(
        compatibility.missingFields.length,
        0
    );
    assert.strictEqual(
        context.getMappedFieldValue(rows[0], "日期"),
        "2026年 8 月 10 日"
    );
    assert.strictEqual(
        context.getMappedFieldValue(rows[0], "开本印张"),
        "大16开 20印张"
    );

    const completeMapping = context.state.fieldMapping;
    const missingOpenFormat = Object.assign({}, completeMapping);
    delete missingOpenFormat["开本印张"];
    context.state.fieldMapping = missingOpenFormat;
    const blocked = context.buildTemplateCompatibilityReport();
    assert(!blocked.compatible);
    assert(blocked.missingFields.includes("开本印张"));

    const dateDefaultMapping = Object.assign({}, completeMapping);
    delete dateDefaultMapping["日期"];
    context.state.fieldMapping = dateDefaultMapping;
    const keepsTemplateDate =
        context.buildTemplateCompatibilityReport();
    assert(keepsTemplateDate.compatible);
    assert(keepsTemplateDate.defaultFields.includes("日期"));
    context.state.fieldMapping = completeMapping;

    const anhuiWorkbook = makeWorkbook([
            {
                name: "同步",
                rows: [
                    ["示例送审白皮样书封面信息"],
                    [
                        "序号",
                        "书名",
                        "责编",
                        "尺寸",
                        "种类",
                        "学科",
                        "使用年级",
                        "册次",
                        "印张数",
                        "配套教材版本"
                    ],
                    [
                        1,
                        "示例教辅 高中英语 必修 第一册",
                        "责编甲",
                        "210mm×297mm",
                        "同步练习类",
                        "英语",
                        "高一",
                        "必修 第一册",
                        14,
                        "示例出版社出版"
                    ],
                    [2, "测试书2", "责编", "210mm×297mm", "同步练习类", "英语", "高一", "必修 第二册", 13.5, "示例出版社出版"],
                    [3, "测试书3", "责编", "210mm×297mm", "同步练习类", "英语", "高一", "必修 第三册", 13.5, "示例出版社出版"],
                    [4, "测试书4", "责编", "210mm×297mm", "同步练习类", "英语", "高一", "选择性必修 第一册", 13.5, "示例出版社出版"],
                    [5, "测试书5", "责编", "210mm×297mm", "同步练习类", "英语", "高二", "选择性必修 第二册", 14, "示例出版社出版"],
                    [6, "测试书6", "责编", "210mm×297mm", "同步练习类", "英语", "高二", "选择性必修 第三册", 14, "示例出版社出版"],
                    [7, "测试书7", "责编", "210mm×297mm", "同步练习类", "英语", "高二", "选择性必修 第四册", 14, "示例出版社出版"],
                    [8, "测试书8", "责编", "210mm×297mm", "毕业年级考试辅导类", "化学", "高三", "全一册", 32, ""],
                    [9, "测试书9", "责编", "210mm×297mm", "毕业年级考试辅导类", "思想政治", "高三", "全一册", 32, ""]
                ]
            }
        ]);
    await openWorkbook(anhuiWorkbook);
    const anhuiSheet =
        context.ExcelReader.detectSheet([], []);
    assert(anhuiSheet);
    assert.strictEqual(anhuiSheet.Name, "同步");
    const anhuiHeaderMap =
        context.ExcelReader
            .getHeaderColumnMap(anhuiSheet);
    const anhuiMapping =
        context.WhitePaperFields
            .buildMapping(anhuiHeaderMap);
    const anhuiRows =
        context.ExcelReader.readDataRows(
            anhuiSheet,
            anhuiHeaderMap
        );
    assert.strictEqual(anhuiRows.length, 9);
    assert(anhuiMapping["序号"]);

    const anhuiTemplateTexts = [
            "编号：             ",
            "          种类：            同步练习类            ",
            "          学科：               英语               ",
            "          使用年级：           高一               ",
            "          册次：           必修 第一册            ",
            "          印张数：              14                ",
            "          配套教材版本： 示例出版社出版   "
        ];
    const anhuiTemplateFields = {};

    anhuiTemplateTexts.forEach(function (text) {
        const fieldName =
            context.XlsxTemplateEngine
                ._test.detectFieldName(text);

        if (fieldName) {
            anhuiTemplateFields[fieldName] = {
                name: fieldName
            };
        }
    });
    [
        "序号",
        "种类",
        "学科",
        "使用年级",
        "册次",
        "印张数",
        "配套教材版本"
    ].forEach(function (fieldName) {
        assert(
            anhuiTemplateFields[fieldName],
            "安徽模板未识别字段：" + fieldName
        );
    });

    context.state.templateType = "xlsx";
    context.state.templateReady = true;
    context.state.templateInfo = null;
    context.state.xlsxTemplateInfo = {
        fields: anhuiTemplateFields
    };
    context.state.fieldMapping = anhuiMapping;
    context.state.dataRows = anhuiRows;

    const anhuiCompatibility =
        context.buildTemplateCompatibilityReport();
    assert(anhuiCompatibility.compatible);
    assert(
        !anhuiCompatibility.templateFields
            .includes("序号")
    );
    const anhuiValues =
        context.buildXlsxTemplateValues(
            anhuiRows[2]
        );
    assert.strictEqual(anhuiValues["序号"], "");
    assert.strictEqual(anhuiValues["册次"], "必修 第三册");
    assert.strictEqual(anhuiValues["印张数"], "13.5");

    const mixedWorkbook = makeWorkbook([
        {
            name: "06白皮制作单",
            rows: [
                ["名称", "备注"],
                ["无关数据", "不应被误选"]
            ]
        },
        {
            name: "真实数据",
            rows: [
                ["教辅名称", "类型", "学科"],
                ["测试书", "考试辅导类", "语文"]
            ]
        }
    ]);
    await openWorkbook(mixedWorkbook);
    const structurallySelected =
        context.ExcelReader.detectSheet(
            ["06白皮制作单"],
            []
        );
    assert(structurallySelected);
    assert.strictEqual(structurallySelected.Name, "真实数据");

    const invalidWorkbook = makeWorkbook([
        {
            name: "06白皮制作单",
            rows: [
                ["名称", "备注"],
                ["无关数据", "不应通过"]
            ]
        }
    ]);
    await openWorkbook(invalidWorkbook);
    assert.strictEqual(
        context.ExcelReader.detectSheet(
            ["06白皮制作单"],
            []
        ),
        null
    );

    console.log("v1.0 regression: PASS");
    console.log("- Excel rows: " + rows.length);
    console.log(
        "- Word fields: " +
        Object.keys(analysis.fields).join(", ")
    );
    console.log("- Template/data compatibility gate: PASS");
    console.log("- Excel-template rows: " + anhuiRows.length);
    console.log("- Sequence/number output clearing: PASS");
    console.log("- Number-line fidelity: PASS");
    console.log("- Empty-line default font (黑体三号): PASS");
    console.log("- Legacy field compatibility: PASS");
    console.log("- False sheet selection prevention: PASS");
}

run().catch(function (error) {
    console.error("v1.0 regression: FAIL");
    console.error(error && error.stack ? error.stack : error);
    process.exitCode = 1;
});
