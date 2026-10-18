/**
 * 白皮批量制作助手 - 任务窗格主逻辑
 *
 * 依赖模块：
 * Logger、ExcelReader、TemplateAnalyzer、WordReplacer、Alignment、
 * FileManager、Validator
 */

// ==================== 全局状态 ====================
var state = {
    excelPath: "",
    excelFile: null,
    templatePath: "",
    templateType: "",
    templateReady: false,
    compatibilityReady: false,
    compatibilityReport: null,
    templateSheetName: "",
    xlsxTemplateInfo: null,
    outputPath: "",
    sheetName: "",
    fieldMapping: {},
    dataRows: [],
    templateInfo: null,
    templateFontWarnings: [],
    templateFontConfirmed: false,
    baseline: null,
    isRunning: false,
    cancelRequested: false,
    successCount: 0,
    failCount: 0,
    results: []
};

var wordApp = null;

if (
    typeof WhitePaperFields === "undefined" ||
    !WhitePaperFields
) {
    throw new Error(
        "字段注册表未加载，请检查 modules/fieldRegistry.js"
    );
}

/*
 * 白皮实际写入字段的标准字形。中文字号“三号”等于 16 磅。
 * 检测只提示、不擅自改模板，用户确认后仍可继续生成。
 */
/* Excel 白皮模板会把“使用年级及册次”拆成两个独立字段。 */
var WHITE_PAPER_FONT_CHECK_FIELDS =
    WhitePaperFields.getFillableFieldNames();

var EXPECTED_FILL_FONT_NAME = "黑体";
var EXPECTED_FILL_FONT_SIZE = 16;

function normalizeFontName(fontName) {
    return String(
        fontName === undefined ||
        fontName === null
            ? ""
            : fontName
    )
        .replace(/[\s_-]+/g, "")
        .toLowerCase();
}

function isExpectedBlackFont(fontName) {
    var normalized =
        normalizeFontName(fontName);

    return (
        normalized.indexOf("黑体") !== -1 ||
        normalized === "simhei" ||
        normalized.indexOf("heitisc") !== -1 ||
        normalized.indexOf("stheiti") !== -1
    );
}

function getFieldFontDisplayName(format) {
    var source = format || {};
    var farEast = source.nameFarEast;
    var normal = source.name;

    if (
        farEast !== undefined &&
        farEast !== null &&
        String(farEast).trim() &&
        String(farEast) !== "9999999"
    ) {
        return String(farEast).trim();
    }

    if (
        normal !== undefined &&
        normal !== null &&
        String(normal).trim() &&
        String(normal) !== "9999999"
    ) {
        return String(normal).trim();
    }

    return "无法确定";
}

function collectTemplateFontWarnings(fields) {
    var result = [];
    var source = fields || {};

    for (
        var i = 0;
        i < WHITE_PAPER_FONT_CHECK_FIELDS.length;
        i++
    ) {
        var fieldName =
            WHITE_PAPER_FONT_CHECK_FIELDS[i];
        var fieldInfo =
            source[fieldName];

        if (!fieldInfo) {
            continue;
        }

        var format =
            fieldInfo.format || {};
        var fontName =
            getFieldFontDisplayName(format);
        var size = Number(format.size);
        var fontOk =
            isExpectedBlackFont(fontName);
        var sizeOk =
            isFinite(size) &&
            Math.abs(
                size - EXPECTED_FILL_FONT_SIZE
            ) <= 0.1;

        if (!fontOk || !sizeOk) {
            result.push({
                fieldName: fieldName,
                fontName: fontName,
                fontSize:
                    isFinite(size) && size > 0
                        ? size
                        : null,
                fontOk: fontOk,
                sizeOk: sizeOk
            });
        }
    }

    return result;
}

function buildTemplateFontWarningMessage(
    warnings,
    templateLabel
) {
    var lines = [];
    var items = warnings || [];

    lines.push(
        "检测到" +
        (templateLabel || "模板") +
        "中以下待填写字段不是黑体三号（16 磅）："
    );
    lines.push("");

    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        var sizeText =
            item.fontSize === null
                ? "字号无法确定"
                : item.fontSize + " 磅";

        lines.push(
            "• " +
            item.fieldName +
            "：" +
            item.fontName +
            " / " +
            sizeText
        );
    }

    lines.push("");
    lines.push(
        "标准要求：黑体、三号（16 磅）。"
    );
    lines.push(
        "是否确认仍使用当前模板继续？"
    );

    return lines.join("\n");
}

function inspectAndConfirmTemplateFonts(
    fields,
    templateLabel
) {
    state.templateFontWarnings =
        collectTemplateFontWarnings(
            fields
        );

    if (
        state.templateFontWarnings.length === 0
    ) {
        state.templateFontConfirmed = true;
        return true;
    }

    Logger.warn(
        "模板待填字段字体不符合黑体三号要求，共 " +
        state.templateFontWarnings.length +
        " 项"
    );

    state.templateFontConfirmed =
        window.confirm(
            buildTemplateFontWarningMessage(
                state.templateFontWarnings,
                templateLabel
            )
        );

    return state.templateFontConfirmed;
}

function ensureTemplateFontConfirmed() {
    if (
        !state.templateFontWarnings ||
        state.templateFontWarnings.length === 0 ||
        state.templateFontConfirmed
    ) {
        return true;
    }

    state.templateFontConfirmed =
        window.confirm(
            buildTemplateFontWarningMessage(
                state.templateFontWarnings,
                getTemplateTypeLabel(
                    state.templateType
                )
            )
        );

    return state.templateFontConfirmed;
}

// ==================== 初始化 ====================

window.onload = function () {
    Logger.info(
        "白皮批量制作助手任务窗格加载完成"
    );

    initWordApp();
    initExcelFilePicker();
};

/**
 * 获取当前任务窗格所属的 WPS 文字 Application 对象
 */
function getWpsApplication() {
    try {
        if (
            typeof Application !== "undefined" &&
            Application
        ) {
            return Application;
        }
    } catch (e1) {
        Logger.warn(
            "直接获取 Application 失败: " +
            e1.message
        );
    }

    try {
        if (
            typeof wps !== "undefined" &&
            wps &&
            typeof wps.WpsApplication === "function"
        ) {
            var app = wps.WpsApplication();

            if (app) {
                return app;
            }
        }
    } catch (e2) {
        Logger.warn(
            "通过 wps.WpsApplication 获取失败: " +
            e2.message
        );
    }

    throw new Error(
        "无法获取当前 WPS 文字 Application 对象"
    );
}

/**
 * 连接当前 WPS 文字实例
 */
function initWordApp() {
    try {
        wordApp = getWpsApplication();

        Logger.info(
            "已连接当前 WPS 文字实例"
        );
    } catch (e) {
        wordApp = null;

        Logger.error(
            "连接 WPS 文字失败",
            e.message
        );

        alert(
            "连接 WPS 文字组件失败：\n" +
            e.message
        );
    }
}
/**
 * 获取文件扩展名
 */
function getTemplateExtension(
    filePath
) {
    var path = String(
        filePath || ""
    );

    var queryIndex =
        path.indexOf("?");

    if (queryIndex >= 0) {
        path =
            path.substring(
                0,
                queryIndex
            );
    }

    var slashIndex =
        Math.max(
            path.lastIndexOf("\\"),
            path.lastIndexOf("/")
        );

    var dotIndex =
        path.lastIndexOf(".");

    if (
        dotIndex < 0 ||
        dotIndex < slashIndex
    ) {
        return "";
    }

    return path
        .substring(dotIndex)
        .toLowerCase();
}

/**
 * 根据扩展名识别模板类型
 */
function detectTemplateType(
    filePath
) {
    var extension =
        getTemplateExtension(
            filePath
        );

    if (
        extension === ".doc" ||
        extension === ".docx"
    ) {
        return "word";
    }

    if (extension === ".xlsx") {
        return "xlsx";
    }

    throw new Error(
        "不支持的模板格式：" +
        (
            extension ||
            "无扩展名"
        )
    );
}

/**
 * 获取模板类型显示文字
 */
function getTemplateTypeLabel(
    templateType
) {
    if (templateType === "word") {
        return "Word 模板（.doc/.docx）";
    }

    if (templateType === "xlsx") {
        return "Excel 模板（.xlsx）";
    }

    return "未识别";
}

/**
 * 清除上一次模板的识别状态
 */
function resetTemplateAnalysisState() {
    state.templateReady = false;
    state.compatibilityReady = false;
    state.compatibilityReport = null;
    state.templateInfo = null;
    state.templateFontWarnings = [];
    state.templateFontConfirmed = false;
    state.baseline = null;
    state.templateSheetName = "";
    state.xlsxTemplateInfo = null;
}

/**
 * 更新模板类型显示区域
 */
function updateTemplateTypeUI() {
    var typeRow =
        document.getElementById(
            "templateTypeRow"
        );

    var typeText =
        document.getElementById(
            "templateTypeText"
        );

    var sheetRow =
        document.getElementById(
            "templateSheetRow"
        );

    var sheetName =
        document.getElementById(
            "templateSheetName"
        );

    if (typeRow) {
        typeRow.style.display =
            state.templateType
                ? "flex"
                : "none";
    }

    if (typeText) {
        typeText.textContent =
            getTemplateTypeLabel(
                state.templateType
            );
    }

    if (
        sheetRow &&
        sheetName
    ) {
        if (
            state.templateType ===
                "xlsx" &&
            state.templateSheetName
        ) {
            sheetRow.style.display =
                "flex";

            sheetName.textContent =
                state.templateSheetName;
        } else {
            sheetRow.style.display =
                "none";

            sheetName.textContent =
                "-";
        }
    }
}
/**
 * 初始化隐藏的浏览器 Excel 文件选择器
 *
 * taskpane.html 中必须存在：
 *
 * <input
 *     type="file"
 *     id="excelFilePicker"
 *     accept=".xlsx,.xls"
 *     style="display:none;"
 * >
 */
function initExcelFilePicker() {
    var picker = document.getElementById(
        "excelFilePicker"
    );

    if (!picker) {
        Logger.error(
            "初始化 Excel 选择器失败",
            "taskpane.html 中不存在 excelFilePicker"
        );

        return;
    }

    picker.onchange = function (event) {
        var files =
            event.target.files;

        if (
            !files ||
            files.length === 0
        ) {
            return;
        }

        var file = files[0];

        state.excelFile = file;
        state.excelPath = file.name;

        var input =
            document.getElementById(
                "excelPath"
            );

        if (input) {
            input.value = file.name;
        }

        Logger.info(
            "已选择 Excel 文件: " +
            file.name
        );

        analyzeExcel();
    };
}

// ==================== 文件选择 ====================

/**
 * 选择 Excel 数据文件
 */
function selectExcel() {
    try {
        var picker =
            document.getElementById(
                "excelFilePicker"
            );

        if (!picker) {
            throw new Error(
                "未找到 Excel 文件选择控件"
            );
        }

        /*
         * 清空旧值。
         * 连续选择同一个文件时仍触发 onchange。
         */
        picker.value = "";
        picker.click();

    } catch (e) {
        Logger.error(
            "选择 Excel 文件失败",
            e.message
        );

        alert(
            "选择 Excel 文件失败：\n" +
            e.message
        );
    }
}

/**
 * 选择 Word 模板文件
 */
/**
 * 选择 Word 或 Excel 白皮模板
 */
function selectTemplate() {
    try {
        var app =
            wordApp ||
            getWpsApplication();

        // 3：文件选择器
        var dialog =
            app.FileDialog(3);

        if (!dialog) {
            throw new Error(
                "WPS 无法创建文件选择窗口"
            );
        }

        dialog.Title =
            "选择白皮模板文件";

        dialog.AllowMultiSelect =
            false;

        try {
            dialog.Filters.Clear();

            dialog.Filters.Add(
                "白皮模板文件",
                "*.doc;*.docx;*.xlsx"
            );

            dialog.Filters.Add(
                "Word 模板",
                "*.doc;*.docx"
            );

            dialog.Filters.Add(
                "Excel 模板",
                "*.xlsx"
            );
        } catch (filterError) {
            Logger.warn(
                "设置模板文件筛选器失败: " +
                filterError.message
            );
        }

        var showResult =
            dialog.Show();

        if (showResult !== -1) {
            return;
        }

        if (
            !dialog.SelectedItems ||
            dialog.SelectedItems.Count < 1
        ) {
            return;
        }

        var selectedPath =
            String(
                dialog.SelectedItems.Item(1)
            );

        var selectedType =
            detectTemplateType(
                selectedPath
            );

        /*
         * 只有通过格式检查后，
         * 才正式写入状态。
         */
        state.templatePath =
            selectedPath;

        state.templateType =
            selectedType;

        resetTemplateAnalysisState();

        var input =
            document.getElementById(
                "templatePath"
            );

        if (input) {
            input.value =
                state.templatePath;
        }

        updateTemplateTypeUI();

        Logger.info(
            "已选择模板：" +
            state.templatePath +
            "，类型：" +
            state.templateType
        );

        /*
         * 根据模板类型进入不同分析流程。
         * Excel 模板绝不能交给
         * wordApp.Documents.Open()。
         */
        analyzeSelectedTemplate();

    } catch (e) {
        resetTemplateAnalysisState();
        checkCanStart();

        Logger.error(
            "选择模板文件失败",
            e.message
        );

        alert(
            "选择模板文件失败：\n" +
            e.message
        );
    }
}

/**
 * 选择输出文件夹
 */
function selectOutput() {
    try {
        var app =
            wordApp ||
            getWpsApplication();

        // 4：文件夹选择器
        var dialog =
            app.FileDialog(4);

        if (!dialog) {
            throw new Error(
                "WPS 无法创建文件夹选择窗口"
            );
        }

        dialog.Title =
            "选择输出文件夹";

        dialog.AllowMultiSelect =
            false;

        if (dialog.Show() === -1) {
            if (
                dialog.SelectedItems &&
                dialog.SelectedItems.Count >= 1
            ) {
                state.outputPath =
                    String(
                        dialog.SelectedItems.Item(1)
                    );

                var input =
                    document.getElementById(
                        "outputPath"
                    );

                if (input) {
                    input.value =
                        state.outputPath;
                }

                checkCanStart();
            }
        }

    } catch (e) {
        Logger.error(
            "选择输出文件夹失败",
            e.message
        );

        alert(
            "选择输出文件夹失败：\n" +
            e.message
        );
    }
}

// ==================== Excel 分析 ====================

/**
 * 分析 Excel 文件，识别工作表、字段和数据行
 */
function analyzeExcel() {
    if (!state.excelFile) {
        alert(
            "请先选择 Excel 数据文件"
        );

        return;
    }

    ExcelReader.openFile(
        state.excelFile,

        function (
            opened,
            errorMessage
        ) {
            if (!opened) {
                alert(
                    "无法打开 Excel 文件：\n" +
                    (
                        errorMessage ||
                        "未知错误"
                    )
                );

                return;
            }

            try {
                var prioritySheets = [
                    "06白皮制作单"
                ];

                var requiredHeaders =
                    getCurrentTemplateFieldNames();

                var sheet =
                    ExcelReader.detectSheet(
                        prioritySheets,
                        requiredHeaders
                    );

                if (!sheet) {
                    ExcelReader.close();

                    state.sheetName = "";
                    state.fieldMapping = {};
                    state.dataRows = [];
                    state.compatibilityReady = false;
                    state.compatibilityReport = null;

                    updateRecognitionUI();
                    updateCompatibilityUI();
                    checkCanStart();

                    alert(
                        "未找到符合要求的工作表。\n\n" +
                        "工作表必须包含“书名/教辅名称”或“序号”，" +
                        "并至少包含一个可填写字段。\n\n" +
                        "v1.0 会自动兼容字段别名和字符间空格，" +
                        "但不会再用低相似度规则误选无关工作表。"
                    );

                    return;
                }

                state.sheetName =
                    sheet.Name || "";

                var headerMap =
                    ExcelReader
                        .getHeaderColumnMap(
                            sheet
                        );

                state.dataRows =
                    ExcelReader
                        .readDataRows(
                            sheet,
                            headerMap
                        );

                state.fieldMapping =
                    buildFieldMapping(
                        headerMap
                    );

                ExcelReader.close();

                refreshTemplateCompatibility(
                    false
                );
                updateRecognitionUI();
                updateCompatibilityUI();
                checkCanStart();

                Logger.info(
                    "Excel 分析完成，识别到 " +
                    state.dataRows.length +
                    " 条有效数据"
                );

            } catch (e) {
                ExcelReader.close();

                Logger.error(
                    "分析 Excel 失败",
                    e.message
                );

                alert(
                    "分析 Excel 失败：\n" +
                    e.message
                );
            }
        }
    );
}

/**
 * 建立 Excel 表头与标准字段名映射
 */
function getStandardFieldAliases(
    standardField
) {
    return WhitePaperFields
        .getAliases(standardField);
}

/**
 * 按标准字段别名查找真实 Excel 表头。
 * 优先精确匹配；最后才使用保守的包含匹配。
 */
function findHeaderByAliases(
    headerMap,
    standardField
) {
    return WhitePaperFields.findHeader(
        headerMap,
        standardField
    );
}

/**
 * 建立 Excel 表头与标准字段名映射
 */
function buildFieldMapping(
    headerMap
) {
    return WhitePaperFields
        .buildMapping(headerMap);
}

function getOrderedTemplateFieldNames(fields) {
    var source = fields || {};
    var result = [];
    var ordered =
        WhitePaperFields
            .getFillableFieldNames();

    for (var i = 0; i < ordered.length; i++) {
        if (source[ordered[i]]) {
            result.push(ordered[i]);
        }
    }

    return result;
}

function getCurrentTemplateFields() {
    if (
        state.templateType === "xlsx" &&
        state.xlsxTemplateInfo
    ) {
        return state.xlsxTemplateInfo.fields || {};
    }

    if (state.templateInfo) {
        return state.templateInfo.fields || {};
    }

    return {};
}

function getCurrentTemplateFieldNames() {
    return getOrderedTemplateFieldNames(
        getCurrentTemplateFields()
    );
}

function buildTemplateCompatibilityReport() {
    var templateFields =
        getCurrentTemplateFieldNames();
    var mapping =
        state.fieldMapping || {};
    var report = {
        templateFields: templateFields,
        matchedFields: {},
        defaultFields: [],
        missingFields: [],
        hasIdentity: Boolean(
            mapping["书名"] ||
            mapping["序号"]
        ),
        compatible: false,
        status: "waiting"
    };

    if (!state.templateReady) {
        report.status = "template-not-ready";
        return report;
    }

    if (
        !state.dataRows ||
        state.dataRows.length === 0
    ) {
        report.status = "data-not-ready";
        return report;
    }

    if (!report.hasIdentity) {
        report.missingFields.push(
            "书名（或序号）"
        );
    }

    for (
        var i = 0;
        i < templateFields.length;
        i++
    ) {
        var fieldName =
            templateFields[i];
        var sourceFields =
            WhitePaperFields
                .getCompatibleSourceFields(
                    fieldName,
                    mapping
                );

        if (sourceFields.length > 0) {
            report.matchedFields[fieldName] =
                sourceFields;
            continue;
        }

        if (
            WhitePaperFields
                .allowsTemplateDefault(
                    fieldName
                )
        ) {
            report.defaultFields.push(
                fieldName
            );
            continue;
        }

        report.missingFields.push(
            fieldName
        );
    }

    report.compatible =
        report.missingFields.length === 0 &&
        templateFields.length > 0;
    report.status = report.compatible
        ? "compatible"
        : "missing-fields";

    return report;
}

function refreshTemplateCompatibility(
    showAlert
) {
    state.compatibilityReport =
        buildTemplateCompatibilityReport();
    state.compatibilityReady = Boolean(
        state.compatibilityReport &&
        state.compatibilityReport.compatible
    );

    updateCompatibilityUI();

    if (
        showAlert &&
        state.compatibilityReport &&
        state.compatibilityReport.status ===
            "missing-fields"
    ) {
        alert(
            "数据表与当前模板不兼容。\n\n" +
            "缺少字段：\n" +
            state.compatibilityReport
                .missingFields
                .join("、") +
            "\n\n请补齐相应 Excel 表头。" +
            "v1.0 不会再让字段不完整的数据表误通过。"
        );
    }

    return state.compatibilityReady;
}

// ==================== 模板分析 ====================

/**
 * 根据模板类型调用对应分析流程
 */
function analyzeSelectedTemplate() {
    resetTemplateAnalysisState();
    updateTemplateTypeUI();
    checkCanStart();

    if (state.templateType === "word") {
        analyzeWordTemplate();
        return;
    }

    if (state.templateType === "xlsx") {
        analyzeXlsxTemplate();
        return;
    }

    alert("无法识别当前模板类型");
}

/**
 * 分析 Word 模板
 */
function analyzeWordTemplate() {
    var doc = null;

    try {
        var app =
            wordApp ||
            getWpsApplication();

        doc = app.Documents.Open(
            state.templatePath,
            false,
            true
        );

        state.templateInfo =
            TemplateAnalyzer.analyzeTemplate(
                doc
            );

        state.baseline =
            Validator.recordBaseline(
                doc
            );

        var fieldCount =
            Object.keys(
                state.templateInfo.fields ||
                {}
            ).length;

        state.templateReady =
            state.templateInfo.success !== false &&
            fieldCount > 0;

        if (state.templateReady) {
            inspectAndConfirmTemplateFonts(
                state.templateInfo.fields || {},
                "Word 模板"
            );
        }

        Logger.info(
            "Word 模板分析完成，识别到 " +
            fieldCount +
            " 个字段"
        );

        updateTemplateTypeUI();
        refreshTemplateCompatibility(
            state.dataRows.length > 0
        );
        updateRecognitionUI();
        checkCanStart();

        if (!state.templateReady) {
            var diagnostics =
                state.templateInfo.diagnostics || {};
            var candidates =
                diagnostics.unrecognizedLabels || [];
            var candidateText = "";

            if (candidates.length > 0) {
                var labels = [];

                for (
                    var c = 0;
                    c < candidates.length && c < 8;
                    c++
                ) {
                    labels.push(
                        "• 第" +
                        candidates[c]
                            .paragraphIndex +
                        "段：" +
                        String(
                            candidates[c].label ||
                            candidates[c].text ||
                            ""
                        ).trim()
                    );
                }

                candidateText =
                    "\n\n检测到但未识别的标签：\n" +
                    labels.join("\n");
            }

            alert(
                "Word 模板中没有识别到可填写字段。\n\n" +
                "v1.0 已支持标签前导空格、汉字间空格、" +
                "中英文冒号和常见字段别名。" +
                candidateText +
                "\n\n请检查这些标签是否属于新的业务字段。"
            );
        }

        if (
            state.templateInfo &&
            state.templateInfo.missingFonts &&
            state.templateInfo.missingFonts.length > 0
        ) {
            alert(
                "警告：模板使用了以下未安装的字体：\n\n" +
                state.templateInfo.missingFonts.join("\n") +
                "\n\n可能导致格式偏差，建议安装对应字体后再使用。"
            );
        }

    } catch (e) {
        state.templateReady = false;

        Logger.error(
            "分析 Word 模板失败",
            e.message
        );

        alert(
            "分析 Word 模板失败：\n" +
            e.message
        );

        checkCanStart();

    } finally {
        try {
            if (doc) {
                doc.Close(false);
            }
        } catch (closeError) {
            Logger.warn(
                "关闭模板文档失败: " +
                closeError.message
            );
        }
    }
}

/**
 * 分析 Excel 白皮模板
 *
 * 当前先完成模板类型分流，
 * 下一步再接入 XlsxTemplateEngine 的实际识别逻辑。
 */
/**
 * 分析 Excel 白皮模板
 */
function analyzeXlsxTemplate() {
    state.templateReady = false;
    state.xlsxTemplateInfo = null;
    state.templateSheetName = "";

    updateTemplateTypeUI();
    checkCanStart();

    if (
        typeof XlsxTemplateEngine ===
        "undefined"
    ) {
        Logger.error(
            "Excel 模板引擎未加载",
            "未找到 XlsxTemplateEngine"
        );

        alert(
            "Excel 模板识别模块未加载。\n\n" +
            "请检查以下文件是否存在且不是空文件：\n" +
            "lib\\jszip.min.js\n" +
            "modules\\xlsxTemplateEngine.js"
        );

        return;
    }

    if (typeof JSZip === "undefined") {
        Logger.error(
            "JSZip 未加载",
            "未找到全局对象 JSZip"
        );

        alert(
            "JSZip 压缩库未加载。\n\n" +
            "请检查：\n" +
            "lib\\jszip.min.js"
        );

        return;
    }

    Logger.info(
        "开始分析 Excel 白皮模板：" +
        state.templatePath
    );

    XlsxTemplateEngine.analyze(
        state.templatePath,

        function (result, errorMessage) {
            if (!result) {
                state.templateReady = false;
                state.xlsxTemplateInfo = null;
                state.templateSheetName = "";

                updateTemplateTypeUI();
                checkCanStart();

                Logger.error(
                    "Excel 模板分析失败",
                    errorMessage || "未知错误"
                );

                alert(
                    "Excel 模板分析失败：\n" +
                    (errorMessage || "未知错误")
                );

                return;
            }

            state.xlsxTemplateInfo = result;
            state.templateSheetName =
                result.sheetName || "";

            state.templateReady =
                Boolean(
                    result.success &&
                    (result.fieldCount || 0) > 0
                );

            if (state.templateReady) {
                inspectAndConfirmTemplateFonts(
                    result.fields || {},
                    "Excel 模板"
                );
            }

            updateTemplateTypeUI();
            refreshTemplateCompatibility(
                state.dataRows.length > 0
            );
            updateRecognitionUI();
            checkCanStart();

            var fieldCount =
                result.fieldCount || 0;

            Logger.info(
                "Excel 模板分析完成，工作表：" +
                (state.templateSheetName || "未识别") +
                "，字段数：" +
                fieldCount +
                "，可生成：" +
                (state.templateReady ? "是" : "否")
            );

            if (!state.templateReady) {
                alert(
                    "Excel 模板中没有识别到可填写字段。\n\n" +
                    "工作表：" +
                    (state.templateSheetName || "未识别") +
                    "\n\nv1.0 会按模板实际字段自适应，" +
                    "不再强制某一套固定字段组合。"
                );
                return;
            }

            alert(
                "Excel 模板识别成功！\n\n" +
                "工作表：" +
                state.templateSheetName +
                "\n识别字段：" +
                fieldCount +
                " 个\n\n" +
                "现在可以预览或批量生成 xlsx 文件。"
            );
        }
    );
}

// ==================== UI 更新 ====================

/**
 * 对文本进行 HTML 转义，避免识别结果或错误日志破坏页面结构。
 */
function escapeHtml(value) {
    return String(
        value === undefined ||
        value === null
            ? ""
            : value
    )
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

/**
 * 显示数据表与当前模板的兼容状态。
 * 只有模板实际需要、且数据能够提供的字段才允许开始生成。
 */
function updateCompatibilityUI() {
    var row =
        document.getElementById(
            "compatibilityRow"
        );
    var statusElement =
        document.getElementById(
            "compatibilityStatus"
        );
    var detailsElement =
        document.getElementById(
            "compatibilityDetails"
        );
    var report =
        state.compatibilityReport;

    if (!row || !statusElement || !detailsElement) {
        return;
    }

    row.style.display = "flex";
    statusElement.className = "value status-wait";

    if (!report) {
        statusElement.textContent = "待检查";
        detailsElement.textContent =
            "选择数据表和模板后自动检查。";
        return;
    }

    if (report.status === "compatible") {
        statusElement.textContent = "已通过";
        statusElement.className =
            "value status-success";

        var defaultText =
            report.defaultFields.length > 0
                ? "；保留模板原值：" +
                    report.defaultFields.join("、")
                : "";

        detailsElement.textContent =
            "模板需要 " +
            report.templateFields.length +
            " 个字段，数据映射完整" +
            defaultText +
            "。";
        return;
    }

    if (report.status === "missing-fields") {
        statusElement.textContent = "未通过";
        statusElement.className =
            "value status-error";
        detailsElement.textContent =
            "缺少：" +
            report.missingFields.join("、") +
            "。已阻止生成，避免产生内容不完整的白皮。";
        return;
    }

    statusElement.textContent = "待检查";
    detailsElement.textContent =
        report.status === "template-not-ready"
            ? "请先选择并识别模板。"
            : "请先选择并识别 Excel 数据表。";
}

/**
 * 更新数据 Excel 的识别结果区域。
 */
function updateRecognitionUI() {
    var resultSection =
        document.getElementById(
            "resultSection"
        );

    if (resultSection) {
        resultSection.style.display =
            "block";
    }

    var sheetNameElement =
        document.getElementById(
            "sheetName"
        );

    if (sheetNameElement) {
        sheetNameElement.textContent =
            state.sheetName || "-";
    }

    var dataCountElement =
        document.getElementById(
            "dataCount"
        );

    if (dataCountElement) {
        dataCountElement.textContent =
            String(
                state.dataRows
                    ? state.dataRows.length
                    : 0
            ) +
            " 条";
    }

    var displayFields =
        getCurrentTemplateFieldNames();

    if (displayFields.length === 0) {
        var ordered =
            WhitePaperFields
                .getFillableFieldNames();

        displayFields = ordered.filter(
            function (fieldName) {
                return Boolean(
                    state.fieldMapping &&
                    state.fieldMapping[fieldName]
                );
            }
        );
    }

    var mapHtml = "";

    for (
        var i = 0;
        i < displayFields.length;
        i++
    ) {
        var standardField =
            displayFields[i];

        var sourceFields =
            WhitePaperFields
                .getCompatibleSourceFields(
                    standardField,
                    state.fieldMapping || {}
                );
        var sourceHeaders = [];

        for (
            var sourceIndex = 0;
            sourceIndex < sourceFields.length;
            sourceIndex++
        ) {
            var sourceHeader =
                state.fieldMapping[
                    sourceFields[sourceIndex]
                ];

            if (sourceHeader) {
                sourceHeaders.push(sourceHeader);
            }
        }

        var excelField =
            sourceHeaders.length > 0
                ? sourceHeaders.join(" + ")
                : WhitePaperFields
                    .allowsTemplateDefault(
                        standardField
                    )
                    ? "(保留模板原值)"
                    : "(未匹配)";

        mapHtml +=
            '<div class="field-map-item">';

        mapHtml +=
            '<span class="excel-field">' +
            escapeHtml(
                excelField
            ) +
            "</span>";

        mapHtml +=
            '<span class="arrow">→</span>';

        mapHtml +=
            '<span class="word-field">' +
            escapeHtml(
                standardField
            ) +
            "</span>";

        mapHtml +=
            "</div>";
    }

    var fieldMappingElement =
        document.getElementById(
            "fieldMapping"
        );

    if (fieldMappingElement) {
        fieldMappingElement.innerHTML =
            mapHtml;
    }

    updateCompatibilityUI();
}

/**
 * 检查当前是否已经具备预览和批量生成条件。
 */
function checkCanStart() {
    var canStart =
        Boolean(
            state.excelFile &&
            state.templatePath &&
            state.templateType &&
            state.templateReady &&
            state.compatibilityReady &&
            state.outputPath &&
            state.dataRows.length > 0
        );

    var startButton =
        document.getElementById(
            "btnStart"
        );

    var previewButton =
        document.getElementById(
            "btnPreview"
        );

    if (startButton) {
        startButton.disabled =
            !canStart;
    }

    if (previewButton) {
        previewButton.disabled =
            !canStart;
    }
}

// ==================== 批量生成控制 ====================

/**
 * 预览第一份
 */
function previewFirst() {
    if (!refreshTemplateCompatibility(true)) {
        return;
    }

    if (!ensureTemplateFontConfirmed()) {
        Logger.info(
            "用户取消了非黑体三号模板的预览生成"
        );
        return;
    }

    if (state.dataRows.length === 0) {
        return;
    }

    Logger.info("开始预览第一份数据");

    setElementDisabled("btnPreview", true);
    setElementDisabled("btnStart", true);

    generateSingleDocument(
        state.dataRows[0],
        1,
        true,
        function (result) {
            checkCanStart();

            if (result.success) {
                alert(
                    "预览生成成功！\n" +
                    "文件：" +
                    result.fileName
                );
            } else {
                alert(
                    "预览生成失败：\n" +
                    result.errors.join("\n")
                );
            }
        }
    );
}

/**
 * 开始批量生成
 */
function startBatch() {
    if (!refreshTemplateCompatibility(true)) {
        return;
    }

    if (!ensureTemplateFontConfirmed()) {
        Logger.info(
            "用户取消了非黑体三号模板的批量生成"
        );
        return;
    }

    if (state.isRunning) {
        return;
    }

    if (
        state.dataRows.length === 0
    ) {
        alert("没有有效数据");
        return;
    }

    try {
        FileManager.ensureFolder(
            state.outputPath
        );
    } catch (e) {
        Logger.error(
            "准备输出目录失败",
            e.message
        );

        alert(
            "准备输出目录失败：\n" +
            e.message
        );

        return;
    }

    state.isRunning = true;
    state.cancelRequested = false;
    state.successCount = 0;
    state.failCount = 0;
    state.results = [];

    Logger.clear();

    Logger.info(
        "开始批量生成，共 " +
        state.dataRows.length +
        " 条数据"
    );

    setElementDisabled(
        "btnStart",
        true
    );

    setElementDisabled(
        "btnPreview",
        true
    );

    setElementDisabled(
        "btnCancel",
        false
    );

    setElementDisplay(
        "progressSection",
        "block"
    );

    setElementDisplay(
        "logSection",
        "none"
    );

    setElementDisplay(
        "resultListSection",
        "none"
    );

    setTimeout(
        processBatch,
        50
    );
}

/**
 * 取消批量任务
 */
function cancelBatch() {
    state.cancelRequested = true;

    setElementDisabled(
        "btnCancel",
        true
    );

    Logger.info(
        "用户请求取消任务"
    );
}

/**
 * 批量处理主循环
 */
function processBatch() {
    var total = state.dataRows.length;
    var current = 0;

    function processNext() {
        if (
            state.cancelRequested ||
            current >= total
        ) {
            finishBatch();
            return;
        }

        var rowData = state.dataRows[current];
        var index = current + 1;

        updateProgress(
            current,
            total,
            rowData
        );

        generateSingleDocument(
            rowData,
            index,
            false,
            function (result) {
                if (result.success) {
                    state.successCount++;
                } else {
                    state.failCount++;
                }

                state.results.push(result);
                updateStats();

                current++;

                setTimeout(
                    processNext,
                    30
                );
            }
        );
    }

    processNext();
}

// ==================== 模板坐标记录 ====================

/**
 * 在替换字段前记录模板中各字段原始文字中心
 */
/**
 * 替换前记录模板字段的真实几何位置
 */
function captureTemplateAlignmentTargets(
    doc,
    templateInfo
) {
    var targets = {};

    var fields =
        templateInfo.fields || {};

    var fieldNames =
        getOrderedTemplateFieldNames(
            fields
        );

    for (
        var i = 0;
        i < fieldNames.length;
        i++
    ) {
        var fieldName =
            fieldNames[i];

        var fieldInfo =
            fields[fieldName];

        if (
            !fieldInfo ||
            !fieldInfo.hasUnderline
        ) {
            continue;
        }

        var geometry =
            Alignment.measureFieldGeometry(
                doc,
                fieldInfo
            );

        if (!geometry) {
            Logger.warn(
                "无法记录模板字段位置: " +
                fieldName
            );

            continue;
        }

        targets[fieldName] = {
            centerX:
                geometry.valueCenterX,

            underlineStartX:
                geometry.underlineStartX,

            underlineEndX:
                geometry.underlineEndX
        };

        Logger.info(
            "模板字段位置: " +
            fieldName +
            "，中心=" +
            geometry.valueCenterX.toFixed(3) +
            "，横线起点=" +
            geometry.underlineStartX.toFixed(3) +
            "，横线终点=" +
            geometry.underlineEndX.toFixed(3)
        );
    }

    /*
     * 用户要求：
     * “同步练习类”的中心、
     * 学段、学科值的中心、
     * 印张数字的中心完全上下对齐。
     *
     * 因此以模板“种类”文字中心为唯一基准。
     */
    if (targets["种类"]) {
        var sharedCenter =
            targets["种类"].centerX;

        for (
            var targetName in targets
        ) {
            if (
                !targets.hasOwnProperty(
                    targetName
                )
            ) {
                continue;
            }

            var definition =
                WhitePaperFields
                    .getDefinition(
                        targetName
                    );

            if (
                definition &&
                definition.alignmentGroup ===
                    "legacy-center"
            ) {
                targets[targetName].centerX =
                    sharedCenter;
            }
        }
    }

    return targets;
}

// ==================== 单文档生成核心 ====================

/**
 * 生成单个文档
 */
/**
 * 读取数据行中的标准字段值。
 */
function normalizeDataHeaderName(
    headerName
) {
    return WhitePaperFields.normalize(
        headerName
    );
}

/**
 * 按候选表头名称直接从当前数据行取值。
 * 支持表头中的星号、空格和全角冒号。
 */
function getRowValueByAliases(
    rowData,
    aliases
) {
    var source = rowData || {};
    var candidates = aliases || [];

    for (
        var i = 0;
        i < candidates.length;
        i++
    ) {
        var expected =
            normalizeDataHeaderName(
                candidates[i]
            );

        for (var key in source) {
            if (
                !source.hasOwnProperty(key) ||
                key === "_rowIndex"
            ) {
                continue;
            }

            if (
                normalizeDataHeaderName(
                    key
                ) !== expected
            ) {
                continue;
            }

            var value = source[key];

            if (
                value !== undefined &&
                value !== null &&
                String(value).trim() !== ""
            ) {
                return String(value).trim();
            }
        }
    }

    return "";
}

/**
 * 取得标准字段值。
 * 同时兼容：
 * 1. 单列“使用年级及册次”；
 * 2. 分列“使用年级”与“册次”。
 */
function getMappedFieldValue(
    rowData,
    standardField
) {
    var source = rowData || {};

    function readDirect(fieldName) {
        var header =
            state.fieldMapping &&
            state.fieldMapping[fieldName];

        if (
            header &&
            source.hasOwnProperty(header)
        ) {
            var mappedValue = source[header];

            if (
                mappedValue !== undefined &&
                mappedValue !== null
            ) {
                return String(mappedValue).trim();
            }
        }

        return getRowValueByAliases(
            source,
            getStandardFieldAliases(
                fieldName
            )
        );
    }

    var value = readDirect(
        standardField
    );

    if (
        standardField ===
        "使用年级及册次"
    ) {
        if (!value) {
            var grade = readDirect(
                "使用年级"
            );
            var volume = readDirect(
                "册次"
            );

            value = [grade, volume]
                .filter(function (item) {
                    return Boolean(item);
                })
                .join(" ");
        }
    } else if (
        (
            standardField ===
                "使用年级" ||
            standardField === "册次"
        ) &&
        !value
    ) {
        var combined = readDirect(
            "使用年级及册次"
        );
        var split =
            WhitePaperFields
                .splitGradeAndVolume(
                    combined
                );

        value =
            standardField === "使用年级"
                ? split.grade
                : split.volume;
    }

    if (
        value === undefined ||
        value === null
    ) {
        return "";
    }

    return WhitePaperFields.formatValue(
        standardField,
        value
    );
}

/**
 * 组织 Excel 模板所需字段。
 */
function buildXlsxTemplateValues(
    rowData
) {
    var values = {};
    var templateFields =
        getCurrentTemplateFieldNames();

    for (
        var i = 0;
        i < templateFields.length;
        i++
    ) {
        var fieldName =
            templateFields[i];
        var sourceFields =
            WhitePaperFields
                .getCompatibleSourceFields(
                    fieldName,
                    state.fieldMapping || {}
                );

        if (
            sourceFields.length === 0 &&
            WhitePaperFields
                .allowsTemplateDefault(
                    fieldName
                )
        ) {
            continue;
        }

        values[fieldName] =
            getMappedFieldValue(
                rowData,
                fieldName
            );
    }

    var clearFields =
        WhitePaperFields
            .getClearInOutputFieldNames();
    var currentFields =
        getCurrentTemplateFields();

    for (
        var clearIndex = 0;
        clearIndex < clearFields.length;
        clearIndex++
    ) {
        if (currentFields[clearFields[clearIndex]]) {
            values[clearFields[clearIndex]] = "";
        }
    }

    return values;
}

/**
 * 使用系统文件关联打开生成的 xlsx 文件。
 */
function openGeneratedXlsx(filePath) {
    try {
        if (
            typeof wps !== "undefined" &&
            wps &&
            wps.OAAssist &&
            typeof wps.OAAssist
                .ShellExecute === "function"
        ) {
            wps.OAAssist.ShellExecute(
                "explorer.exe",
                '"' + filePath + '"'
            );
            return true;
        }
    } catch (e) {
        Logger.warn(
            "打开生成的 Excel 文件失败: " +
            e.message
        );
    }

    return false;
}

/**
 * 根据 Excel 模板生成单个 xlsx 文件。
 */
function generateSingleXlsxDocument(
    rowData,
    index,
    isPreview,
    callback
) {
    var result = {
        index: index,
        bookName: getBookName(rowData),
        fileName: "",
        outputPath: "",
        success: false,
        errors: [],
        warnings: []
    };

    try {
        var preventOverwriteElement =
            document.getElementById(
                "preventOverwrite"
            );

        var preventOverwrite =
            preventOverwriteElement
                ? Boolean(
                    preventOverwriteElement.checked
                )
                : true;

        var nameInfo =
            FileManager.generateFileName(
                rowData,
                index,
                ".xlsx",
                state.outputPath,
                preventOverwrite
            );

        result.fileName =
            nameInfo.fileName;
        result.outputPath =
            nameInfo.fullPath;

        if (
            !preventOverwrite &&
            FileManager.exists(
                nameInfo.fullPath
            )
        ) {
            if (
                !FileManager.removeFile(
                    nameInfo.fullPath
                )
            ) {
                throw new Error(
                    "无法覆盖已有文件：" +
                    nameInfo.fullPath
                );
            }
        }

        var fieldValues =
            buildXlsxTemplateValues(
                rowData
            );

        XlsxTemplateEngine.generate(
            state.templatePath,
            nameInfo.fullPath,
            fieldValues,

            function (
                generateResult,
                errorMessage
            ) {
                if (!generateResult) {
                    result.errors.push(
                        "Excel 生成失败: " +
                        (errorMessage || "未知错误")
                    );

                    Logger.error(
                        "生成 Excel 白皮失败",
                        result.bookName +
                        ": " +
                        (errorMessage || "未知错误")
                    );
                } else {
                    result.success = true;

                    Logger.info(
                        "生成成功: " +
                        result.fileName
                    );

                    if (isPreview) {
                        var opened =
                            openGeneratedXlsx(
                                nameInfo.fullPath
                            );

                        if (!opened) {
                            result.warnings.push(
                                "文件已生成，但未能自动打开"
                            );
                        }
                    }
                }

                if (
                    typeof callback ===
                    "function"
                ) {
                    callback(result);
                }
            }
        );

    } catch (e) {
        result.errors.push(
            "生成异常: " +
            e.message
        );

        Logger.error(
            "生成 Excel 白皮异常",
            result.bookName +
            ": " +
            e.message
        );

        if (
            typeof callback ===
            "function"
        ) {
            callback(result);
        }
    }
}

/**
 * 根据模板类型生成单份文件。
 * Word 流程同步执行，Excel 流程异步执行。
 */
function generateSingleDocument(
    rowData,
    index,
    isPreview,
    callback
) {
    if (state.templateType === "xlsx") {
        generateSingleXlsxDocument(
            rowData,
            index,
            isPreview,
            callback
        );
        return;
    }

    var result =
        generateSingleWordDocument(
            rowData,
            index,
            isPreview
        );

    if (
        typeof callback ===
        "function"
    ) {
        callback(result);
    }
}

function generateSingleWordDocument(
    rowData,
    index,
    isPreview
) {
    var result = {
        index: index,
        bookName:
            getBookName(
                rowData
            ),
        fileName: "",
        outputPath: "",
        success: false,
        errors: [],
        warnings: []
    };

    var doc = null;

    try {
        var templateExt =
            FileManager
                .getFileExt(
                    state.templatePath
                );

        var preventOverwriteElement =
            document.getElementById(
                "preventOverwrite"
            );

        var preventOverwrite =
            preventOverwriteElement
                ? Boolean(
                    preventOverwriteElement
                        .checked
                )
                : true;

        var nameInfo =
            FileManager
                .generateFileName(
                    rowData,
                    index,
                    templateExt,
                    state.outputPath,
                    preventOverwrite
                );

        result.fileName =
            nameInfo.fileName;

        result.outputPath =
            nameInfo.fullPath;

        var copied =
            FileManager
                .copyTemplate(
                    state.templatePath,
                    nameInfo.fullPath
                );

        if (!copied) {
            result.errors.push(
                "复制模板文件失败"
            );

            return result;
        }

        doc =
            wordApp.Documents.Open(
                nameInfo.fullPath
            );

        /*
         * 记录模板基线。
         */
        var baseline =
            Validator
                .recordBaseline(
                    doc
                );

        /*
         * 替换前分析模板。
         */
        var templateInfo =
            TemplateAnalyzer
                .analyzeTemplate(
                    doc
                );

        /*
         * 替换前记录各字段原始中心。
         */
        var alignmentTargets =
            captureTemplateAlignmentTargets(
                doc,
                templateInfo
            );

        var fieldData = {};
        var runtimeFields = {};

        var replaceFields =
            getOrderedTemplateFieldNames(
                templateInfo.fields || {}
            );

        for (
            var f = 0;
            f < replaceFields.length;
            f++
        ) {
            var fieldName =
                replaceFields[f];

            var sourceFields =
                WhitePaperFields
                    .getCompatibleSourceFields(
                        fieldName,
                        state.fieldMapping || {}
                    );

            if (
                sourceFields.length === 0 &&
                WhitePaperFields
                    .allowsTemplateDefault(
                        fieldName
                    )
            ) {
                Logger.info(
                    "信息表未提供" +
                    fieldName +
                    "，保留模板原值"
                );

                continue;
            }

            if (sourceFields.length === 0) {
                result.errors.push(
                    "数据表缺少模板字段: " +
                    fieldName
                );
                continue;
            }

            var value =
                getMappedFieldValue(
                    rowData,
                    fieldName
                );

            value =
                String(value);

            fieldData[
                fieldName
            ] = value;

            var fieldInfo =
                templateInfo
                    .fields[
                        fieldName
                    ];

            if (!fieldInfo) {
                result.warnings.push(
                    "模板中未找到字段: " +
                    fieldName
                );

                Logger.warn(
                    "模板中未识别到字段: " +
                    fieldName
                );

                continue;
            }

            /*
             * 新版 WordReplacer 返回：
             *
             * {
             *   success: true,
             *   fieldInfo: runtimeFieldInfo
             * }
             */
            var replaceResult =
                WordReplacer
                    .replaceField(
                        doc,
                        fieldInfo,
                        value
                    );

            if (
                !replaceResult ||
                !replaceResult.success
            ) {
                result.errors.push(
                    "字段替换失败: " +
                    fieldName
                );

                continue;
            }

            runtimeFields[
                fieldName
            ] =
                replaceResult
                    .fieldInfo;
        }

        /*
         * 序号/编号只用于识别数据行和必要时生成文件名，白皮成品中
         * 一律清空。空编号字段完全跳过；已有编号仅原位清值，禁止
         * 重建横线或改变横线起止位置。
         */
        var clearOutputFields =
            WhitePaperFields
                .getClearInOutputFieldNames();

        for (
            var clearFieldIndex = 0;
            clearFieldIndex <
                clearOutputFields.length;
            clearFieldIndex++
        ) {
            var clearFieldName =
                clearOutputFields[
                    clearFieldIndex
                ];
            var clearFieldInfo =
                templateInfo.fields[
                    clearFieldName
                ];

            if (!clearFieldInfo) {
                continue;
            }

            var clearResult =
                WordReplacer
                    .clearValueKeepUnderline(
                        doc,
                        clearFieldInfo
                    );

            if (
                !clearResult ||
                !clearResult.success
            ) {
                result.errors.push(
                    "清除字段失败: " +
                    clearFieldName
                );
            } else {
                Logger.info(
                    "已清除成品中的字段: " +
                    clearFieldName
                );
            }
        }

        /*
         * 使用替换器返回的真实字段范围，
         * 不再重新分析并猜测字段位置。
         */
        var autoCenterElement =
            document.getElementById(
                "autoCenter"
            );

        var autoCenter =
            autoCenterElement
                ? Boolean(
                    autoCenterElement
                        .checked
                )
                : true;

        if (autoCenter) {
            doAlignment(
                doc,
                {
                    fields:
                        runtimeFields
                },
                alignmentTargets
            );
        }

        doc.Save();

        /*
         * 即时校验。
         */
        var validateResult =
            Validator
                .validateOutput(
                    doc,
                    baseline,
                    fieldData
                );

        if (
            !validateResult.passed
        ) {
            result.errors =
                result.errors.concat(
                    validateResult
                        .errors || []
                );
        }

        result.warnings =
            result.warnings.concat(
                validateResult
                    .warnings || []
            );

        var verifyAfterSaveElement =
            document.getElementById(
                "verifyAfterSave"
            );

        var verifyAfterSave =
            verifyAfterSaveElement
                ? Boolean(
                    verifyAfterSaveElement
                        .checked
                )
                : true;

        if (verifyAfterSave) {
            doc.Close();
            doc = null;

            var reopenResult =
                Validator
                    .verifyByReopen(
                        nameInfo.fullPath,
                        baseline,
                        fieldData,
                        wordApp
                    );

            if (
                !reopenResult.passed
            ) {
                result.errors =
                    result.errors.concat(
                        reopenResult
                            .errors || []
                    );
            }

            if (
                isPreview &&
                result.errors.length === 0
            ) {
                wordApp.Documents.Open(
                    nameInfo.fullPath
                );

                try {
                    wordApp.Visible =
                        true;
                } catch (
                    visibleError
                ) {}
            }

        } else {
            if (
                isPreview &&
                result.errors.length === 0
            ) {
                try {
                    doc.ActiveWindow.Visible =
                        true;

                    wordApp.Visible =
                        true;
                } catch (
                    visibleError2
                ) {}

            } else {
                doc.Close();
            }

            doc = null;
        }

        result.success =
            result.errors.length === 0;

        if (result.success) {
            Logger.info(
                "生成成功: " +
                result.fileName
            );
        } else {
            Logger.error(
                "生成失败: " +
                result.fileName,
                result.errors.join(
                    "; "
                )
            );
        }

    } catch (e) {
        result.errors.push(
            "生成异常: " +
            e.message
        );

        Logger.error(
            "生成文档异常",
            result.bookName +
            ": " +
            e.message
        );

        try {
            if (doc) {
                doc.Close(false);
                doc = null;
            }
        } catch (
            closeError
        ) {}
    }

    return result;
}

/**
 * 按模板原始坐标对齐字段
 */
/**
 * 同时恢复文字中心和横线长度
 */
function doAlignment(
    doc,
    templateInfo,
    alignmentTargets
) {
    try {
        var fields =
            templateInfo.fields || {};

        var targets =
            alignmentTargets || {};

        var fieldNames =
            getOrderedTemplateFieldNames(
                fields
            );

        for (
            var i = 0;
            i < fieldNames.length;
            i++
        ) {
            var fieldName =
                fieldNames[i];

            if (
                !fields[fieldName] ||
                !fields[fieldName]
                    .hasUnderline ||
                !targets[fieldName]
            ) {
                continue;
            }

            var passed =
                Alignment
                    .fitFieldToTemplate(
                        doc,
                        fields[fieldName],
                        targets[fieldName]
                    );

            if (!passed) {
                Logger.warn(
                    fieldName +
                    " 未达到模板对齐精度"
                );
            }
        }

        Logger.info(
            "字段中心与横线长度校准完成"
        );

    } catch (e) {
        Logger.warn(
            "模板坐标对齐异常: " +
            e.message
        );
    }
}

// ==================== 进度与统计 ====================

function setElementDisabled(
    id,
    disabled
) {
    var element =
        document.getElementById(
            id
        );

    if (element) {
        element.disabled =
            disabled;
    }
}

function setElementDisplay(
    id,
    displayValue
) {
    var element =
        document.getElementById(
            id
        );

    if (element) {
        element.style.display =
            displayValue;
    }
}

function setElementText(
    id,
    text
) {
    var element =
        document.getElementById(
            id
        );

    if (element) {
        element.textContent =
            text;
    }
}

function updateProgress(
    current,
    total,
    rowData
) {
    var percent =
        total > 0
            ? Math.floor(
                (
                    current /
                    total
                ) *
                100
            )
            : 0;

    var progressFill =
        document.getElementById(
            "progressFill"
        );

    if (progressFill) {
        progressFill.style.width =
            percent + "%";
    }

    setElementText(
        "progressPercent",
        percent + "%"
    );

    setElementText(
        "currentBook",
        "正在处理: " +
        getBookName(
            rowData
        )
    );
}

function updateStats() {
    setElementText(
        "successCount",
        state.successCount
    );

    setElementText(
        "failCount",
        state.failCount
    );
}

// ==================== 完成处理 ====================

function finishBatch() {
    state.isRunning = false;

    setElementDisabled(
        "btnCancel",
        true
    );

    setElementDisabled(
        "btnStart",
        false
    );

    setElementDisabled(
        "btnPreview",
        false
    );

    var progressFill =
        document.getElementById(
            "progressFill"
        );

    if (progressFill) {
        progressFill.style.width =
            "100%";
    }

    setElementText(
        "progressPercent",
        "100%"
    );

    setElementText(
        "currentBook",
        state.cancelRequested
            ? "已取消"
            : "生成完成"
    );

    var errors =
        Logger.getErrors();

    if (
        errors.length > 0
    ) {
        setElementDisplay(
            "logSection",
            "block"
        );

        var logHtml = "";

        for (
            var i = 0;
            i < errors.length;
            i++
        ) {
            logHtml +=
                '<div class="error-log-item">[' +
                escapeHtml(
                    errors[i].time
                ) +
                "] " +
                escapeHtml(
                    errors[i].message
                ) +
                "</div>";
        }

        var errorLog =
            document.getElementById(
                "errorLog"
            );

        if (errorLog) {
            errorLog.innerHTML =
                logHtml;
        }
    }

    setElementDisplay(
        "resultListSection",
        "block"
    );

    var listHtml = "";

    for (
        var j = 0;
        j < state.results.length;
        j++
    ) {
        var item =
            state.results[j];

        listHtml +=
            '<div class="result-item ' +
            (
                item.success
                    ? "success"
                    : "fail"
            ) +
            '">';

        listHtml +=
            '<span class="name" title="' +
            escapeHtml(
                item.fileName
            ) +
            '">' +
            escapeHtml(
                item.bookName
            ) +
            "</span>";

        listHtml +=
            "<span>" +
            (
                item.success
                    ? "✓"
                    : "✗"
            ) +
            "</span>";

        listHtml +=
            "</div>";
    }

    var resultList =
        document.getElementById(
            "resultList"
        );

    if (resultList) {
        resultList.innerHTML =
            listHtml;
    }

    Logger.info(
        "批量任务结束。成功: " +
        state.successCount +
        "，失败: " +
        state.failCount
    );
}

// ==================== 辅助函数 ====================

function getBookName(
    rowData
) {
    var name =
        getMappedFieldValue(
            rowData,
            "书名"
        );

    if (name) {
        return String(name);
    }

    var serialNumber =
        rowData["序号"] ||
        rowData["序号*"] ||
        "";

    if (serialNumber) {
        return (
            "序号" +
            serialNumber
        );
    }

    return "未命名";
}

/**
 * 尝试打开输出文件夹，不使用 ActiveXObject
 */
function openOutputFolder() {
    if (!state.outputPath) {
        return;
    }

    try {
        if (
            typeof wps !==
                "undefined" &&
            wps &&
            wps.OAAssist &&
            typeof wps
                .OAAssist
                .ShellExecute ===
                "function"
        ) {
            wps.OAAssist
                .ShellExecute(
                    "explorer.exe",
                    '"' +
                    state.outputPath +
                    '"'
                );

            return;
        }
    } catch (e1) {
        Logger.warn(
            "直接打开输出文件夹失败: " +
            e1.message
        );
    }

    try {
        if (
            navigator.clipboard &&
            typeof navigator
                .clipboard
                .writeText ===
                "function"
        ) {
            navigator.clipboard
                .writeText(
                    state.outputPath
                );

            alert(
                "当前 WPS 版本不能直接打开文件夹。\n" +
                "输出路径已复制：\n" +
                state.outputPath
            );

            return;
        }
    } catch (e2) {}

    alert(
        "输出文件夹：\n" +
        state.outputPath
    );
}
// ==================== 清理 ====================

window.onbeforeunload =
    function () {
        Logger.info(
            "任务窗格关闭，清理任务窗格资源"
        );

        /*
         * wordApp 指向用户当前使用的 WPS 实例，
         * 不能调用 wordApp.Quit()。
         */
        wordApp = null;

        try {
            ExcelReader.quit();
        } catch (e) {
            Logger.warn(
                "释放 Excel 内存数据失败: " +
                e.message
            );
        }
    };
