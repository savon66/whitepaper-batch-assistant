/**
 * Excel 数据读取模块
 *
 * 直接在 WPS 任务窗格中通过 FileReader + SheetJS 读取 Excel。
 * 不启动 WPS 表格后台进程，不使用 ActiveXObject，不使用 CreateObject。
 *
 * 读取时优先使用单元格格式化显示文本，避免：
 * 1. 数字 6 显示为 6.0；
 * 2. 日期变成序列号；
 * 3. 前导零丢失；
 * 4. 百分比显示为小数。
 */
var ExcelReader = (function() {
    var workbook = null;

    /**
     * 检查 SheetJS 是否加载成功
     */
    function ensureXlsxLoaded() {
        if (
            typeof XLSX === "undefined" ||
            !XLSX ||
            typeof XLSX.read !== "function"
        ) {
            throw new Error(
                "未加载xlsx.full.min.js，请检查taskpane.html中的脚本引用顺序"
            );
        }
    }

    /**
     * 读取用户选择的 Excel 文件
     *
     * @param {File} file 浏览器文件选择器返回的File对象
     * @param {Function} callback 回调函数 callback(success, errorMessage)
     */
    function openFile(file, callback) {
        try {
            ensureXlsxLoaded();

            if (!file) {
                callback(false, "未取得Excel文件对象");
                return;
            }

            close();

            var reader = new FileReader();

            reader.onload = function(event) {
                try {
                    var arrayBuffer = event.target.result;

                    workbook = XLSX.read(arrayBuffer, {
                        type: "array",

                        // 保留格式化文字
                        cellText: true,

                        // 保留数字格式代码
                        cellNF: true,

                        // 日期保持Excel数字逻辑，显示值从cell.w读取
                        cellDates: false,

                        // 不限制读取行数
                        sheetRows: 0
                    });

                    if (
                        !workbook ||
                        !workbook.SheetNames ||
                        workbook.SheetNames.length === 0
                    ) {
                        throw new Error("Excel文件中没有可读取的工作表");
                    }

                    Logger.info(
                        "Excel文件读取成功，工作表数量: " +
                        workbook.SheetNames.length
                    );

                    callback(true, "");
                } catch (parseError) {
                    workbook = null;

                    Logger.error(
                        "解析Excel文件失败",
                        parseError.message
                    );

                    callback(
                        false,
                        "解析Excel文件失败: " +
                        parseError.message
                    );
                }
            };

            reader.onerror = function() {
                workbook = null;

                var message = "浏览器读取Excel文件失败";

                if (reader.error && reader.error.message) {
                    message += ": " + reader.error.message;
                }

                Logger.error("读取Excel文件失败", message);
                callback(false, message);
            };

            /*
             * 必须读取为ArrayBuffer。
             * 不得使用readAsText，否则二进制XLS/XLSX会损坏。
             */
            reader.readAsArrayBuffer(file);

        } catch (e) {
            workbook = null;

            Logger.error(
                "打开Excel文件失败",
                e.message
            );

            callback(
                false,
                "打开Excel文件失败: " +
                e.message
            );
        }
    }

    /**
     * 获取指定名称的工作表
     */
    function getSheetByName(sheetName) {
        if (
            !workbook ||
            !workbook.Sheets ||
            !workbook.Sheets[sheetName]
        ) {
            return null;
        }

        var sheet = workbook.Sheets[sheetName];

        /*
         * 为了兼容原taskpane.js中的sheet.Name，
         * 在SheetJS工作表对象上补充Name属性。
         */
        sheet.Name = sheetName;

        return sheet;
    }

    /**
     * 取得工作表有效区域
     */
    function getSheetRange(sheet) {
        if (!sheet || !sheet["!ref"]) {
            return null;
        }

        try {
            return XLSX.utils.decode_range(sheet["!ref"]);
        } catch (e) {
            Logger.warn(
                "解析工作表有效区域失败: " +
                e.message
            );

            return null;
        }
    }

    /**
     * 获取单元格对象
     *
     * rowIndex和columnIndex均为0开始的绝对索引。
     */
    function getCell(sheet, rowIndex, columnIndex) {
        if (!sheet) {
            return null;
        }

        var address = XLSX.utils.encode_cell({
            r: rowIndex,
            c: columnIndex
        });

        return sheet[address] || null;
    }

    function padDatePart(value) {
        var number = Number(value);

        return number < 10
            ? "0" + number
            : String(number);
    }

    /**
     * 日期单元格统一读取为 yyyy-mm-dd，避免 Excel 显示文本
     * “8/10/26”在不同区域设置下被理解成不同日期。
     */
    function getCanonicalDateText(cell) {
        if (!cell) {
            return "";
        }

        try {
            if (
                cell.t === "d" &&
                cell.v instanceof Date &&
                !isNaN(cell.v.getTime())
            ) {
                return (
                    cell.v.getFullYear() +
                    "-" +
                    padDatePart(
                        cell.v.getMonth() + 1
                    ) +
                    "-" +
                    padDatePart(
                        cell.v.getDate()
                    )
                );
            }

            if (
                cell.t === "n" &&
                typeof cell.v === "number" &&
                XLSX.SSF &&
                typeof XLSX.SSF.parse_date_code ===
                    "function"
            ) {
                var formatIsDate = Boolean(
                    cell.z &&
                    typeof XLSX.SSF.is_date ===
                        "function" &&
                    XLSX.SSF.is_date(cell.z)
                );
                var displayLooksLikeDate =
                    /^\s*\d{1,4}[\/.\-]\d{1,2}[\/.\-]\d{1,4}\s*$/.test(
                        String(cell.w || "")
                    );

                if (
                    !formatIsDate &&
                    !displayLooksLikeDate
                ) {
                    return "";
                }

                var parts =
                    XLSX.SSF.parse_date_code(
                        cell.v
                    );

                if (
                    parts &&
                    parts.y &&
                    parts.m &&
                    parts.d
                ) {
                    return (
                        parts.y +
                        "-" +
                        padDatePart(parts.m) +
                        "-" +
                        padDatePart(parts.d)
                    );
                }
            }
        } catch (e) {
            Logger.warn(
                "规范化日期单元格失败: " +
                e.message
            );
        }

        return "";
    }

    /**
     * 获取单元格显示文本
     *
     * 优先级：
     * 1. cell.w：Excel格式化后的显示文本；
     * 2. XLSX.utils.format_cell；
     * 3. cell.v：单元格原始值。
     */
    function getCellText(sheet, rowIndex, columnIndex) {
        var cell = getCell(
            sheet,
            rowIndex,
            columnIndex
        );

        if (!cell) {
            return "";
        }

        var canonicalDate =
            getCanonicalDateText(cell);

        if (canonicalDate) {
            return canonicalDate;
        }

        if (
            cell.w !== undefined &&
            cell.w !== null
        ) {
            return String(cell.w).trim();
        }

        try {
            var formatted = XLSX.utils.format_cell(cell);

            if (
                formatted !== undefined &&
                formatted !== null
            ) {
                return String(formatted).trim();
            }
        } catch (formatError) {
            Logger.warn(
                "格式化单元格失败: " +
                formatError.message
            );
        }

        if (
            cell.v !== undefined &&
            cell.v !== null
        ) {
            return String(cell.v).trim();
        }

        return "";
    }

    /**
     * 查找指定坐标所在的合并区域。
     *
     * SheetJS 只在合并区域左上角保存真实值，其余单元格为空。
     * 表头识别仍使用原始单元格，避免顶部横向合并标题被误判为表头；
     * 仅数据行读取时调用本逻辑继承左上角值。
     */
    function findMergedRange(
        sheet,
        rowIndex,
        columnIndex
    ) {
        if (
            !sheet ||
            !sheet["!merges"] ||
            !sheet["!merges"].length
        ) {
            return null;
        }

        var merges = sheet["!merges"];

        for (var i = 0; i < merges.length; i++) {
            var merge = merges[i];

            if (
                rowIndex >= merge.s.r &&
                rowIndex <= merge.e.r &&
                columnIndex >= merge.s.c &&
                columnIndex <= merge.e.c
            ) {
                return merge;
            }
        }

        return null;
    }

    /**
     * 获取数据单元格显示文本。
     * 当前坐标为空且位于合并区域时，继承该区域左上角单元格的值。
     */
    function getDataCellText(
        sheet,
        rowIndex,
        columnIndex
    ) {
        var directText = getCellText(
            sheet,
            rowIndex,
            columnIndex
        );

        if (directText) {
            return directText;
        }

        var merge = findMergedRange(
            sheet,
            rowIndex,
            columnIndex
        );

        if (!merge) {
            return "";
        }

        if (
            merge.s.r === rowIndex &&
            merge.s.c === columnIndex
        ) {
            return directText;
        }

        return getCellText(
            sheet,
            merge.s.r,
            merge.s.c
        );
    }

    /**
     * 规范化表头，用于匹配
     */
    function normalizeHeader(value) {
        if (
            typeof WhitePaperFields !==
                "undefined" &&
            WhitePaperFields &&
            typeof WhitePaperFields
                .resolveHeader === "function"
        ) {
            var canonical =
                WhitePaperFields
                    .resolveHeader(value);

            if (canonical) {
                return canonical;
            }

            return WhitePaperFields
                .normalize(value);
        }

        var normalized = String(value || "")
            .replace(/\*/g, "")
            .replace(/[：:]/g, "")
            .replace(/\s+/g, "")
            .trim();

        if (
            normalized === "书目名称" ||
            normalized === "教辅材料" ||
            normalized === "教辅名称"
        ) {
            return "书名";
        }

        if (normalized === "适用学段") {
            return "学段";
        }

        if (
            normalized === "所配版本" ||
            normalized === "教材版本"
        ) {
            return "配套教材版本";
        }

        return normalized;
    }

    function inspectHeaderRow(
        sheet,
        range,
        rowIndex
    ) {
        var nonEmpty = 0;
        var recognized = {};

        for (
            var col = range.s.c;
            col <= range.e.c;
            col++
        ) {
            var text = getCellText(
                sheet,
                rowIndex,
                col
            );

            if (!text) {
                continue;
            }

            nonEmpty++;

            if (
                typeof WhitePaperFields !==
                    "undefined" &&
                WhitePaperFields &&
                typeof WhitePaperFields
                    .resolveHeader === "function"
            ) {
                var fieldName =
                    WhitePaperFields
                        .resolveHeader(text);

                if (fieldName) {
                    recognized[fieldName] = true;
                }
            }
        }

        var recognizedCount =
            Object.keys(recognized).length;
        var hasIdentity = Boolean(
            recognized["书名"] ||
            recognized["序号"]
        );

        return {
            rowIndex: rowIndex,
            nonEmpty: nonEmpty,
            recognized: recognized,
            recognizedCount: recognizedCount,
            hasIdentity: hasIdentity,
            score:
                recognizedCount * 100 +
                (hasIdentity ? 50 : 0) +
                Math.min(nonEmpty, 40)
        };
    }

    /**
     * 扫描有效区域前20行，以“已知字段数”为主、非空单元格数为辅
     * 选择表头。这样顶部说明文字或横向合并标题不会再因文字较多而
     * 抢占表头行。
     */
    function findBestHeaderRow(
        sheet,
        range
    ) {
        var firstRow = range.s.r;
        var lastCandidateRow = Math.min(
            range.e.r,
            firstRow + 19
        );
        var best = null;

        for (
            var row = firstRow;
            row <= lastCandidateRow;
            row++
        ) {
            var report = inspectHeaderRow(
                sheet,
                range,
                row
            );

            if (
                !best ||
                report.score > best.score
            ) {
                best = report;
            }
        }

        return best
            ? best.rowIndex
            : firstRow;
    }

    /**
     * 获取工作表表头。
     */
    function getSheetHeaders(sheet) {
        var headers = [];
        var range = getSheetRange(sheet);

        if (!range) {
            return headers;
        }

        var bestRow =
            findBestHeaderRow(
                sheet,
                range
            );

        for (
            var column = range.s.c;
            column <= range.e.c;
            column++
        ) {
            var headerText = getCellText(
                sheet,
                bestRow,
                column
            );

            if (headerText) {
                headers.push(headerText);
            }
        }

        return headers;
    }

    function buildHeaderMapping(headers) {
        var headerMap = {};
        var values = headers || [];

        for (var i = 0; i < values.length; i++) {
            headerMap[values[i]] = i;
        }

        if (
            typeof WhitePaperFields !==
                "undefined" &&
            WhitePaperFields &&
            typeof WhitePaperFields
                .buildMapping === "function"
        ) {
            return WhitePaperFields
                .buildMapping(headerMap);
        }

        var fallback = {};

        for (var h = 0; h < values.length; h++) {
            fallback[normalizeHeader(values[h])] =
                values[h];
        }

        return fallback;
    }

    function getExpectedMatchCount(
        mapping,
        expectedFields
    ) {
        var source = mapping || {};
        var expected = expectedFields || [];
        var count = 0;

        for (var i = 0; i < expected.length; i++) {
            var fieldName =
                typeof WhitePaperFields !==
                    "undefined" &&
                WhitePaperFields
                    ? WhitePaperFields
                        .resolveHeader(
                            expected[i]
                        ) || expected[i]
                    : normalizeHeader(
                        expected[i]
                    );

            var compatible = [];

            if (
                typeof WhitePaperFields !==
                    "undefined" &&
                WhitePaperFields &&
                typeof WhitePaperFields
                    .getCompatibleSourceFields ===
                    "function"
            ) {
                compatible =
                    WhitePaperFields
                        .getCompatibleSourceFields(
                            fieldName,
                            source
                        );
            } else if (source[fieldName]) {
                compatible = [fieldName];
            }

            if (compatible.length > 0) {
                count++;
            }
        }

        return count;
    }

    function evaluateSheetCandidate(
        sheetName,
        priorityNames,
        expectedFields
    ) {
        var sheet = getSheetByName(sheetName);
        var headers = getSheetHeaders(sheet);
        var mapping = buildHeaderMapping(headers);
        var fieldNames = Object.keys(mapping);
        var hasIdentity = Boolean(
            mapping["书名"] ||
            mapping["序号"]
        );
        var contentFieldCount = 0;

        for (var i = 0; i < fieldNames.length; i++) {
            if (
                fieldNames[i] !== "书名" &&
                fieldNames[i] !== "序号"
            ) {
                contentFieldCount++;
            }
        }

        var expectedMatchCount =
            getExpectedMatchCount(
                mapping,
                expectedFields
            );
        var priorityIndex = -1;

        for (
            var p = 0;
            p < (priorityNames || []).length;
            p++
        ) {
            if (priorityNames[p] === sheetName) {
                priorityIndex = p;
                break;
            }
        }

        var valid =
            hasIdentity &&
            contentFieldCount > 0;
        var score =
            expectedMatchCount * 1000 +
            fieldNames.length * 100 +
            contentFieldCount * 10 +
            (
                priorityIndex >= 0
                    ? Math.max(
                        1,
                        50 - priorityIndex
                    )
                    : 0
            );

        return {
            sheet: sheet,
            sheetName: sheetName,
            headers: headers,
            mapping: mapping,
            hasIdentity: hasIdentity,
            contentFieldCount:
                contentFieldCount,
            expectedMatchCount:
                expectedMatchCount,
            recognizedFieldCount:
                fieldNames.length,
            priorityIndex: priorityIndex,
            valid: valid,
            score: score
        };
    }

    /**
     * 识别数据工作表。
     *
     * v1.0 不再使用“0.4 相似度即可通过”的模糊规则，也不允许指定
     * 工作表名绕过表头校验。候选表必须同时具有书名/序号和至少一个
     * 内容字段，再按当前模板字段覆盖度、总字段数和工作表名排序。
     */
    function detectSheet(
        priorityNames,
        requiredHeaders
    ) {
        if (
            !workbook ||
            !workbook.SheetNames
        ) {
            return null;
        }

        var best = null;

        for (
            var s = 0;
            s < workbook.SheetNames.length;
            s++
        ) {
            var report =
                evaluateSheetCandidate(
                    workbook.SheetNames[s],
                    priorityNames || [],
                    requiredHeaders || []
                );

            if (
                report.valid &&
                (
                    !best ||
                    report.score > best.score
                )
            ) {
                best = report;
            }
        }

        if (best) {
            best.sheet._whitePaperDetection =
                best;

            Logger.info(
                "找到工作表（结构化字段匹配）: " +
                best.sheetName +
                "，识别字段 " +
                best.recognizedFieldCount +
                " 个，当前模板匹配 " +
                best.expectedMatchCount +
                " 个"
            );

            return best.sheet;
        }

        Logger.error(
            "未找到符合要求的工作表"
        );

        return null;
    }

    /**
     * 获取表头列索引映射
     *
     * 返回值中的列索引为SheetJS实际0基索引。
     */
    function getHeaderColumnMap(sheet) {
        var map = {};
        var range = getSheetRange(sheet);

        if (!range) {
            return map;
        }

        var headerRow =
            findBestHeaderRow(
                sheet,
                range
            );

        for (
            var column = range.s.c;
            column <= range.e.c;
            column++
        ) {
            var headerText = getCellText(
                sheet,
                headerRow,
                column
            );

            if (headerText) {
                map[headerText] = column;
            }
        }

        map._headerRow = headerRow;

        Logger.info(
            "表头位于第" +
            (headerRow + 1) +
            "行，共识别" +
            (
                Object.keys(map).length - 1
            ) +
            "个字段"
        );

        return map;
    }

    /**
     * 按字段名获取行数据
     *
     * 兼容表头带星号的情况，例如“书名*”。
     */
    function getRowField(
        rowData,
        standardName
    ) {
        if (rowData[standardName]) {
            return rowData[standardName];
        }

        if (rowData[standardName + "*"]) {
            return rowData[
                standardName + "*"
            ];
        }

        for (var key in rowData) {
            if (
                key === "_rowIndex" ||
                !rowData.hasOwnProperty(key)
            ) {
                continue;
            }

            if (
                normalizeHeader(key) ===
                normalizeHeader(standardName)
            ) {
                return rowData[key];
            }
        }

        return "";
    }

    /**
     * 读取所有有效数据行
     */
    function readDataRows(
        sheet,
        headerMap
    ) {
        var data = [];
        var range = getSheetRange(sheet);

        if (
            !range ||
            headerMap._headerRow === undefined
        ) {
            return data;
        }

        var startRow =
            headerMap._headerRow + 1;

        var endRow = range.e.r;

        Logger.info(
            "数据范围: 第" +
            (startRow + 1) +
            "行到第" +
            (endRow + 1) +
            "行"
        );

        for (
            var row = startRow;
            row <= endRow;
            row++
        ) {
            var rowData = {};
            var hasContent = false;

            for (var key in headerMap) {
                if (
                    !headerMap.hasOwnProperty(key) ||
                    key === "_headerRow"
                ) {
                    continue;
                }

                var columnIndex =
                    headerMap[key];

                var cellValue =
                    getDataCellText(
                        sheet,
                        row,
                        columnIndex
                    );

                rowData[key] = cellValue;

                if (cellValue) {
                    hasContent = true;
                }
            }

            var bookName =
                getRowField(
                    rowData,
                    "书名"
                );

            var serialNumber =
                getRowField(
                    rowData,
                    "序号"
                );

            var mainFields =
                typeof WhitePaperFields !==
                    "undefined" &&
                WhitePaperFields &&
                typeof WhitePaperFields
                    .getFillableFieldNames ===
                    "function"
                    ? WhitePaperFields
                        .getFillableFieldNames()
                    : [
                        "种类",
                        "学段",
                        "学科",
                        "使用年级及册次",
                        "印张数",
                        "配套教材版本"
                    ];

            var mainFieldsEmpty = true;

            for (
                var f = 0;
                f < mainFields.length;
                f++
            ) {
                if (
                    mainFields[f] === "书名" ||
                    mainFields[f] === "序号"
                ) {
                    continue;
                }

                if (
                    getRowField(
                        rowData,
                        mainFields[f]
                    )
                ) {
                    mainFieldsEmpty = false;
                    break;
                }
            }

            if (
                (bookName || serialNumber) &&
                !mainFieldsEmpty &&
                hasContent
            ) {
                // 使用Excel界面中的实际行号
                rowData._rowIndex = row + 1;

                data.push(rowData);
            }
        }

        Logger.info(
            "读取到有效数据: " +
            data.length +
            "条"
        );

        return data;
    }

    /**
     * 清理当前工作簿对象
     */
    function close() {
        workbook = null;
    }

    /**
     * SheetJS不创建WPS表格进程，
     * 因此这里只需要释放内存中的工作簿。
     */
    function quit() {
        close();
        Logger.info(
            "Excel内存数据已释放"
        );
    }

    return {
        openFile: openFile,
        detectSheet: detectSheet,
        getHeaderColumnMap:
            getHeaderColumnMap,
        readDataRows: readDataRows,
        close: close,
        quit: quit,
        _test: {
            normalizeHeader:
                normalizeHeader,
            getCanonicalDateText:
                getCanonicalDateText,
            buildHeaderMapping:
                buildHeaderMapping,
            inspectHeaderRow:
                inspectHeaderRow,
            findBestHeaderRow:
                findBestHeaderRow
        }
    };
})();
