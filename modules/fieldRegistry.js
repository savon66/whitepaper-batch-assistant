/**
 * 白皮字段注册表（v1.0）
 *
 * Excel 表头、Word 模板、Excel 模板、字段映射和生成校验统一从这里
 * 读取字段定义，避免不同模块各自维护别名而再次出现“Excel 已识别，
 * Word 模板却识别为 0 个字段”的情况。
 */
var WhitePaperFields = (function () {
    "use strict";

    var DEFINITIONS = [
        {
            name: "书名",
            aliases: [
                "书名",
                "书目名称",
                "教辅材料",
                "教辅名称",
                "图书名称"
            ],
            category: "identity",
            fillable: true
        },
        {
            name: "种类",
            aliases: [
                "种类",
                "类别",
                "类型",
                "教辅类型",
                "材料类型"
            ],
            category: "content",
            fillable: true,
            alignmentGroup: "legacy-center"
        },
        {
            name: "学段",
            aliases: [
                "学段",
                "适用学段",
                "使用学段"
            ],
            category: "content",
            fillable: true,
            allowTemplateDefault: true,
            alignmentGroup: "legacy-center"
        },
        {
            name: "学科",
            aliases: [
                "学科",
                "科目",
                "适用学科"
            ],
            category: "content",
            fillable: true,
            alignmentGroup: "legacy-center"
        },
        {
            name: "使用年级及册次",
            aliases: [
                "使用年级及册次",
                "使用年级和册次",
                "使用年级与册次",
                "年级及册次",
                "年级和册次",
                "年级册次",
                "使用年级册次"
            ],
            category: "content",
            fillable: true
        },
        {
            name: "使用年级",
            aliases: [
                "使用年级",
                "适用年级",
                "年级"
            ],
            category: "content",
            fillable: true
        },
        {
            name: "册次",
            aliases: [
                "册次",
                "分册",
                "上下册"
            ],
            category: "content",
            fillable: true
        },
        {
            name: "印张数",
            aliases: [
                "印张数",
                "印章数",
                "印张",
                "印章"
            ],
            category: "content",
            fillable: true,
            alignmentGroup: "legacy-center"
        },
        {
            name: "开本印张",
            aliases: [
                "开本印张",
                "开本及印张",
                "开本和印张",
                "开本与印张",
                "开本/印张",
                "开本／印张"
            ],
            category: "content",
            fillable: true
        },
        {
            name: "教材印张",
            aliases: [
                "教材印张",
                "教材印张数",
                "所配教材印张",
                "配套教材印张"
            ],
            category: "content",
            fillable: true
        },
        {
            name: "配套教材版本",
            aliases: [
                "配套教材版本",
                "所配版本",
                "教材版本",
                "配套版本"
            ],
            category: "content",
            fillable: true
        },
        {
            name: "日期",
            aliases: [
                "日期",
                "制作日期",
                "送评日期",
                "填报日期",
                "年月日"
            ],
            category: "content",
            fillable: true,
            allowTemplateDefault: true,
            valueType: "date"
        },
        {
            name: "序号",
            aliases: [
                "序号",
                "编号"
            ],
            category: "identity",
            fillable: false,
            clearInOutput: true
        }
    ];

    var DEFINITION_MAP = {};
    var ALIAS_MAP = {};

    function cloneArray(values) {
        return (values || []).slice();
    }

    function stripTrailingAnnotation(value) {
        var text = String(value || "").trim();
        var previous = "";

        while (text !== previous) {
            previous = text;
            text = text.replace(
                /\s*[（(【\[][^）)】\]]*[）)】\]]\s*$/,
                ""
            );
        }

        return text;
    }

    /**
     * 字段名规范化：
     * - 忽略半角/全角/不换行/零宽空格；
     * - 忽略中英文冒号、必填星号；
     * - 忽略字段名末尾的括号说明；
     * - 不删除“开本”“教材”等语义字符，防止不同印张字段串线。
     */
    function normalize(value) {
        return stripTrailingAnnotation(value)
            .replace(/[\u0000-\u001f]/g, "")
            .replace(/[\s\u00a0\u1680\u180e\u2000-\u200d\u202f\u205f\u2060\u3000\ufeff]/g, "")
            .replace(/[：:＊*]/g, "")
            .trim()
            .toLowerCase();
    }

    function initialize() {
        for (var i = 0; i < DEFINITIONS.length; i++) {
            var definition = DEFINITIONS[i];
            DEFINITION_MAP[definition.name] = definition;

            var allAliases = [definition.name].concat(
                definition.aliases || []
            );

            for (var a = 0; a < allAliases.length; a++) {
                var key = normalize(allAliases[a]);

                if (key && !ALIAS_MAP[key]) {
                    ALIAS_MAP[key] = definition.name;
                }
            }
        }
    }

    initialize();

    function resolve(value) {
        var key = normalize(value);

        if (!key) {
            return "";
        }

        return ALIAS_MAP[key] || "";
    }

    /**
     * 对“字段名（必填）”“字段名-请填写”等表头做保守兼容。
     * 仅允许已知字段名出现在开头，且尾部必须是明显说明语，禁止双向
     * 子串匹配，避免“印张”误吞“开本印张/教材印张”。
     */
    function resolveHeader(value) {
        var exact = resolve(value);

        if (exact) {
            return exact;
        }

        var key = normalize(value);
        var allowedSuffixes = [
            "请填写",
            "填写",
            "必填",
            "选填",
            "可选",
            "信息",
            "内容",
            "名称"
        ];

        for (var aliasKey in ALIAS_MAP) {
            if (!ALIAS_MAP.hasOwnProperty(aliasKey)) {
                continue;
            }

            if (key.indexOf(aliasKey) !== 0) {
                continue;
            }

            var suffix = key.substring(aliasKey.length);

            for (var s = 0; s < allowedSuffixes.length; s++) {
                if (suffix === allowedSuffixes[s]) {
                    return ALIAS_MAP[aliasKey];
                }
            }
        }

        return "";
    }

    function findColon(text) {
        var value = String(text || "");

        for (var i = 0; i < value.length; i++) {
            if (
                value.charAt(i) === "：" ||
                value.charAt(i) === ":"
            ) {
                return i;
            }
        }

        return -1;
    }

    /**
     * 识别 Word/Excel 模板中的标签。标签可有前导空格，汉字之间也可
     * 任意插入空格；返回原文坐标，供替换器继续按原版式操作。
     */
    function matchTemplateText(value) {
        var text = String(value || "")
            .replace(/[\r\u0007]+$/g, "");
        var colonPos = findColon(text);

        if (colonPos >= 0) {
            var rawLabel = text.substring(0, colonPos);
            var fieldName = resolve(rawLabel);

            if (fieldName) {
                return {
                    fieldName: fieldName,
                    matchType: "label",
                    rawLabel: rawLabel,
                    normalizedLabel: normalize(rawLabel),
                    colonPos: colonPos,
                    labelEndChar: colonPos + 1,
                    areaStartChar: colonPos + 1,
                    areaEndChar: text.length
                };
            }
        }

        /* 无标签的独立中文日期行，例如“2026年 8 月 10 日”。 */
        var dateMatch = /^(\s*)(\d{2,4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)(\s*)$/.exec(
            text
        );

        if (dateMatch) {
            var start = dateMatch[1].length;

            return {
                fieldName: "日期",
                matchType: "date",
                rawLabel: "",
                normalizedLabel: "日期",
                colonPos: -1,
                labelEndChar: start,
                areaStartChar: start,
                areaEndChar: start + dateMatch[2].length,
                valueStartChar: start,
                valueEndChar: start + dateMatch[2].length
            };
        }

        return null;
    }

    function getDefinition(fieldName) {
        return DEFINITION_MAP[fieldName] || null;
    }

    function getDefinitions() {
        return DEFINITIONS.slice();
    }

    function getFieldNames() {
        var result = [];

        for (var i = 0; i < DEFINITIONS.length; i++) {
            result.push(DEFINITIONS[i].name);
        }

        return result;
    }

    function getFillableFieldNames() {
        var result = [];

        for (var i = 0; i < DEFINITIONS.length; i++) {
            if (DEFINITIONS[i].fillable) {
                result.push(DEFINITIONS[i].name);
            }
        }

        return result;
    }

    function getClearInOutputFieldNames() {
        var result = [];

        for (var i = 0; i < DEFINITIONS.length; i++) {
            if (DEFINITIONS[i].clearInOutput) {
                result.push(DEFINITIONS[i].name);
            }
        }

        return result;
    }

    function shouldClearInOutput(fieldName) {
        var definition = getDefinition(fieldName);

        return Boolean(
            definition &&
            definition.clearInOutput
        );
    }

    function getAliases(fieldName) {
        var definition = getDefinition(fieldName);

        if (!definition) {
            return [fieldName];
        }

        var result = [definition.name].concat(
            definition.aliases || []
        );
        var unique = [];
        var seen = {};

        for (var i = 0; i < result.length; i++) {
            var key = normalize(result[i]);

            if (!key || seen[key]) {
                continue;
            }

            seen[key] = true;
            unique.push(result[i]);
        }

        return unique;
    }

    function getAliasesMap() {
        var result = {};

        for (var i = 0; i < DEFINITIONS.length; i++) {
            result[DEFINITIONS[i].name] =
                getAliases(DEFINITIONS[i].name);
        }

        return result;
    }

    function findHeader(headerMap, fieldName) {
        var source = headerMap || {};

        for (var header in source) {
            if (
                !source.hasOwnProperty(header) ||
                header === "_headerRow"
            ) {
                continue;
            }

            if (resolveHeader(header) === fieldName) {
                return header;
            }
        }

        return "";
    }

    function buildMapping(headerMap) {
        var result = {};
        var source = headerMap || {};

        for (var header in source) {
            if (
                !source.hasOwnProperty(header) ||
                header === "_headerRow"
            ) {
                continue;
            }

            var fieldName = resolveHeader(header);

            if (fieldName && !result[fieldName]) {
                result[fieldName] = header;
            }
        }

        return result;
    }

    function splitGradeAndVolume(value) {
        var text = String(value || "")
            .replace(/[\r\n\t]+/g, " ")
            .replace(/\s+/g, " ")
            .trim();

        if (!text) {
            return {
                grade: "",
                volume: ""
            };
        }

        var patterns = [
            /^(.*?(?:年级|高一|高二|高三))\s*(上册|下册|全一册|必修\s*[^\s]*|选择性必修\s*[^\s]*|选修\s*[^\s]*)$/,
            /^(.*?)(上册|下册|全一册)$/
        ];

        for (var i = 0; i < patterns.length; i++) {
            var match = text.match(patterns[i]);

            if (match) {
                return {
                    grade: String(match[1] || "").trim(),
                    volume: String(match[2] || "").trim()
                };
            }
        }

        var spaceIndex = text.indexOf(" ");

        if (spaceIndex > 0) {
            return {
                grade: text.substring(0, spaceIndex).trim(),
                volume: text.substring(spaceIndex + 1).trim()
            };
        }

        return {
            grade: text,
            volume: ""
        };
    }

    function getCompatibleSourceFields(
        fieldName,
        mapping
    ) {
        var source = mapping || {};
        var result = [];

        if (source[fieldName]) {
            result.push(fieldName);
            return result;
        }

        if (fieldName === "使用年级及册次") {
            if (source["使用年级"]) {
                result.push("使用年级");
            }

            if (source["册次"]) {
                result.push("册次");
            }
        } else if (
            fieldName === "使用年级" ||
            fieldName === "册次"
        ) {
            if (source["使用年级及册次"]) {
                result.push("使用年级及册次");
            }
        }

        return result;
    }

    function parseDateParts(value) {
        if (value instanceof Date && !isNaN(value.getTime())) {
            return {
                year: value.getFullYear(),
                month: value.getMonth() + 1,
                day: value.getDate()
            };
        }

        var text = String(value || "").trim();

        if (!text) {
            return null;
        }

        var chinese = /^(\d{2,4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日$/.exec(
            text
        );

        var year;
        var month;
        var day;

        if (chinese) {
            year = Number(chinese[1]);
            month = Number(chinese[2]);
            day = Number(chinese[3]);
        } else {
            var iso = /^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/.exec(
                text
            );

            if (iso) {
                year = Number(iso[1]);
                month = Number(iso[2]);
                day = Number(iso[3]);
            } else {
                /* Excel 的常见 m/d/yy 显示文本。 */
                var shortDate = /^(\d{1,2})[\/.](\d{1,2})[\/.](\d{2,4})$/.exec(
                    text
                );

                if (!shortDate) {
                    return null;
                }

                month = Number(shortDate[1]);
                day = Number(shortDate[2]);
                year = Number(shortDate[3]);
            }
        }

        if (year < 100) {
            year += year < 70 ? 2000 : 1900;
        }

        var date = new Date(year, month - 1, day);

        if (
            date.getFullYear() !== year ||
            date.getMonth() + 1 !== month ||
            date.getDate() !== day
        ) {
            return null;
        }

        return {
            year: year,
            month: month,
            day: day
        };
    }

    function formatValue(fieldName, value) {
        var text = String(
            value === undefined || value === null
                ? ""
                : value
        ).trim();
        var definition = getDefinition(fieldName);

        if (
            definition &&
            definition.valueType === "date"
        ) {
            var parts = parseDateParts(value);

            if (parts) {
                return (
                    parts.year +
                    "年 " +
                    parts.month +
                    " 月 " +
                    parts.day +
                    " 日"
                );
            }
        }

        return text;
    }

    function allowsTemplateDefault(fieldName) {
        var definition = getDefinition(fieldName);

        return Boolean(
            definition &&
            definition.allowTemplateDefault
        );
    }

    return {
        definitions: getDefinitions,
        getDefinition: getDefinition,
        getFieldNames: getFieldNames,
        getFillableFieldNames: getFillableFieldNames,
        getClearInOutputFieldNames:
            getClearInOutputFieldNames,
        shouldClearInOutput:
            shouldClearInOutput,
        getAliases: getAliases,
        getAliasesMap: getAliasesMap,
        normalize: normalize,
        resolve: resolve,
        resolveHeader: resolveHeader,
        matchTemplateText: matchTemplateText,
        findHeader: findHeader,
        buildMapping: buildMapping,
        splitGradeAndVolume: splitGradeAndVolume,
        getCompatibleSourceFields:
            getCompatibleSourceFields,
        parseDateParts: parseDateParts,
        formatValue: formatValue,
        allowsTemplateDefault:
            allowsTemplateDefault,
        _test: {
            stripTrailingAnnotation:
                stripTrailingAnnotation,
            findColon: findColon,
            cloneArray: cloneArray
        }
    };
})();
