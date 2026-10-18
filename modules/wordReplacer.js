var WordReplacer = (function () {
    /*
     * 用固定数量的不换行空格重建字段两侧留白。
     *
     * 旧实现只替换字段值。长文本替换成短文本后，整条下划线会先
     * 随段落缩短，再依赖模板中数量不确定的空格补偿宽度。WPS 对
     * 普通空格、Tab 和混合 Run 的字符间距处理并不完全一致，因此
     * 会产生文字偏心和横线终点漂移。
     *
     * 新实现给左右两侧各保留一组格式一致的“宽度锚点”，后续
     * Alignment 模块只改变这些锚点的字符间距，不改变字段文字。
     */
    var STABLE_PADDING_COUNT = 12;
    var DEFAULT_PADDING_CHARACTER = " ";
    var EMPTY_TEMPLATE_VALUE_FONT_NAME = "黑体";
    var EMPTY_TEMPLATE_VALUE_FONT_SIZE = 16;
    var CLEAR_VALUE_TOLERANCE = 0.5;
    var CLEAR_VALUE_MAX_ITERATIONS = 6;
    var CLEAR_VALUE_MAX_SPACING = 18;

    function cloneFieldInfo(fieldInfo) {
        var copy = {};

        for (var key in fieldInfo) {
            if (fieldInfo.hasOwnProperty(key)) {
                copy[key] = fieldInfo[key];
            }
        }

        return copy;
    }

    function cloneFormat(format) {
        var copy = {};
        var source = format || {};

        for (var key in source) {
            if (source.hasOwnProperty(key)) {
                copy[key] = source[key];
            }
        }

        return copy;
    }

    /**
     * v1.0 规则：模板横线上原本没有文字时，不能继承空格 Run 的仿宋、
     * 楷体或不确定字号；新填入的字段值统一使用黑体三号（16 磅）。
     * 有原值的模板仍完整保留原值格式。
     */
    function applyEmptyTemplateValueDefaults(
        format,
        fieldInfo
    ) {
        var result = cloneFormat(format);
        var originalValue = String(
            fieldInfo &&
            fieldInfo.originalValue !==
                undefined &&
            fieldInfo.originalValue !== null
                ? fieldInfo.originalValue
                : ""
        ).replace(
            /[\s\u00a0\u3000]/g,
            ""
        );

        if (originalValue) {
            return result;
        }

        result.name =
            EMPTY_TEMPLATE_VALUE_FONT_NAME;
        result.nameFarEast =
            EMPTY_TEMPLATE_VALUE_FONT_NAME;
        result.size =
            EMPTY_TEMPLATE_VALUE_FONT_SIZE;

        return result;
    }

    function repeatText(text, count) {
        var result = "";

        for (var i = 0; i < count; i++) {
            result += text;
        }

        return result;
    }

    function clamp(value, minimum, maximum) {
        return Math.max(
            minimum,
            Math.min(maximum, value)
        );
    }

    function isBlankValue(value) {
        return String(
            value === undefined || value === null
                ? ""
                : value
        ).replace(
            /[\s\u00a0\u1680\u180e\u2000-\u200d\u202f\u205f\u2060\u3000\ufeff]/g,
            ""
        ).length === 0;
    }

    function getRangeBoundaryX(range, atEnd) {
        try {
            var point = range.Duplicate;

            if (atEnd) {
                point.Start = point.End;
            } else {
                point.End = point.Start;
            }

            var position = Number(
                point.Information(5)
            );

            if (
                isFinite(position) &&
                position >= 0
            ) {
                return position;
            }
        } catch (e) {}

        return NaN;
    }

    function getRangeWidth(range) {
        var startX = getRangeBoundaryX(
            range,
            false
        );
        var endX = getRangeBoundaryX(
            range,
            true
        );

        if (
            !isFinite(startX) ||
            !isFinite(endX)
        ) {
            return NaN;
        }

        return Math.max(0, endX - startX);
    }

    function repaginate(doc) {
        try {
            if (
                doc &&
                typeof doc.Repaginate === "function"
            ) {
                doc.Repaginate();
            }
        } catch (e) {}

        try {
            if (
                doc &&
                doc.Application &&
                doc.Application.ActiveDocument &&
                typeof doc.Application
                    .ActiveDocument
                    .Repaginate === "function"
            ) {
                doc.Application
                    .ActiveDocument
                    .Repaginate();
            }
        } catch (e2) {}
    }

    function getSafeSpacing(range) {
        try {
            var spacing = Number(
                range.Font.Spacing
            );

            if (
                isFinite(spacing) &&
                Math.abs(spacing) < 100
            ) {
                return spacing;
            }
        } catch (e) {}

        return 0;
    }

    /**
     * 把已有编号替换为空白后，用空白字符间距补偿原值宽度。
     * 这样字符数量、下划线起止字符和页面上的横线终点均保持不变。
     */
    function restoreClearedRangeWidth(
        doc,
        range,
        targetWidth,
        characterCount
    ) {
        if (
            !isFinite(targetWidth) ||
            targetWidth <= 0 ||
            characterCount <= 0
        ) {
            return;
        }

        for (
            var iteration = 0;
            iteration < CLEAR_VALUE_MAX_ITERATIONS;
            iteration++
        ) {
            repaginate(doc);

            var currentWidth = getRangeWidth(
                range
            );

            if (!isFinite(currentWidth)) {
                return;
            }

            var widthError =
                targetWidth - currentWidth;

            if (
                Math.abs(widthError) <=
                CLEAR_VALUE_TOLERANCE
            ) {
                return;
            }

            var nextSpacing = clamp(
                getSafeSpacing(range) +
                    widthError /
                        characterCount,
                -CLEAR_VALUE_MAX_SPACING,
                CLEAR_VALUE_MAX_SPACING
            );

            try {
                range.Font.Spacing =
                    nextSpacing;
            } catch (e) {
                return;
            }
        }
    }

    function getRangeText(range) {
        try {
            return String(
                range && range.Text
                    ? range.Text
                    : ""
            );
        } catch (e) {
            return "";
        }
    }

    /**
     * 尽量沿用模板原来的空白字符。
     * 原模板使用普通空格时，不再强制改成不换行空格，避免 WPS
     * 对“下划线 + NBSP + 字符间距”采用不同的线条渲染。
     */
    function detectPaddingCharacter(
        leftRange,
        rightRange
    ) {
        var text =
            getRangeText(leftRange) +
            getRangeText(rightRange);

        if (
            text.indexOf(" ") !== -1
        ) {
            return " ";
        }

        if (
            text.indexOf("\u00A0") !== -1
        ) {
            return "\u00A0";
        }

        if (
            text.indexOf("\u3000") !== -1
        ) {
            return "\u3000";
        }

        return DEFAULT_PADDING_CHARACTER;
    }

    function createParagraphRange(
        paragraph,
        startChar,
        endChar
    ) {
        var range =
            paragraph.Range.Duplicate;

        var paragraphStart =
            paragraph.Range.Start;

        range.Start =
            paragraphStart + startChar;

        range.End =
            paragraphStart + endChar;

        return range;
    }

    function saveEdgeFormat(
        range,
        fromEnd,
        fallback
    ) {
        try {
            if (
                range &&
                range.End > range.Start
            ) {
                var probe =
                    range.Duplicate;

                if (fromEnd) {
                    probe.Start =
                        probe.End - 1;
                } else {
                    probe.End =
                        probe.Start + 1;
                }

                return saveRangeFormat(
                    probe
                );
            }
        } catch (e) {}

        return cloneFormat(
            fallback
        );
    }

    function normalizePaddingFormat(
        format,
        underlineFormat
    ) {
        var result =
            cloneFormat(
                format
            );

        var underlineSource =
            underlineFormat || {};

        if (
            underlineSource.underline !==
            undefined
        ) {
            result.underline =
                underlineSource.underline;
        }

        if (
            underlineSource.underlineColor !==
            undefined
        ) {
            result.underlineColor =
                underlineSource
                    .underlineColor;
        }

        /*
         * 宽度锚点必须具有统一、可预测的初始间距。
         * 实际宽度随后由 Alignment 根据模板坐标精确校准。
         */
        result.spacing = 0;

        if (
            !isFinite(
                Number(result.scaling)
            ) ||
            Number(result.scaling) <= 0
        ) {
            delete result.scaling;
        }

        return result;
    }

    /**
     * 模板横线两侧本来就有可调空白时，优先只替换字段文字。
     * 这样原有 Run、空格字符和下划线格式全部保留；Alignment 仅
     * 调整左右空白的字符间距，属于“格式优先、坐标尽量对齐”。
     */
    function canPreserveOriginalUnderline(
        paragraph,
        fieldInfo
    ) {
        if (
            !canUseStableUnderline(
                fieldInfo
            ) ||
            Number(
                fieldInfo.leftPaddingCount
            ) <= 0 ||
            Number(
                fieldInfo.rightPaddingCount
            ) <= 0
        ) {
            return false;
        }

        try {
            var underlineRange =
                createParagraphRange(
                    paragraph,
                    fieldInfo
                        .underlineStartChar,
                    fieldInfo
                        .underlineEndChar
                );

            /*
             * Tab 的 Font.Spacing 行为不稳定，含 Tab 的模板继续使用
             * 固定锚点回退方案；普通空格模板则保留原横线结构。
             */
            return getRangeText(
                underlineRange
            ).indexOf("\t") === -1;
        } catch (e) {
            return false;
        }
    }

    function canUseStableUnderline(
        fieldInfo
    ) {
        return Boolean(
            fieldInfo &&
            fieldInfo.hasUnderline &&
            typeof fieldInfo
                .underlineStartChar ===
                "number" &&
            typeof fieldInfo
                .underlineEndChar ===
                "number" &&
            typeof fieldInfo
                .valueStartChar ===
                "number" &&
            typeof fieldInfo
                .valueEndChar ===
                "number" &&
            fieldInfo.underlineStartChar <=
                fieldInfo.valueStartChar &&
            fieldInfo.valueStartChar <=
                fieldInfo.valueEndChar &&
            fieldInfo.valueEndChar <=
                fieldInfo.underlineEndChar
        );
    }

    function replaceStableUnderline(
        paragraph,
        fieldInfo,
        value
    ) {
        var underlineStart =
            fieldInfo.underlineStartChar;

        var underlineEnd =
            fieldInfo.underlineEndChar;

        var oldUnderlineLength =
            underlineEnd -
            underlineStart;

        var underlineRange =
            createParagraphRange(
                paragraph,
                underlineStart,
                underlineEnd
            );

        var valueRange =
            createParagraphRange(
                paragraph,
                fieldInfo.valueStartChar,
                fieldInfo.valueEndChar
            );

        var leftPaddingRange =
            createParagraphRange(
                paragraph,
                underlineStart,
                fieldInfo.valueStartChar
            );

        var rightPaddingRange =
            createParagraphRange(
                paragraph,
                fieldInfo.valueEndChar,
                underlineEnd
            );

        var underlineFormat =
            saveEdgeFormat(
                underlineRange,
                false,
                {}
            );

        var valueFormat =
            applyEmptyTemplateValueDefaults(
                saveEdgeFormat(
                    valueRange,
                    false,
                    underlineFormat
                ),
                fieldInfo
            );

        var leftFormat =
            normalizePaddingFormat(
                saveEdgeFormat(
                    leftPaddingRange,
                    true,
                    underlineFormat
                ),
                underlineFormat
            );

        var rightFormat =
            normalizePaddingFormat(
                saveEdgeFormat(
                    rightPaddingRange,
                    false,
                    underlineFormat
                ),
                underlineFormat
            );

        var paddingCharacter =
            detectPaddingCharacter(
                leftPaddingRange,
                rightPaddingRange
            );

        var leftPadding =
            repeatText(
                paddingCharacter,
                STABLE_PADDING_COUNT
            );

        var rightPadding =
            repeatText(
                paddingCharacter,
                STABLE_PADDING_COUNT
            );

        var replacementText =
            leftPadding +
            value +
            rightPadding;

        underlineRange.Text =
            replacementText;

        var runtime =
            cloneFieldInfo(
                fieldInfo
            );

        runtime.underlineStartChar =
            underlineStart;

        runtime.valueStartChar =
            underlineStart +
            leftPadding.length;

        runtime.valueEndChar =
            runtime.valueStartChar +
            value.length;

        runtime.underlineEndChar =
            underlineStart +
            replacementText.length;

        runtime.leftPaddingStartChar =
            runtime.underlineStartChar;

        runtime.leftPaddingEndChar =
            runtime.valueStartChar;

        runtime.rightPaddingStartChar =
            runtime.valueEndChar;

        runtime.rightPaddingEndChar =
            runtime.underlineEndChar;

        runtime.leftPaddingCount =
            STABLE_PADDING_COUNT;

        runtime.rightPaddingCount =
            STABLE_PADDING_COUNT;

        runtime.paddingCharacter =
            paddingCharacter;

        runtime.preservedUnderlineStructure =
            false;

        runtime.underlineTotalLength =
            replacementText.length;

        runtime.originalValue =
            value;

        runtime.originalText =
            value;

        runtime.originalLength =
            value.length;

        var lengthDelta =
            replacementText.length -
            oldUnderlineLength;

        if (
            typeof fieldInfo.areaEndChar ===
            "number"
        ) {
            runtime.areaEndChar =
                fieldInfo.areaEndChar +
                lengthDelta;
        }

        if (
            typeof fieldInfo.areaStartChar ===
            "number"
        ) {
            runtime.underlineStartOffset =
                runtime.underlineStartChar -
                fieldInfo.areaStartChar;

            runtime.underlineEndOffset =
                runtime.underlineEndChar -
                fieldInfo.areaStartChar;
        }

        var newLeftRange =
            createParagraphRange(
                paragraph,
                runtime
                    .leftPaddingStartChar,
                runtime
                    .leftPaddingEndChar
            );

        var newValueRange =
            createParagraphRange(
                paragraph,
                runtime.valueStartChar,
                runtime.valueEndChar
            );

        var newRightRange =
            createParagraphRange(
                paragraph,
                runtime
                    .rightPaddingStartChar,
                runtime
                    .rightPaddingEndChar
            );

        restoreRangeFormat(
            newLeftRange,
            leftFormat
        );

        restoreRangeFormat(
            newValueRange,
            valueFormat
        );

        restoreRangeFormat(
            newRightRange,
            rightFormat
        );

        return runtime;
    }

    function replaceValueOnly(
        paragraph,
        fieldInfo,
        value
    ) {
        var base =
            paragraph.Range.Start;

        var start =
            base +
            fieldInfo.valueStartChar;

        var end =
            base +
            fieldInfo.valueEndChar;

        var oldRange =
            paragraph.Range.Duplicate;

        oldRange.Start = start;
        oldRange.End = end;

        var format =
            applyEmptyTemplateValueDefaults(
                saveRangeFormat(oldRange),
                fieldInfo
            );

        var oldLength =
            end - start;

        oldRange.Text = value;

        var newRange =
            paragraph.Range.Duplicate;

        newRange.Start = start;
        newRange.End =
            start + value.length;

        restoreRangeFormat(
            newRange,
            format
        );

        var delta =
            value.length -
            oldLength;

        var runtime =
            cloneFieldInfo(
                fieldInfo
            );

        runtime.valueStartChar =
            fieldInfo.valueStartChar;

        runtime.valueEndChar =
            fieldInfo.valueStartChar +
            value.length;

        runtime.originalValue =
            value;

        runtime.originalText =
            value;

        runtime.originalLength =
            value.length;

        runtime.preservedUnderlineStructure =
            Boolean(
                fieldInfo.hasUnderline
            );

        if (
            typeof fieldInfo.underlineEndChar ===
            "number"
        ) {
            runtime.underlineEndChar =
                fieldInfo.underlineEndChar +
                delta;
        }

        if (
            typeof fieldInfo.areaEndChar ===
            "number"
        ) {
            runtime.areaEndChar =
                fieldInfo.areaEndChar +
                delta;
        }

        runtime.rightPaddingStartChar =
            runtime.valueEndChar;

        if (
            typeof fieldInfo.rightPaddingEndChar ===
            "number"
        ) {
            runtime.rightPaddingEndChar =
                fieldInfo.rightPaddingEndChar +
                delta;
        }

        return runtime;
    }

    function replaceField(doc, fieldInfo, newValue) {
        try {
            var paragraph =
                doc.Paragraphs.Item(
                    fieldInfo.paragraphIndex
                );

            if (!paragraph) {
                throw new Error("找不到字段段落");
            }

            if (
                typeof fieldInfo.valueStartChar !== "number" ||
                typeof fieldInfo.valueEndChar !== "number"
            ) {
                throw new Error(
                    "字段缺少 valueStartChar/valueEndChar"
                );
            }

            var value = String(
                newValue === undefined ||
                newValue === null
                    ? ""
                    : newValue
            );

            var runtime;

            if (
                canPreserveOriginalUnderline(
                    paragraph,
                    fieldInfo
                )
            ) {
                runtime =
                    replaceValueOnly(
                        paragraph,
                        fieldInfo,
                        value
                    );

                Logger.info(
                    "字段保留模板原横线结构: " +
                    fieldInfo.fieldName
                );
            } else if (
                canUseStableUnderline(
                    fieldInfo
                )
            ) {
                runtime =
                    replaceStableUnderline(
                        paragraph,
                        fieldInfo,
                        value
                    );

                Logger.warn(
                    "字段无法原位保留横线，使用固定空白锚点: " +
                    fieldInfo.fieldName
                );
            } else {
                runtime =
                    replaceValueOnly(
                        paragraph,
                        fieldInfo,
                        value
                    );
            }

            Logger.info(
                "字段替换成功: " +
                fieldInfo.fieldName
            );

            return {
                success: true,
                fieldInfo: runtime
            };

        } catch (e) {
            Logger.error(
                "替换字段失败: " +
                (
                    fieldInfo
                        ? fieldInfo.fieldName
                        : ""
                ),
                e.message
            );

            return {
                success: false,
                error: e.message,
                fieldInfo: null
            };
        }
    }

    function saveRangeFormat(range) {
        var fmt = {};

        try {
            fmt.name =
                range.Font.Name;

            fmt.nameFarEast =
                range.Font.NameFarEast;

            fmt.size =
                range.Font.Size;

            fmt.bold =
                range.Font.Bold;

            fmt.italic =
                range.Font.Italic;

            fmt.underline =
                range.Font.Underline;

            fmt.underlineColor =
                range.Font.UnderlineColor;

            fmt.color =
                range.Font.Color;

            fmt.spacing =
                range.Font.Spacing;

            fmt.scaling =
                range.Font.Scaling;

            fmt.position =
                range.Font.Position;

            fmt.strikeThrough =
                range.Font.StrikeThrough;

            fmt.subscript =
                range.Font.Subscript;

            fmt.superscript =
                range.Font.Superscript;

        } catch (e) {}

        return fmt;
    }

    function isTrueFontFlag(value) {
        return (
            value === true ||
            Number(value) === -1 ||
            Number(value) === 1
        );
    }

    function isSafeNumber(
        value,
        minimum,
        maximum
    ) {
        var numeric = Number(value);

        return (
            isFinite(numeric) &&
            numeric >= minimum &&
            numeric <= maximum &&
            numeric !== 9999999
        );
    }

    function restoreRangeFormat(
        range,
        format
    ) {
        if (!range || !format) {
            return;
        }

        try {
            if (format.name !== undefined) {
                range.Font.Name =
                    format.name;
            }

            if (
                format.nameFarEast !==
                undefined
            ) {
                range.Font.NameFarEast =
                    format.nameFarEast;
            }

            if (
                isSafeNumber(
                    format.size,
                    1,
                    300
                )
            ) {
                range.Font.Size =
                    format.size;
            }

            /*
             * 不把 WPS 的 False/混合值重新写回文档。写入 Bold=0、
             * Italic=0 等会新增显式 OOXML 属性，部分 WPS 版本会让
             * 下划线和字形的视觉粗细发生变化。替换文字本身已继承
             * 原 Range 的非加粗状态；仅对模板明确为 True 的属性补写。
             */
            if (isTrueFontFlag(format.bold)) {
                range.Font.Bold =
                    -1;
            }

            if (isTrueFontFlag(format.italic)) {
                range.Font.Italic =
                    -1;
            }

            if (
                isSafeNumber(
                    format.underline,
                    0,
                    100
                )
            ) {
                range.Font.Underline =
                    format.underline;
            }

            if (
                format.underlineColor !==
                undefined
            ) {
                range.Font.UnderlineColor =
                    format.underlineColor;
            }

            if (format.color !== undefined) {
                range.Font.Color =
                    format.color;
            }

            if (
                isSafeNumber(
                    format.spacing,
                    -100,
                    100
                ) &&
                Math.abs(
                    Number(format.spacing)
                ) > 0.001
            ) {
                range.Font.Spacing =
                    format.spacing;
            }

            if (
                isSafeNumber(
                    format.scaling,
                    1,
                    600
                ) &&
                Math.abs(
                    Number(format.scaling) -
                    100
                ) > 0.001
            ) {
                range.Font.Scaling =
                    format.scaling;
            }

            if (
                isSafeNumber(
                    format.position,
                    -100,
                    100
                ) &&
                Math.abs(
                    Number(format.position)
                ) > 0.001
            ) {
                range.Font.Position =
                    format.position;
            }

            if (
                isTrueFontFlag(
                    format.strikeThrough
                )
            ) {
                range.Font.StrikeThrough =
                    -1;
            }

            if (
                isTrueFontFlag(
                    format.subscript
                )
            ) {
                range.Font.Subscript =
                    -1;
            }

            if (
                isTrueFontFlag(
                    format.superscript
                )
            ) {
                range.Font.Superscript =
                    -1;
            }

        } catch (e) {
            Logger.warn(
                "恢复字段格式部分失败: " +
                e.message
            );
        }
    }

    function clearValueKeepUnderline(
        doc,
        fieldInfo
    ) {
        try {
            if (!fieldInfo) {
                throw new Error("字段信息为空");
            }

            /*
             * 编号横线本来就是空的时，绝对不能调用通用替换流程。
             * 通用流程会为零长度值创建左右各 12 个空白锚点，导致
             * 重庆模板的 18 空格横线被重建成 24 空格。
             */
            if (isBlankValue(fieldInfo.originalValue)) {
                Logger.info(
                    "字段原本为空，完整保留模板横线: " +
                    fieldInfo.fieldName
                );

                return {
                    success: true,
                    fieldInfo:
                        cloneFieldInfo(fieldInfo),
                    skipped: true,
                    preservedUnderlineStructure: true
                };
            }

            if (
                typeof fieldInfo.valueStartChar !==
                    "number" ||
                typeof fieldInfo.valueEndChar !==
                    "number" ||
                fieldInfo.valueEndChar <
                    fieldInfo.valueStartChar
            ) {
                throw new Error(
                    "字段缺少有效的值范围"
                );
            }

            var paragraph =
                doc.Paragraphs.Item(
                    fieldInfo.paragraphIndex
                );

            if (!paragraph) {
                throw new Error("找不到字段段落");
            }

            var valueLength =
                fieldInfo.valueEndChar -
                fieldInfo.valueStartChar;
            var valueRange =
                createParagraphRange(
                    paragraph,
                    fieldInfo.valueStartChar,
                    fieldInfo.valueEndChar
                );
            var targetWidth =
                getRangeWidth(valueRange);
            var format =
                saveRangeFormat(valueRange);
            var paddingCharacter =
                detectPaddingCharacter(
                    valueRange,
                    valueRange
                );
            var blankText = repeatText(
                paddingCharacter,
                valueLength
            );

            valueRange.Text = blankText;

            var clearedRange =
                createParagraphRange(
                    paragraph,
                    fieldInfo.valueStartChar,
                    fieldInfo.valueStartChar +
                        blankText.length
                );

            restoreRangeFormat(
                clearedRange,
                format
            );
            restoreClearedRangeWidth(
                doc,
                clearedRange,
                targetWidth,
                blankText.length
            );

            var runtime =
                cloneFieldInfo(fieldInfo);

            runtime.originalValue = "";
            runtime.originalText = "";
            runtime.originalLength = 0;
            runtime.preservedUnderlineStructure =
                true;

            Logger.info(
                "字段值已清空，横线长度保持不变: " +
                fieldInfo.fieldName
            );

            return {
                success: true,
                fieldInfo: runtime,
                skipped: false,
                preservedUnderlineStructure: true
            };
        } catch (e) {
            Logger.error(
                "清除字段失败: " +
                (
                    fieldInfo
                        ? fieldInfo.fieldName
                        : ""
                ),
                e.message
            );

            return {
                success: false,
                error: e.message,
                fieldInfo: null
            };
        }
    }

    return {
        replaceField:
            replaceField,

        saveRangeFormat:
            saveRangeFormat,

        restoreRangeFormat:
            restoreRangeFormat,

        clearValueKeepUnderline:
            clearValueKeepUnderline,

        STABLE_PADDING_COUNT:
            STABLE_PADDING_COUNT,

        _test: {
            applyEmptyTemplateValueDefaults:
                applyEmptyTemplateValueDefaults,
            isBlankValue:
                isBlankValue
        }
    };
})();
