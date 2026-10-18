/**
 * 模板字段识别模块
 * 自动识别Word模板中的字段标签、占位区域和下划线
 * 兼容中英文冒号、空格变体、多Run拆分等情况
 */
var TemplateAnalyzer = (function() {
    var fieldConfigs =
        typeof WhitePaperFields !== "undefined"
            ? WhitePaperFields.definitions()
            : [];

    function ensureFieldRegistry() {
        if (
            typeof WhitePaperFields === "undefined" ||
            !WhitePaperFields ||
            typeof WhitePaperFields.matchTemplateText !==
                "function"
        ) {
            throw new Error(
                "字段注册表未加载，请检查 modules/fieldRegistry.js"
            );
        }
    }

    function matchParagraphText(text) {
        ensureFieldRegistry();

        return WhitePaperFields.matchTemplateText(
            stripParagraphMark(text)
        );
    }
    
    // 主入口：分析整个模板
    function analyzeTemplate(doc) {
        var result = {
            success: false,
            error: "",
            fields: {},
            baseline: {},
            fonts: [],
            missingFonts: [],
            diagnostics: {
                unrecognizedLabels: [],
                duplicateFields: []
            }
        };
        
        try {
            ensureFieldRegistry();

            // 记录页面基线信息
            result.baseline.pageCount = doc.ComputeStatistics(2); // wdStatisticPages
            result.baseline.paragraphCount = doc.Paragraphs.Count;
            result.baseline.pageWidth = doc.PageSetup.PageWidth;
            result.baseline.pageHeight = doc.PageSetup.PageHeight;
            result.baseline.topMargin = doc.PageSetup.TopMargin;
            result.baseline.bottomMargin = doc.PageSetup.BottomMargin;
            result.baseline.leftMargin = doc.PageSetup.LeftMargin;
            result.baseline.rightMargin = doc.PageSetup.RightMargin;
            result.baseline.orientation = doc.PageSetup.Orientation;
            
            Logger.info("模板基线: " + result.baseline.pageCount + "页, " + 
                        result.baseline.paragraphCount + "段, " +
                        result.baseline.pageWidth + "x" + result.baseline.pageHeight + "磅");
            
            // 收集所有使用的字体
            var fontSet = {};
            collectFontsFromRange(doc.Content, fontSet);
            result.fonts = Object.keys(fontSet);
            Logger.info("模板使用字体: " + result.fonts.join(", "));
            
            // 检查缺失字体
            result.missingFonts = checkMissingFonts(result.fonts);
            if (result.missingFonts.length > 0) {
                Logger.warn("缺失字体: " + result.missingFonts.join(", "));
            }
            
            // 逐段落扫描识别字段
            var paraCount = doc.Paragraphs.Count;
            for (var i = 1; i <= paraCount; i++) {
                var paragraph = doc.Paragraphs.Item(i);
                var paraText = paragraph.Range.Text;
                
                // 跳过空段落
                if (!paraText || paraText === '\r') continue;
                
                var match =
                    matchParagraphText(
                        paraText
                    );

                if (match) {
                    if (!result.fields[match.fieldName]) {
                        var fieldInfo =
                            analyzeFieldPosition(
                                paragraph,
                                match,
                                i
                            );

                        result.fields[match.fieldName] =
                            fieldInfo;

                        Logger.info(
                            "识别到字段: " +
                            match.fieldName +
                            " (第" +
                            i +
                            "段)"
                        );
                    } else {
                        result.diagnostics
                            .duplicateFields
                            .push({
                                fieldName:
                                    match.fieldName,
                                paragraphIndex: i,
                                text:
                                    stripParagraphMark(
                                        paraText
                                    )
                            });
                    }
                } else if (
                    /[：:]/.test(paraText)
                ) {
                    var colonPos = Math.max(
                        paraText.indexOf("："),
                        paraText.indexOf(":")
                    );

                    if (colonPos < 0) {
                        colonPos =
                            paraText.indexOf("：") >= 0
                                ? paraText.indexOf("：")
                                : paraText.indexOf(":");
                    }

                    result.diagnostics
                        .unrecognizedLabels
                        .push({
                            paragraphIndex: i,
                            label:
                                paraText.substring(
                                    0,
                                    colonPos
                                ),
                            text:
                                stripParagraphMark(
                                    paraText
                                )
                        });
                }
            }
            
            Logger.info("模板识别完成，共找到 " + Object.keys(result.fields).length + " 个字段");
            result.success = true;
            
        } catch (e) {
            result.error = e.message;
            Logger.error("分析模板失败", e.message);
        }
        
        return result;
    }
    
    // 分析单个字段的详细位置信息
function isPaddingCharacter(ch) {
    return (
        ch === " " ||
        ch === "\u00A0" ||
        ch === "\u3000" ||
        ch === "\t"
    );
}

function stripParagraphMark(text) {
    var value =
        String(text || "");

    if (
        value.length > 0 &&
        value.charAt(
            value.length - 1
        ) === "\r"
    ) {
        value =
            value.substring(
                0,
                value.length - 1
            );
    }

    return value;
}

function getRangeBoundaryX(
    range,
    atEnd
) {
    try {
        var point =
            range.Duplicate;

        if (atEnd) {
            point.Start =
                point.End;
        } else {
            point.End =
                point.Start;
        }

        var x =
            Number(
                point.Information(5)
            );

        if (
            isFinite(x) &&
            x >= 0
        ) {
            return x;
        }
    } catch (e) {}

    var text =
        String(range.Text || "");

    if (!text) {
        return 0;
    }

    if (!atEnd) {
        return getCharXPosition(
            range,
            0
        );
    }

    var last =
        text.length - 1;

    var lastX =
        getCharXPosition(
            range,
            last
        );

    var width = 0;

    if (text.length >= 2) {
        width =
            lastX -
            getCharXPosition(
                range,
                last - 1
            );
    }

    if (
        !isFinite(width) ||
        width <= 0
    ) {
        var fontSize = 16;

        try {
            fontSize =
                Number(
                    range.Font.Size
                ) || 16;
        } catch (e2) {}

        width =
            /[\u3400-\u9fff]/.test(
                text.charAt(last)
            )
                ? fontSize
                : fontSize * 0.55;
    }

    return lastX + width;
}


    function analyzeFieldPosition(
    paragraph,
        match,
        paraIndex
    ) {
    var fieldName =
        typeof match === "string"
            ? match
            : match.fieldName;

    var paragraphText =
        stripParagraphMark(
            paragraph.Range.Text
        );

    var colonPos =
        match &&
        typeof match.colonPos === "number"
            ? match.colonPos
            : -1;

    if (colonPos < 0) {
        for (
            var i = 0;
            i < paragraphText.length;
            i++
        ) {
            if (
                paragraphText.charAt(i) === "：" ||
                paragraphText.charAt(i) === ":"
            ) {
                colonPos = i;
                break;
            }
        }
    }

    var isDateMatch =
        match &&
        match.matchType === "date";

    var labelEndChar =
        match &&
        typeof match.labelEndChar === "number"
            ? match.labelEndChar
            : Math.max(0, colonPos + 1);

    /*
     * 字段区域从冒号后立即开始。
     * 这里不能跳过前置空格。
     */
    var areaStartChar =
        match &&
        typeof match.areaStartChar === "number"
            ? match.areaStartChar
            : labelEndChar;

    var areaEndChar =
        match &&
        typeof match.areaEndChar === "number"
            ? match.areaEndChar
            : paragraphText.length;

    var areaRange =
        paragraph.Range.Duplicate;

    areaRange.Start =
        paragraph.Range.Start +
        areaStartChar;

    areaRange.End =
        paragraph.Range.Start +
        areaEndChar;

    var underlineInfo =
        analyzeUnderline(
            areaRange
        );

    var underlineStartChar =
        areaStartChar +
        underlineInfo.startOffset;

    var underlineEndChar =
        areaStartChar +
        underlineInfo.endOffset;

    var underlineText =
        paragraphText.substring(
            underlineStartChar,
            underlineEndChar
        );

    var localStart = 0;
    var localEnd =
        underlineText.length;

    while (
        localStart <
            underlineText.length &&
        isPaddingCharacter(
            underlineText.charAt(
                localStart
            )
        )
    ) {
        localStart++;
    }

    while (
        localEnd >
            localStart &&
        isPaddingCharacter(
            underlineText.charAt(
                localEnd - 1
            )
        )
    ) {
        localEnd--;
    }

    var valueStartChar =
        isDateMatch &&
        typeof match.valueStartChar === "number"
            ? match.valueStartChar
            : underlineStartChar +
                localStart;

    var valueEndChar =
        isDateMatch &&
        typeof match.valueEndChar === "number"
            ? match.valueEndChar
            : underlineStartChar +
                localEnd;

    var valueRange =
        paragraph.Range.Duplicate;

    valueRange.Start =
        paragraph.Range.Start +
        valueStartChar;

    valueRange.End =
        paragraph.Range.Start +
        valueEndChar;

    var originalValue =
        paragraphText.substring(
            valueStartChar,
            valueEndChar
        );

    /*
     * 占位区可能只有带下划线的空格，valueRange 此时为零长度。
     * WPS 对零长度 Range 返回的字体信息不稳定，因此字体检查改用
     * 下划线区域中的首个字符作为格式采样点。
     */
    var formatRange =
        valueRange;

    if (
        valueStartChar >= valueEndChar &&
        underlineEndChar > underlineStartChar
    ) {
        formatRange =
            paragraph.Range.Duplicate;

        formatRange.Start =
            paragraph.Range.Start +
            underlineStartChar;

        formatRange.End =
            Math.min(
                paragraph.Range.Start +
                    underlineEndChar,
                formatRange.Start + 1
            );
    } else if (
        valueStartChar >= valueEndChar &&
        areaStartChar > 0
    ) {
        formatRange =
            paragraph.Range.Duplicate;

        formatRange.Start =
            paragraph.Range.Start +
            areaStartChar - 1;

        formatRange.End =
            paragraph.Range.Start +
            areaStartChar;
    }

    var valueStartX =
        getRangeBoundaryX(
            valueRange,
            false
        );

    var valueEndX =
        getRangeBoundaryX(
            valueRange,
            true
        );

    var valueCenterX =
        (
            valueStartX +
            valueEndX
        ) / 2;

    return {
        paragraphIndex:
            paraIndex,

        fieldName:
            fieldName,

        matchType:
            match && match.matchType
                ? match.matchType
                : "label",

        rawLabel:
            match && match.rawLabel
                ? match.rawLabel
                : "",

        normalizedLabel:
            match && match.normalizedLabel
                ? match.normalizedLabel
                : "",

        labelEndChar:
            labelEndChar,

        areaStartChar:
            areaStartChar,

        areaEndChar:
            areaEndChar,

        underlineStartChar:
            underlineStartChar,

        underlineEndChar:
            underlineEndChar,

        underlineStartOffset:
            underlineInfo.startOffset,

        underlineEndOffset:
            underlineInfo.endOffset,

        underlineTotalLength:
            underlineEndChar -
            underlineStartChar,

        hasUnderline:
            underlineInfo.hasUnderline,

        valueStartChar:
            valueStartChar,

        valueEndChar:
            valueEndChar,

        leftPaddingStartChar:
            underlineStartChar,

        leftPaddingEndChar:
            valueStartChar,

        rightPaddingStartChar:
            valueEndChar,

        rightPaddingEndChar:
            underlineEndChar,

        leftPaddingCount:
            valueStartChar -
            underlineStartChar,

        rightPaddingCount:
            underlineEndChar -
            valueEndChar,

        originalValue:
            originalValue,

        originalText:
            originalValue,

        originalLength:
            originalValue.length,

        startX:
            valueStartX,

        endX:
            valueEndX,

        centerX:
            valueCenterX,

        valueCenterX:
            valueCenterX,

        underlineStartX:
            underlineInfo.startX,

        underlineEndX:
            underlineInfo.endX,

        underlineCenterX:
            (
                underlineInfo.startX +
                underlineInfo.endX
            ) / 2,

        format:
            saveRangeFormat(
                formatRange
            )
    };
}
    
    // 分析范围内的下划线区域
    function analyzeUnderline(range) {
    var text =
        String(range.Text || "");

    var result = {
        hasUnderline: false,
        startOffset: 0,
        endOffset: text.length,
        startX: 0,
        endX: 0
    };

    try {
        var start = -1;
        var end = -1;

        for (
            var i = 0;
            i < text.length;
            i++
        ) {
            var charRange =
                range.Duplicate;

            charRange.Start =
                range.Start + i;

            charRange.End =
                range.Start + i + 1;

            if (
                Number(
                    charRange
                        .Font
                        .Underline
                ) !== 0
            ) {
                if (start < 0) {
                    start = i;
                }

                end = i + 1;
            }
        }

        if (
            start >= 0 &&
            end > start
        ) {
            result.hasUnderline =
                true;

            result.startOffset =
                start;

            result.endOffset =
                end;
        }

        var underlineRange =
            range.Duplicate;

        underlineRange.Start =
            range.Start +
            result.startOffset;

        underlineRange.End =
            range.Start +
            result.endOffset;

        result.startX =
            getRangeBoundaryX(
                underlineRange,
                false
            );

        result.endX =
            getRangeBoundaryX(
                underlineRange,
                true
            );

    } catch (e) {
        Logger.warn(
            "分析下划线失败: " +
            e.message
        );
    }

    return result;
}
    
    // 获取指定字符的X坐标（相对于页面，单位：磅）
    function getCharXPosition(range, charOffset) {
        try {
            var charRange = range.Duplicate;
            charRange.Start = range.Start + charOffset;
            charRange.End = range.Start + charOffset + 1;
            
            // wdHorizontalPositionRelativeToPage = 8
            return charRange.Information(5);
        } catch (e) {
            Logger.warn(
                "获取字符水平坐标失败: " +
                e.message
             );

            return -1;
        }
    }
    
    // 保存范围的完整格式信息
    function saveRangeFormat(range) {
        var fmt = {};
        try {
            fmt.name = range.Font.Name;
            fmt.nameFarEast = range.Font.NameFarEast;
            fmt.size = range.Font.Size;
            fmt.bold = range.Font.Bold;
            fmt.italic = range.Font.Italic;
            fmt.underline = range.Font.Underline;
            fmt.underlineColor = range.Font.UnderlineColor;
            fmt.color = range.Font.Color;
            fmt.spacing = range.Font.Spacing;
            fmt.scaling = range.Font.Scaling;
            fmt.position = range.Font.Position;
            fmt.strikeThrough = range.Font.StrikeThrough;
            fmt.subscript = range.Font.Subscript;
            fmt.superscript = range.Font.Superscript;
        } catch(e) {}
        return fmt;
    }
    
    // 从Range中收集所有字体
    function collectFontsFromRange(range, fontSet) {
        try {
            if (range.Font.Name) fontSet[range.Font.Name] = true;
            if (range.Font.NameFarEast) fontSet[range.Font.NameFarEast] = true;
        } catch(e) {}
    }
    
    // 检查系统中缺失的字体
function checkMissingFonts(fonts) {
    var missing = [];

    try {
        if (
            typeof Application ===
                "undefined" ||
            !Application ||
            !Application.FontNames
        ) {
            return missing;
        }

        var fontNames =
            Application.FontNames;

        var installed = {};

        for (
            var i = 1;
            i <= fontNames.Count;
            i++
        ) {
            installed[
                String(
                    fontNames.Item(i)
                )
            ] = true;
        }

        for (
            var f = 0;
            f < fonts.length;
            f++
        ) {
            if (
                !installed[
                    String(fonts[f])
                ]
            ) {
                missing.push(
                    String(fonts[f])
                );
            }
        }

    } catch (e) {
        Logger.warn(
            "字体检查跳过: " +
            e.message
        );
    }

    return missing;
}   
    
    return {
        analyzeTemplate: analyzeTemplate,
        fieldConfigs: fieldConfigs,
        getCharXPosition: getCharXPosition,
        _test: {
            matchParagraphText:
                matchParagraphText,
            stripParagraphMark:
                stripParagraphMark
        }
    };
})();
