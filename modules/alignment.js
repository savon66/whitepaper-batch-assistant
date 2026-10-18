/**
 * 白皮字段精密对齐模块
 *
 * 同时约束：
 * 1. 字段文字中心位置；
 * 2. 下划线起点位置；
 * 3. 下划线终点位置。
 *
 * 字段两侧由 WordReplacer 建立固定数量的宽度锚点；本模块只调整
 * 锚点的字符间距，不拉伸字段文字本身。
 */
var Alignment = (function () {

    /*
     * 格式优先：允许亚磅级位置误差，但限制字符间距补偿幅度。
     * 旧值最高可写入 80 磅间距，WPS 会明显改变下划线观感。
     */
    var TOLERANCE = 0.75;
    var MAX_ITERATIONS = 12;
    var MAX_SPACING_STEP = 4.0;
    var MAX_ABSOLUTE_SPACING = 18.0;

    function clamp(value, minimum, maximum) {
        return Math.max(
            minimum,
            Math.min(maximum, value)
        );
    }

    function getParagraphText(paragraph) {
        var text = String(
            paragraph.Range.Text || ""
        );

        if (
            text.length > 0 &&
            text.charAt(text.length - 1) === "\r"
        ) {
            text = text.substring(
                0,
                text.length - 1
            );
        }

        return text;
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

    function resolveUnderlineStart(
        fieldInfo
    ) {
        if (
            typeof fieldInfo.underlineStartChar ===
            "number"
        ) {
            return fieldInfo.underlineStartChar;
        }

        if (
            typeof fieldInfo.areaStartChar ===
                "number" &&
            typeof fieldInfo.underlineStartOffset ===
                "number"
        ) {
            return (
                fieldInfo.areaStartChar +
                fieldInfo.underlineStartOffset
            );
        }

        return null;
    }

    function resolveUnderlineEnd(
        fieldInfo
    ) {
        if (
            typeof fieldInfo.underlineEndChar ===
            "number"
        ) {
            return fieldInfo.underlineEndChar;
        }

        if (
            typeof fieldInfo.areaStartChar ===
                "number" &&
            typeof fieldInfo.underlineEndOffset ===
                "number"
        ) {
            return (
                fieldInfo.areaStartChar +
                fieldInfo.underlineEndOffset
            );
        }

        return null;
    }

    /**
     * 获取 Range 起点或终点的实际水平坐标
     */
    function getBoundaryX(
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

            var position =
                Number(
                    point.Information(5)
                );

            if (
                isFinite(position) &&
                position >= 0
            ) {
                return position;
            }
        } catch (e) {}

        var text =
            String(range.Text || "");

        if (!text) {
            return 0;
        }

        if (!atEnd) {
            return TemplateAnalyzer
                .getCharXPosition(
                    range,
                    0
                );
        }

        var lastIndex =
            text.length - 1;

        var lastX =
            TemplateAnalyzer
                .getCharXPosition(
                    range,
                    lastIndex
                );

        var width = 0;

        if (text.length >= 2) {
            width =
                lastX -
                TemplateAnalyzer
                    .getCharXPosition(
                        range,
                        lastIndex - 1
                    );
        }

        if (
            !isFinite(width) ||
            width <= 0 ||
            width > 100
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
                    text.charAt(lastIndex)
                )
                    ? fontSize
                    : fontSize * 0.55;
        }

        return lastX + width;
    }

    function getSafeSpacing(range) {
        try {
            var spacing =
                Number(
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
     * WPS 在替换 Range.Text 或改变 Font.Spacing 后可能延迟重排。
     * 若直接读取 Information(5)，拿到的仍可能是上一次布局坐标。
     */
    function repaginate(doc) {
        try {
            if (
                doc &&
                typeof doc.Repaginate ===
                    "function"
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
                    .Repaginate ===
                    "function"
            ) {
                doc.Application
                    .ActiveDocument
                    .Repaginate();
            }
        } catch (e2) {}
    }

    function setRangeSpacing(
        range,
        spacing
    ) {
        var safeSpacing =
            clamp(
                spacing,
                -MAX_ABSOLUTE_SPACING,
                MAX_ABSOLUTE_SPACING
            );

        range.Font.Spacing =
            safeSpacing;
    }

    /**
     * 测量字段当前几何位置
     */
    function measureFieldGeometry(
        doc,
        fieldInfo
    ) {
        if (!fieldInfo) {
            return null;
        }

        repaginate(doc);

        var paragraph =
            doc.Paragraphs.Item(
                fieldInfo.paragraphIndex
            );

        if (!paragraph) {
            return null;
        }

        var underlineStartChar =
            resolveUnderlineStart(
                fieldInfo
            );

        var underlineEndChar =
            resolveUnderlineEnd(
                fieldInfo
            );

        if (
            underlineStartChar === null ||
            underlineEndChar === null ||
            typeof fieldInfo.valueStartChar !==
                "number" ||
            typeof fieldInfo.valueEndChar !==
                "number"
        ) {
            Logger.warn(
                "字段坐标信息不完整: " +
                fieldInfo.fieldName
            );

            return null;
        }

        var paragraphLength =
            getParagraphText(
                paragraph
            ).length;

        underlineStartChar =
            clamp(
                underlineStartChar,
                0,
                paragraphLength
            );

        underlineEndChar =
            clamp(
                underlineEndChar,
                underlineStartChar,
                paragraphLength
            );

        var valueStartChar =
            clamp(
                fieldInfo.valueStartChar,
                underlineStartChar,
                underlineEndChar
            );

        var valueEndChar =
            clamp(
                fieldInfo.valueEndChar,
                valueStartChar,
                underlineEndChar
            );

        var underlineRange =
            createParagraphRange(
                paragraph,
                underlineStartChar,
                underlineEndChar
            );

        var valueRange =
            createParagraphRange(
                paragraph,
                valueStartChar,
                valueEndChar
            );

        var leftPaddingRange =
            createParagraphRange(
                paragraph,
                underlineStartChar,
                valueStartChar
            );

        var rightPaddingRange =
            createParagraphRange(
                paragraph,
                valueEndChar,
                underlineEndChar
            );

        var underlineStartX =
            getBoundaryX(
                underlineRange,
                false
            );

        var underlineEndX =
            getBoundaryX(
                underlineRange,
                true
            );

        var valueStartX =
            getBoundaryX(
                valueRange,
                false
            );

        var valueEndX =
            getBoundaryX(
                valueRange,
                true
            );

        return {
            paragraph: paragraph,

            underlineStartChar:
                underlineStartChar,

            underlineEndChar:
                underlineEndChar,

            valueStartChar:
                valueStartChar,

            valueEndChar:
                valueEndChar,

            underlineRange:
                underlineRange,

            valueRange:
                valueRange,

            leftPaddingRange:
                leftPaddingRange,

            rightPaddingRange:
                rightPaddingRange,

            leftPaddingCount:
                valueStartChar -
                underlineStartChar,

            rightPaddingCount:
                underlineEndChar -
                valueEndChar,

            underlineStartX:
                underlineStartX,

            underlineEndX:
                underlineEndX,

            valueStartX:
                valueStartX,

            valueEndX:
                valueEndX,

            valueCenterX:
                (
                    valueStartX +
                    valueEndX
                ) / 2,

            valueWidth:
                valueEndX -
                valueStartX,

            leftPaddingWidth:
                valueStartX -
                underlineStartX,

            rightPaddingWidth:
                underlineEndX -
                valueEndX
        };
    }

    /**
     * 将字段恢复到模板记录的中心和横线边界
     */
    function fitFieldToTemplate(
        doc,
        fieldInfo,
        target
    ) {
        try {
            if (
                !fieldInfo ||
                !target
            ) {
                return false;
            }

            repaginate(doc);

            var targetCenterX =
                Number(
                    target.centerX
                );

            var targetUnderlineStartX =
                Number(
                    target.underlineStartX
                );

            var targetUnderlineEndX =
                Number(
                    target.underlineEndX
                );

            if (
                !isFinite(targetCenterX) ||
                !isFinite(
                    targetUnderlineStartX
                ) ||
                !isFinite(
                    targetUnderlineEndX
                )
            ) {
                Logger.warn(
                    "模板目标坐标无效: " +
                    fieldInfo.fieldName
                );

                return false;
            }

            for (
                var iteration = 0;
                iteration < MAX_ITERATIONS;
                iteration++
            ) {
                var geometry =
                    measureFieldGeometry(
                        doc,
                        fieldInfo
                    );

                if (!geometry) {
                    return false;
                }

                var centerError =
                    geometry.valueCenterX -
                    targetCenterX;

                var endError =
                    geometry.underlineEndX -
                    targetUnderlineEndX;

                var startError =
                    geometry.underlineStartX -
                    targetUnderlineStartX;

                if (
                    Math.abs(centerError) <=
                        TOLERANCE &&
                    Math.abs(endError) <=
                        TOLERANCE &&
                    Math.abs(startError) <=
                        TOLERANCE
                ) {
                    Logger.info(
                        fieldInfo.fieldName +
                        " 对齐完成，中心误差=" +
                        centerError.toFixed(3) +
                        "，横线终点误差=" +
                        endError.toFixed(3)
                    );

                    return true;
                }

                /*
                 * 目标左侧留白宽度：
                 * 目标中心 - 文字宽度一半 - 横线起点。
                 */
                var desiredLeftWidth =
                    targetCenterX -
                    targetUnderlineStartX -
                    geometry.valueWidth / 2;

                /*
                 * 目标右侧留白宽度：
                 * 横线终点 - 目标中心 - 文字宽度一半。
                 */
                var desiredRightWidth =
                    targetUnderlineEndX -
                    targetCenterX -
                    geometry.valueWidth / 2;

                if (
                    desiredLeftWidth < 0 ||
                    desiredRightWidth < 0
                ) {
                    Logger.warn(
                        fieldInfo.fieldName +
                        " 内容宽度超过模板可用区域"
                    );

                    return false;
                }

                var leftWidthError =
                    desiredLeftWidth -
                    geometry.leftPaddingWidth;

                var rightWidthError =
                    desiredRightWidth -
                    geometry.rightPaddingWidth;

                var changed = false;

                if (
                    geometry.leftPaddingCount > 0 &&
                    Math.abs(leftWidthError) >
                        TOLERANCE
                ) {
                    var oldLeftSpacing =
                        getSafeSpacing(
                            geometry
                                .leftPaddingRange
                        );

                    var leftSpacingDelta =
                        clamp(
                            leftWidthError /
                            geometry
                                .leftPaddingCount,
                            -MAX_SPACING_STEP,
                            MAX_SPACING_STEP
                        );

                    setRangeSpacing(
                        geometry
                            .leftPaddingRange,
                        oldLeftSpacing +
                        leftSpacingDelta
                    );

                    changed = true;
                }

                /*
                 * 右侧字符间距负责把横线终点
                 * 恢复到模板原位置。
                 */
                if (
                    geometry.rightPaddingCount > 0 &&
                    Math.abs(rightWidthError) >
                        TOLERANCE
                ) {
                    var oldRightSpacing =
                        getSafeSpacing(
                            geometry
                                .rightPaddingRange
                        );

                    var rightSpacingDelta =
                        clamp(
                            rightWidthError /
                            geometry
                                .rightPaddingCount,
                            -MAX_SPACING_STEP,
                            MAX_SPACING_STEP
                        );

                    setRangeSpacing(
                        geometry
                            .rightPaddingRange,
                        oldRightSpacing +
                        rightSpacingDelta
                    );

                    changed = true;
                }

                if (!changed) {
                    break;
                }

                /*
                 * 强制 WPS 完成本轮排版后再测量。没有这一步时，
                 * 连续迭代可能一直基于旧坐标，最终过度补偿。
                 */
                repaginate(doc);
            }

            repaginate(doc);

            var finalGeometry =
                measureFieldGeometry(
                    doc,
                    fieldInfo
                );

            if (!finalGeometry) {
                return false;
            }

            var finalCenterError =
                finalGeometry.valueCenterX -
                targetCenterX;

            var finalEndError =
                finalGeometry.underlineEndX -
                targetUnderlineEndX;

            var finalStartError =
                finalGeometry.underlineStartX -
                targetUnderlineStartX;

            var passed =
                Math.abs(finalCenterError) <=
                    TOLERANCE &&
                Math.abs(finalEndError) <=
                    TOLERANCE &&
                Math.abs(finalStartError) <=
                    TOLERANCE;

            Logger.info(
                fieldInfo.fieldName +
                " 最终中心误差=" +
                finalCenterError.toFixed(3) +
                "，横线起点误差=" +
                finalStartError.toFixed(3) +
                "，横线终点误差=" +
                finalEndError.toFixed(3)
            );

            return passed;

        } catch (e) {
            Logger.error(
                "字段模板对齐失败: " +
                fieldInfo.fieldName,
                e.message
            );

            return false;
        }
    }

    return {
        measureFieldGeometry:
            measureFieldGeometry,

        fitFieldToTemplate:
            fitFieldToTemplate,

        TOLERANCE:
            TOLERANCE
    };

})();
