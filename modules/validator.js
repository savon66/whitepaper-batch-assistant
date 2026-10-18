/**
 * 生成前后校验模块
 * 确保输出文件格式与模板一致，内容正确
 */
var Validator = (function() {
    
    /**
     * 生成前记录基线信息
     */
    function recordBaseline(doc) {
        var baseline = {};
        
        try {
            // 页面信息
            baseline.pageCount = doc.ComputeStatistics(2); // wdStatisticPages
            baseline.paragraphCount = doc.Paragraphs.Count;
            baseline.pageWidth = doc.PageSetup.PageWidth;
            baseline.pageHeight = doc.PageSetup.PageHeight;
            baseline.topMargin = doc.PageSetup.TopMargin;
            baseline.bottomMargin = doc.PageSetup.BottomMargin;
            baseline.leftMargin = doc.PageSetup.LeftMargin;
            baseline.rightMargin = doc.PageSetup.RightMargin;
            baseline.orientation = doc.PageSetup.Orientation;
            baseline.paperSize = doc.PageSetup.PaperSize;
            
            // 段落格式摘要（前50段）
            baseline.paragraphFormats = [];
            var checkCount = Math.min(doc.Paragraphs.Count, 50);
            for (var i = 1; i <= checkCount; i++) {
                var para = doc.Paragraphs.Item(i);
                baseline.paragraphFormats.push({
                    alignment: para.Alignment,
                    lineSpacing: para.LineSpacing,
                    lineSpacingRule: para.LineSpacingRule,
                    spaceBefore: para.SpaceBefore,
                    spaceAfter: para.SpaceAfter,
                    firstLineIndent: para.FirstLineIndent,
                    leftIndent: para.LeftIndent,
                    rightIndent: para.RightIndent,
                    outlineLevel: para.OutlineLevel
                });
            }
            
            // 节信息
            baseline.sectionCount = doc.Sections.Count;
            
            // 页眉页脚存在性
            try {
                baseline.hasHeader = doc.Sections.Item(1).Headers.Item(1).Range.Text.length > 2;
                baseline.hasFooter = doc.Sections.Item(1).Footers.Item(1).Range.Text.length > 2;
            } catch(e) {}
            
        } catch (e) {
            Logger.warn("记录基线失败: " + e.message);
        }
        
        return baseline;
    }
    
    /**
     * 生成后即时校验
     */
    function validateOutput(doc, baseline, fieldData) {
        var errors = [];
        var warnings = [];
        
        try {
            // 1. 页数校验
            var currentPages = doc.ComputeStatistics(2);
            if (currentPages !== baseline.pageCount) {
                errors.push("页数变化：模板" + baseline.pageCount + "页，输出" + currentPages + "页");
            }
            
            // 2. 页面尺寸校验（允许0.1磅误差）
            if (Math.abs(doc.PageSetup.PageWidth - baseline.pageWidth) > 0.1) {
                errors.push("页面宽度变化");
            }
            if (Math.abs(doc.PageSetup.PageHeight - baseline.pageHeight) > 0.1) {
                errors.push("页面高度变化");
            }
            
            // 3. 页边距校验
            if (Math.abs(doc.PageSetup.TopMargin - baseline.topMargin) > 0.1) {
                errors.push("上边距变化");
            }
            if (Math.abs(doc.PageSetup.BottomMargin - baseline.bottomMargin) > 0.1) {
                errors.push("下边距变化");
            }
            if (Math.abs(doc.PageSetup.LeftMargin - baseline.leftMargin) > 0.1) {
                errors.push("左边距变化");
            }
            if (Math.abs(doc.PageSetup.RightMargin - baseline.rightMargin) > 0.1) {
                errors.push("右边距变化");
            }
            
            // 4. 段落数量校验
            if (doc.Paragraphs.Count !== baseline.paragraphCount) {
                errors.push("段落数量变化：模板" + baseline.paragraphCount + "段，输出" + doc.Paragraphs.Count + "段");
            }
            
            // 5. 节数量校验
            if (doc.Sections.Count !== baseline.sectionCount) {
                errors.push("节数量变化");
            }
            
            // 6. 字段内容校验
            if (fieldData) {
                var fullText = doc.Content.Text;
                for (var fieldName in fieldData) {
                    var value = fieldData[fieldName];
                    if (value && value.length > 0) {
                        if (fullText.indexOf(value) === -1) {
                            errors.push("字段内容缺失: " + fieldName + " = " + value);
                        }
                    }
                }
            }
            
            // 7. 段落格式抽样校验
            var checkCount = Math.min(doc.Paragraphs.Count, baseline.paragraphFormats.length, 30);
            var formatErrors = 0;
            for (var i = 1; i <= checkCount; i++) {
                var para = doc.Paragraphs.Item(i);
                var baseFmt = baseline.paragraphFormats[i - 1];
                if (!baseFmt) continue;
                
                if (para.Alignment !== baseFmt.alignment) formatErrors++;
                if (Math.abs(para.LineSpacing - baseFmt.lineSpacing) > 0.1) formatErrors++;
                if (Math.abs(para.SpaceBefore - baseFmt.spaceBefore) > 0.1) formatErrors++;
                if (Math.abs(para.SpaceAfter - baseFmt.spaceAfter) > 0.1) formatErrors++;
            }
            
            if (formatErrors > 3) {
                warnings.push("段落格式有 " + formatErrors + " 处差异");
            }
            
            // 8. 文件完整性
            if (doc.Paragraphs.Count < 1) {
                errors.push("文档为空或损坏");
            }
            
        } catch (e) {
            errors.push("校验异常: " + e.message);
        }
        
        return {
            passed: errors.length === 0,
            errors: errors,
            warnings: warnings
        };
    }
    
    /**
     * 保存后重新打开校验（二次验证）
     */
    function verifyByReopen(filePath, baseline, fieldData, wordApp) {
        try {
            // 只读方式重新打开
            var doc = wordApp.Documents.Open(filePath, false, true);
            var result = validateOutput(doc, baseline, fieldData);
            doc.Close(false);
            return result;
        } catch (e) {
            return {
                passed: false,
                errors: ["重新打开校验失败: " + e.message],
                warnings: []
            };
        }
    }
    
    /**
     * 校验字体一致性
     */
    function validateFonts(doc, expectedFonts) {
        var errors = [];
        try {
            var actualFonts = {};
            for (var i = 1; i <= Math.min(doc.Paragraphs.Count, 30); i++) {
                var f = doc.Paragraphs.Item(i).Range.Font;
                if (f.Name) actualFonts[f.Name] = true;
                if (f.NameFarEast) actualFonts[f.NameFarEast] = true;
            }
            
            // 检查关键字体是否还在
            for (var e = 0; e < expectedFonts.length; e++) {
                if (!actualFonts[expectedFonts[e]]) {
                    // 可能是字体名略有不同，做模糊检查
                    var found = false;
                    for (var a in actualFonts) {
                        if (a.indexOf(expectedFonts[e]) !== -1 || expectedFonts[e].indexOf(a) !== -1) {
                            found = true;
                            break;
                        }
                    }
                    if (!found) {
                        errors.push("字体可能变化: " + expectedFonts[e]);
                    }
                }
            }
        } catch(e) {}
        
        return errors;
    }
    
    return {
        recordBaseline: recordBaseline,
        validateOutput: validateOutput,
        verifyByReopen: verifyByReopen,
        validateFonts: validateFonts
    };
})();
