/**
 * 文件管理模块
 *
 * 使用 WPS 加载项自带的 Application.FileSystem。
 * 不使用 ActiveXObject、Scripting.FileSystemObject 或 WScript.Shell。
 */
var FileManager = (function() {

    /**
     * 获取 WPS FileSystem 对象
     */
    function getFileSystem() {
        if (
            typeof Application === "undefined" ||
            !Application
        ) {
            throw new Error(
                "无法获取当前 WPS Application 对象"
            );
        }

        if (!Application.FileSystem) {
            throw new Error(
                "当前 WPS 环境未提供 FileSystem 接口"
            );
        }

        return Application.FileSystem;
    }

    /**
     * 判断文件或文件夹是否存在
     */
    function exists(filePath) {
        var fs = getFileSystem();

        try {
            if (typeof fs.Exists === "function") {
                return Boolean(
                    fs.Exists(String(filePath))
                );
            }

            if (typeof fs.existsSync === "function") {
                return Boolean(
                    fs.existsSync(String(filePath))
                );
            }

            throw new Error(
                "FileSystem不存在Exists或existsSync接口"
            );
        } catch (e) {
            Logger.warn(
                "检查路径是否存在失败: " +
                filePath +
                "，" +
                e.message
            );

            return false;
        }
    }

    /**
     * 连接文件夹路径与文件名
     */
    function joinPath(folderPath, fileName) {
        var folder = String(folderPath || "");
        var name = String(fileName || "");

        folder = folder.replace(/[\\\/]+$/, "");
        name = name.replace(/^[\\\/]+/, "");

        return folder + "\\" + name;
    }

    /**
     * 获取文件扩展名，包含点号
     *
     * 示例：
     * C:\\模板.docx → .docx
     */
    function getFileExt(filePath) {
        var path = String(filePath || "");

        var slashIndex = Math.max(
            path.lastIndexOf("\\"),
            path.lastIndexOf("/")
        );

        var dotIndex = path.lastIndexOf(".");

        if (
            dotIndex === -1 ||
            dotIndex < slashIndex
        ) {
            return "";
        }

        return path
            .substring(dotIndex)
            .toLowerCase();
    }

    /**
     * 清理 Windows 文件名中的非法字符
     */
    function sanitizeFileName(fileName) {
        var value = String(fileName || "");

        // 删除 Windows 文件名禁止字符
        value = value.replace(
            /[\\\/:*?"<>|]/g,
            "_"
        );

        // 删除控制字符
        value = value.replace(
            /[\u0000-\u001f]/g,
            ""
        );

        // 合并连续空格
        value = value.replace(
            /\s+/g,
            " "
        );

        value = value.trim();

        // Windows 文件名不能以句点或空格结束
        value = value.replace(
            /[. ]+$/g,
            ""
        );

        // 避免文件名过长
        if (value.length > 150) {
            value = value.substring(0, 150);
        }

        return value;
    }

    /**
     * 获取书名
     */
    function getBookName(rowData, index) {
        var bookName =
            rowData["书名"] ||
            rowData["书名*"] ||
            rowData["书目名称"] ||
            rowData["书目名称*"] ||
            rowData["教辅材料"] ||
            rowData["教辅材料*"] ||
            rowData["教辅名称"] ||
            rowData["教辅名称*"] ||
            "";

        if (bookName) {
            return sanitizeFileName(bookName);
        }

        var serialNumber =
            rowData["序号"] ||
            rowData["序号*"] ||
            "";

        if (serialNumber) {
            return sanitizeFileName(
                "序号" + serialNumber + "_白皮"
            );
        }

        return sanitizeFileName(
            String(index) + "_白皮"
        );
    }

    /**
     * 生成输出文件名和完整路径
     *
     * @param rowData Excel 行数据
     * @param index 当前数据序号
     * @param templateExt 模板扩展名
     * @param outputPath 输出文件夹
     * @param preventOverwrite 是否禁止覆盖
     */
    function generateFileName(
        rowData,
        index,
        templateExt,
        outputPath,
        preventOverwrite
    ) {
        var extension =
            String(templateExt || "");

        if (
            extension &&
            extension.charAt(0) !== "."
        ) {
            extension = "." + extension;
        }

        if (!extension) {
            extension = ".docx";
        }

        var baseName =
            getBookName(rowData, index);

        if (!baseName) {
            baseName =
                String(index) + "_白皮";
        }

        var fileName =
            baseName + extension;

        var fullPath =
            joinPath(
                outputPath,
                fileName
            );

        /*
         * 禁止覆盖时：
         * 优先增加Excel序号，
         * 仍重复时继续增加数字。
         */
        if (
            preventOverwrite &&
            exists(fullPath)
        ) {
            var serialNumber =
                rowData["序号"] ||
                rowData["序号*"] ||
                index;

            var suffix =
                "_" +
                sanitizeFileName(
                    String(serialNumber)
                );

            fileName =
                baseName +
                suffix +
                extension;

            fullPath =
                joinPath(
                    outputPath,
                    fileName
                );

            var repeatIndex = 2;

            while (exists(fullPath)) {
                fileName =
                    baseName +
                    suffix +
                    "_" +
                    repeatIndex +
                    extension;

                fullPath =
                    joinPath(
                        outputPath,
                        fileName
                    );

                repeatIndex++;

                if (repeatIndex > 9999) {
                    throw new Error(
                        "无法生成不重复的输出文件名"
                    );
                }
            }
        }

        return {
            fileName: fileName,
            fullPath: fullPath
        };
    }

    /**
     * 确保输出文件夹存在
     */
    function ensureFolder(folderPath) {
        var path = String(folderPath || "");

        if (!path) {
            throw new Error(
                "输出文件夹路径为空"
            );
        }

        if (exists(path)) {
            return true;
        }

        var fs = getFileSystem();

        try {
            if (
                typeof fs.mkdirSync === "function"
            ) {
                fs.mkdirSync(path);
            } else if (
                typeof fs.Mkdir === "function"
            ) {
                fs.Mkdir(path);
            } else {
                throw new Error(
                    "FileSystem没有创建文件夹接口"
                );
            }

            if (!exists(path)) {
                throw new Error(
                    "创建文件夹后仍无法访问该目录"
                );
            }

            Logger.info(
                "已创建输出文件夹: " +
                path
            );

            return true;

        } catch (e) {
            Logger.error(
                "创建输出文件夹失败",
                e.message
            );

            throw new Error(
                "无法创建输出文件夹: " +
                e.message
            );
        }
    }

    /**
     * 删除已有文件
     */
    function removeFile(filePath) {
        if (!exists(filePath)) {
            return true;
        }

        var fs = getFileSystem();

        try {
            if (
                typeof fs.unlinkSync === "function"
            ) {
                fs.unlinkSync(
                    String(filePath)
                );
            } else if (
                typeof fs.Remove === "function"
            ) {
                fs.Remove(
                    String(filePath)
                );
            } else {
                throw new Error(
                    "FileSystem没有删除文件接口"
                );
            }

            return !exists(filePath);

        } catch (e) {
            Logger.error(
                "删除已有文件失败",
                e.message
            );

            return false;
        }
    }

    /**
     * 复制 Word 模板
     *
     * @param sourcePath 模板完整路径
     * @param destinationPath 输出文件完整路径
     */
    function copyTemplate(
        sourcePath,
        destinationPath
    ) {
        try {
            var source =
                String(sourcePath || "");

            var destination =
                String(destinationPath || "");

            if (!source) {
                throw new Error(
                    "模板文件路径为空"
                );
            }

            if (!destination) {
                throw new Error(
                    "输出文件路径为空"
                );
            }

            if (!exists(source)) {
                throw new Error(
                    "模板文件不存在: " +
                    source
                );
            }

            /*
             * generateFileName在禁止覆盖时已生成新名称。
             * 用户取消“禁止覆盖”时，则删除已有文件。
             */
            if (exists(destination)) {
                var removed =
                    removeFile(destination);

                if (!removed) {
                    throw new Error(
                        "无法覆盖已有文件: " +
                        destination
                    );
                }
            }

            var fs = getFileSystem();

            if (
                typeof fs.copyFileSync !==
                "function"
            ) {
                throw new Error(
                    "当前WPS不支持FileSystem.copyFileSync"
                );
            }

            fs.copyFileSync(
                source,
                destination
            );

            if (!exists(destination)) {
                throw new Error(
                    "复制完成后未找到输出文件"
                );
            }

            Logger.info(
                "模板复制成功: " +
                destination
            );

            return true;

        } catch (e) {
            Logger.error(
                "复制模板失败",
                e.message
            );

            return false;
        }
    }

    return {
        exists: exists,
        joinPath: joinPath,
        getFileExt: getFileExt,
        sanitizeFileName: sanitizeFileName,
        generateFileName:
            generateFileName,
        ensureFolder: ensureFolder,
        copyTemplate: copyTemplate,
        removeFile: removeFile
    };
})();