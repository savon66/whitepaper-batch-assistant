/**
 * WPS 加载项标准入口
 * WPS 会自动创建 index.html，并从该文件加载 main.js
 */

(function loadBusinessScripts() {
    var scripts = [
        "./modules/logger.js",
        "./modules/fieldRegistry.js",
        "./modules/templateAnalyzer.js",
        "./modules/wordReplacer.js",
        "./modules/alignment.js",

        // 其他模块继续写在这里
        // "./modules/excelReader.js",
        // "./modules/documentGenerator.js",

        "./taskpane.js",
        "./app.js"
    ];

    for (var i = 0; i < scripts.length; i++) {
        document.write(
            '<script type="text/javascript" src="' +
            scripts[i] +
            '"><\/script>'
        );
    }
})();

/**
 * 获取当前加载项所在的网址目录
 */
function GetUrlPath() {
    var url = decodeURI(document.location.toString());

    if (url.indexOf("/") !== -1) {
        url = url.substring(0, url.lastIndexOf("/"));
    }

    return url;
}

/**
 * 点击“批量制作白皮”后，打开右侧任务窗格
 */
function OnWhitePaperButtonClick(control) {
    try {
        var taskPaneId = wps.PluginStorage.getItem(
            "whitepaper_taskpane_id"
        );

        var taskPane;

        if (!taskPaneId) {
            // taskpane.html 当前位于项目根目录
            var taskPaneUrl = GetUrlPath() + "/taskpane.html";

            taskPane = wps.CreateTaskPane(
                taskPaneUrl,
                "白皮批量制作助手 v1.0"
            );

            wps.PluginStorage.setItem(
                "whitepaper_taskpane_id",
                taskPane.ID
            );

            // 设置任务窗格宽度
            taskPane.Width = 420;

            // 显示任务窗格
            taskPane.Visible = true;
        } else {
            // 已经创建过任务窗格时，直接重新显示
            taskPane = wps.GetTaskPane(Number(taskPaneId));
            taskPane.Visible = true;
        }

        return true;
    } catch (error) {
        alert(
            "打开白皮批量制作助手失败：\n" +
            error.message
        );

        return false;
    }
}
