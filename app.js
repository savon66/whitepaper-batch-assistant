/**
 * 白皮批量制作助手 - WPS JS 加载项入口
 * 功能区回调与插件生命周期管理
 *
 * 安装方式：WPS 开发工具 -> 加载项 -> 添加 -> 选择本文件
 */

// 功能区加载回调
function OnAddinLoad(ribbonUI) {
    try {
        console.log("白皮批量制作助手 v1.0 加载成功");
    } catch (e) {
        console.error("插件加载失败: " + e.message);
    }
}
