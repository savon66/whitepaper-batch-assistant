# 第三方组件说明

本项目在 `lib/` 下以本地副本形式内置以下第三方库，加载项运行时不需要联网获取：

## SheetJS Community Edition

- 文件：`lib/xlsx.full.min.js`
- 版本：0.20.3
- 许可证：Apache License 2.0
- 主页：<https://sheetjs.com/>
- 源码：<https://github.com/SheetJS/sheetjs>

## JSZip

- 文件：`lib/jszip.min.js`
- 版本：3.10.1
- 许可证：MIT 或 GPLv3（本项目按 MIT 使用）
- 主页：<https://stuk.github.io/jszip/>
- 源码：<https://github.com/Stuk/jszip>

## 开发依赖

`package.json` 中的 `vite`、`wpsjs`、`wps-jsapi-declare` 仅用于开发调试，不随加载项发布，其许可证以各自 npm 包声明为准。

---

本项目自身代码以 [MIT License](LICENSE) 发布。
