# Logo 文件目录

这里保留周报顶部 logo 的本地放置位置。

默认脚本读取：

- `assets/logo/report_logo.png`

这个 logo 是内部专用资产，不要提交到 git，不要上传到外部平台。`.gitignore` 已忽略常见图片文件后缀。

如需替换 logo，保持文件名 `report_logo.png` 不变即可；如需使用其他文件名，请同步修改 `config/excel_layout.json` 的 `logo.path`。
