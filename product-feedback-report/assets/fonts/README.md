# 字体文件目录

这里保留内部字体文件的本地放置位置。

需要的字体：

- `方正FW筑紫黑简 R.ttf`
- `Heytea Sans Serif Regular.otf`

这两个字体是内部独家字体，不要提交到 git，不要上传到外部平台。`.gitignore` 已忽略常见字体文件后缀。

默认配置在 `config/font_files.json` 中。字体来源统一由 `fontDirectory` 和各字体的 `fileName` 决定；不要配置本机绝对路径。把字体文件放到本目录后，确认 `fileName` 与实际文件名一致即可。

如果字体文件缺失，脚本会使用 `config/font_files.json` 里的默认字体配置，不会阻断生成。
