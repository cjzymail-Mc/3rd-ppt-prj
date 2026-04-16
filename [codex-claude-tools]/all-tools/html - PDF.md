## 0. 截图生成高清 PDF（已验证流程）
### 适用场景
当 HTML 视觉效果很好，后续用户需要将 html 转为 PDF 文件时，使用本流程。

### 推荐做法
1. 用 Playwright 按 `section.slide` 逐页截图（元素截图，不手写 clip）。
2. 使用 `deviceScaleFactor=4` 进行高采样。
3. 用 Pillow 合并为单个 PDF。

### 关键参数
1. 中文字体：
- `FONTCONFIG_FILE=/tmp/fontconfig-windows.conf`
- 字体目录包含 `/mnt/c/Windows/Fonts`
2. 高清采样：
- `deviceScaleFactor=4`
- 实测可达 `4944x2880` 单页（适合大屏投屏）

### 常见问题
1. 中文乱码：优先检查字体配置与字体目录挂载。
2. 浏览器启动失败：用 `ldd chrome-headless-shell` 查缺失 `.so`。
3. 图片模糊：仅改 viewport 无效，必须提升 `deviceScaleFactor`。
4. 截图越界：改用 `element.screenshot()`。