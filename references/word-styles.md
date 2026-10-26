# 样式和编号实现

所有正文必须以 `正文_改` 样式控制，Normal仅作继承基础。样式的 rFonts 明确 eastAsia、ascii、hAnsi、cs，移除同节点主题字体属性。拉丁字体统一 Times New Roman；数学对象内部字体不强制更改。图中文字由图片来源负责。

Heading 1/2/3 链接同一个 `w:num` / `w:abstractNum` 的 0/1/2 级：每级含 start=1、decimal、pStyle、suff=space、lvlText=%1 / %1.%2 / %1.%2.%3。使用与源报告一致的完整9级定义，用户入口只支持前三层；省略lvlRestart，使用原生“遇到上一级后重置”的默认规则。Word实际往返验证表明，这比最小3级加显式lvlRestart更稳定。每级rPr显式TNR、粗体、对应字号。段落样式numPr链接numId/ilvl；标题段落也显式指定同一numId/ilvl，outlineLvl=0/1/2。编号数字不写入标题文本。底层修改后按OOXML序列重排pPr/rPr/sectPr，避免Word导入时修复错误。

参考文献标题单独样式、一级大纲，无 numPr。文献编号通过 `SEQ Reference` 生成而非手写 `[1]`；方括号是静态标点，括号内是真实域。标题编号修复采用新增编号定义，避免破坏无关列表；默认仅识别已经有意义的标题样式。

格式化器另存文件、保留段落内文本节点、域、书签和对象。它不能可靠地从任意视觉排版猜测标题或从静态“图2-1”推断引用语义；需要 style-map 或结构化输入。未知样式报告 warning，校验可继续报告 FAIL，不将其伪装成已规范化。python-docx 对未知扩展的无损性不保证，重要文档应比较源/结果包部件和渲染。
