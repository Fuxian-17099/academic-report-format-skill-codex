# 公式与 MathType

自动模式写入真正 `m:oMath`，支持行内 segments 和独立 equation block。AST文本仅限叶节点数学符号，用 `type` 构造结构：row(items)、fraction(numerator/denominator)、sub/sup(base/script)、subsup(base/sub/sup)、sqrt(body)、root(degree/body)、integral/sum(lower/upper/body)、matrix(rows)。Greek字符直接用Unicode，如 α、β、ω、Σ。矩阵必须矩形。未知结构报错。

也接受 `mathml` XML字符串和可选 `xsl` 本机文件路径，通过 Microsoft Office 的 MML2OMML.XSL 转换，转换后检查根为 oMath。默认探测 `C:/Program Files/Microsoft Office/root/Office16/MML2OMML.XSL`。不随包分发微软XSL；不安装Office也可用不依赖XSL的AST。MathML输入禁止外部实体与网络解析。

不内置任意LaTeX解释器；Codex可把明确的公式转换为AST，复杂LaTeX可用经确认的外部转换器转MathML。必须核对分式、上下标和数学含义，不能把LaTeX源码写入 m:t 伪装成可编辑结构。

MathType模式设置 equation 的 `mode: mathtype`。创建“【在此插入 MathType 公式】”可见位置，保留右侧SEQ和书签。用户只替换公式内容，不删除制表位、编号域或书签。Python不生成MathType专有OLE，也未验证本机插件调用；原有OLE在保守格式化时跳过，需用户Word复核。

编号在公式右侧，bookmark仅包住STYLEREF、点和SEQ，不含括号。正文自行写“式（[[ref:eq-id]]）”。修改所在章节后更新全部域与目录。宽公式和复杂矩阵的行高/居中必须在Word中检查。
