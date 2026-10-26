# JSON接口

UTF-8对象：`title`必需；`institution`默认空；`report_type`默认学术报告；`header`默认题目；`metadata`为封面标签到值的对象。省略的个人信息保持空白，不从样本取值。

`blocks`为顺序数组，每项含type：

| type | 必需字段 | 可选 |
|---|---|---|
| heading | level=1/2/3, text（不含编号） | 不允许跳级 |
| paragraph | text 或 segments | text 支持 [[ref:id]]、[[cite:id]] |
| figure | id, path, caption | width_inches，自动限制在页面内并保持比例 |
| table | id, caption, headers, rows | 行长度与列数一致 |
| equation | id, math(AST) 或 mathml | mode=auto/mathtype；后者无需math |

segments为字符串与数学对象混合列表，如 `["定义", {"math":{"type":"sub","base":"f","script":"out"}}, "如上"]`。

`references`为 `[{"id":"ref-a","text":"用户核实的作者和题录"}]`。text不含手工[1]。可靠DOI/BibTeX需由调用方先解析核实，再构造条目。没有可靠文献不得编造。

可选 `markdown` 为正文，支持 #/##/### 标题、空行分段，会追加在blocks后。Markdown图表、列表、代码块、LaTeX、粗体等需由Codex先转换成结构化blocks，基础解析器不承诺完整CommonMark支持。实验表格数据可由CSV等先转换成headers/rows；不自动分析或补齐测量结果。

生成器对重复ID、不存在引用、不支持结构、缺文件、标题手写编号和跳级报错并不输出最终文件。示例图片是为验证生成器重新绘制的流程图，与原报告无关。
