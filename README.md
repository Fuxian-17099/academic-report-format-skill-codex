# academic-report-format-skill-codex

个人学术报告排版 Codex Skill，正式版本 **1.0.0**。

根据提供的正文、实验数据、图片、公式和参考文献，生成或规范化 Microsoft Word 学术报告。

## 功能

- 生成学术报告、实验报告和科研结题报告。
- 规范化已有 Word 文档。
- 修复 Word 原生多级标题编号。
- 插入可编辑的 OMML 数学公式。
- 管理图片、表格、公式及参考文献的自动编号和交叉引用。
- 生成自动目录、页眉页脚和页码。
- 执行格式校验并报告具体错误。

## 排版规范

- 正文：正文_改样式，宋体12 pt，固定20 pt行距，首行缩进2字符。
- 英文和数字：Times New Roman。
- 三级标题：黑体16、14、12 pt，全部加粗。
- 标题编号：Word原生多级列表，数字为Times New Roman粗体。
- 页面：A4，四边页边距约3 cm。
- 参考文献：独立样式、自动编号及正文交叉引用。

## 安装

1. 下载本仓库或 Releases 中的完整压缩包并解压。
2. 将包含 SKILL.md 的完整文件夹命名为 academic-report-format。
3. 将该文件夹放到：

   ```text
   %USERPROFILE%\.agents\skills\academic-report-format
   ```

4. 确认 SKILL.md 直接位于上述目录内，避免多嵌套一层文件夹。
5. 使用运行脚本的 Python 安装依赖：

   ```powershell
   python -m pip install -r requirements.txt
   ```

6. 重新打开 Codex，然后调用 $academic-report-format。

## 使用示例

在 Codex 中输入：

```text
使用 $academic-report-format，根据附件 Markdown、实验数据、
图片和参考文献生成 Word 实验报告。

题目：……
章节要求：……

使用可编辑公式和自动交叉引用，不补造实验结果。
```

也可以输入：

```text
使用 $academic-report-format，规范化这份 Word 报告，
修复多级标题编号，保留原文实质内容。
```

## 脚本使用

在本仓库根目录运行：

```powershell
python scripts/generate_report.py examples/example-input.json report.docx
python scripts/validate_report.py report.docx
python -m unittest discover -s tests -v
```

结构化输入格式见 examples/example-input.json，
详细规则见 references 文件夹。

## 验证情况

- 23项自动测试通过，包含负向测试。
- 示例通过Microsoft Word实际域更新和保存检查。
- 五页示例经Word导出并逐页检查。

## 使用限制

- 新生成或修改的报告需要在桌面Word中更新域和整个目录。
- MathType模式提供公式位置、编号和引用，由用户通过插件插入公式。
- 不生成MathType专有OLE对象。
- 已有静态引用和未知样式需要确认映射，不能自动猜测。
- 不自动编造实验数据或参考文献。
- 复杂公式、长表格及最终分页仍需检查实际渲染效果。
