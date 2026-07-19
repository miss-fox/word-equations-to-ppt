# word-equations-to-ppt

日常 Office 工具：把 Word 里的数学公式（OMML）批量写入 PowerPoint，**保持可编辑**，不转图片。

## 输入

| 参数 | 必需 | 说明 |
|------|------|------|
| `--docx` | ✅ | 源 Word 文件 |
| `--out` | ✅ | 输出 pptx 路径 |
| `--per-page` | 推荐 | 每页几道题，如 `2`、`6` |
| `--pptx` | 可选 | 自定义样式模板；不传则用内置默认模板 |
| `--target` | 可选 | `both`（默认）/ `office` / `wps` |

**最简用法（只需 1 个 Word 文件 + 输出路径）：**

```bash
python3 word_to_ppt.py --docx input.docx --out output.pptx --per-page 2
```

默认 `--target both`，同时生成两份：

| 文件 | 给谁 |
|------|------|
| `output.pptx` | Microsoft PowerPoint |
| `output-wps.pptx` | **WPS 演示** |

页数自动计算：`ceil(Word 公式数 / 每页题数)`

## 示例

```bash
# 每页 2 题
python3 word_to_ppt.py --docx 训练.docx --out 输出.pptx --per-page 2

# 每页 6 题（默认模板标签布局）
python3 word_to_ppt.py --docx 训练.docx --out 输出.pptx --per-page 6

# 自定义模板（仅换样式/页眉，题数仍用 --per-page 控制）
python3 word_to_ppt.py --docx 训练.docx --out 输出.pptx --per-page 4 --pptx my-template.pptx
```

## 统一公式格式

Office 可编辑公式的通用格式是 **OMML**，脚本注入前会统一套用：

| 属性 | 默认值 | 说明 |
|------|--------|------|
| 字体 | Cambria Math | Office 公式标准字体 |
| 字号 | 14pt | `--math-pt 14` 可调 |
| 样式 | plain (`m:sty=p`) | 禁止 script/挤压样式 |
| 分式 | bar 型 | 统一分式线 |

```bash
python3 word_to_ppt.py --docx in.docx --out out.pptx --per-page 6 --math-pt 14
```

- Python 3.10+
- Word 源文件公式必须是**公式编辑器对象**（OMML），不能是图片
- 用 **Microsoft PowerPoint** 打开验收（WPS 对 OMML 注入兼容性差）

## 验收

- 点公式 → 「公式工具 / 设计」
- 公式文本框左对齐
- 页数 = ceil(公式数 / 每页题数)

## 原理

1. 从 docx 解压提取 `m:oMath`
2. 用内置/自定义 pptx 作外壳（主题、母版、页面尺寸）
3. 按 `--per-page` 计算坐标，注入 OMML（`a14:m` 包裹，PowerPoint 可编辑）
