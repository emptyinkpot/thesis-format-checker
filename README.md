# thesis-format-checker

DOCX 论文格式合规校验器，支持自定义规则集 (preset) 和自动微调。

## 安装

```bash
pip install -e .
```

依赖 [pandoc](https://pandoc.org/) 用于内容提取（需确保 pandoc 在 PATH 中）。

## 统一入口

这个仓库只需要记两个入口：

```powershell
# 业务入口：基于 Downloads 里的最新版本自动生成下一版 DOCX/PDF/报告
python "E:\My Project\thesis-format-checker\delivery\run_delivery.py"

# 测试入口：真实调用业务入口生成下一版，然后跑完整验收
python "E:\My Project\thesis-format-checker\tests\run_full_tests.py"
```

`delivery/run_delivery.py` 是业务 main；`tests/run_full_tests.py` 是测试 main。`src/`、`presets/` 是内部实现或规则数据，不作为日常启动入口。

## 底层检测 CLI

`thesis-check` 是格式检测器的底层 CLI，给业务入口和调试使用，不是论文迭代的主入口。

```bash
# 校验（默认使用 ncwu preset）
thesis-check 论文.docx

# 指定规则集
thesis-check 论文.docx --preset ncwu

# 使用自定义规则文件
thesis-check 论文.docx --rules my-school.yaml

# JSON 输出（CI 友好）
thesis-check 论文.docx --json

# 自动修正可修项，输出到新文件
thesis-check 论文.docx --fix 论文_fixed.docx

# 列出所有规则
thesis-check list-rules --preset ncwu
```

## 退出码

| Code | 含义 |
|------|------|
| 0 | 全部通过 |
| 1 | 有 warning |
| 2 | 有 error |
| 3 | 工具/输入错误 |

## 内置规则

| ID | 说明 | 可自动修正 |
|----|------|-----------|
| page-margins | 页边距 | Y |
| header-text-match | 页眉文本 | Y |
| header-on-all-sections | 所有节有页眉 | - |
| body-font-size | 正文字号 | Y |
| body-east-asia-font | 正文中文字体 | Y |
| body-line-spacing | 正文行距 | Y |
| heading1-style | 一级标题样式 | Y |
| heading2-style | 二级标题样式 | Y |
| heading3-style | 三级标题样式 | Y |
| heading-style-applied | 标题段落使用正确 Heading | Y |
| chapter-page-break | 章前分页 | - |
| abstract-zh-length | 中文摘要字数 | - |
| abstract-en-length | 英文摘要词数 | - |
| foreign-translation-length | 外文译文字数 | - |
| toc-present | 目录存在 | - |
| cover-fields | 封面字段完整 | - |

## 自定义规则

参考 `presets/ncwu.yaml` 和 `examples/custom-school.yaml` 编写你自己学校的规则集。

## License

MIT
