# 自定义模板

[variables]
title: 在此输入标题
author: 在此输入作者
date: 2026-01-01

[content]
## 在此输入标题1
在此输入正文内容...

## 在此输入标题2
在此输入正文内容...

## 在此输入标题3
在此输入正文内容...

---

## 变量说明

你可以在命令行中使用 -v 参数自定义变量：

```bash
python document_generator.py custom -o mydoc.docx \
  -v title="我的标题" \
  -v author="张三" \
  -v date="2026-02-10"
```

## 可用变量

| 变量名 | 说明 | 默认值 |
|--------|------|--------|
| title | 文档标题 | 在此输入标题 |
| author | 作者/单位 | 在此输入作者 |
| date | 日期 | 2026-01-01 |
| content | 正文内容 | 在此输入正文内容... |

## 自定义方法

1. 复制本模板到新文件
2. 修改模板内容
3. 保存为新文件名
4. 使用时指定新模板名

```bash
python document_generator.py your_template_name -o output.docx
```

## 格式说明

- `#` 开头表示一级标题
- `##` 开头表示二级标题
- `- ` 开头表示列表项
- `| |` 用于创建表格
