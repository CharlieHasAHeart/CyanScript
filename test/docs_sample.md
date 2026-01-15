# 1. 项目背景

软件著作权申请要求提交规范的说明书，但排版复杂、格式繁琐。CyanScript 通过 Markdown + 模板的方式，让内容与样式解耦。

# 2. 功能概览

解析 Markdown 并映射到模板样式，标题分别对应“标题 1~4”，正文对应“正文”样式。

模板占位符会被替换：封面与页眉的 {{ software_name }} 与 {{ version }}。

# 3. 系统架构

CyanScript 包含三大核心模块。

## 3.1 模板渲染

基于 docxtpl 渲染模板并注入主内容。

## 3.2 Markdown 解析

将 Markdown 转换为 HTML，再映射为 Word 段落与图片。

# 4. 效果示例

![示例图片](./assets/diagram.png)

# 4.1 第二张图示例

![流程示意图](./assets/flow.png)

# 4.2 代码块示例

```python
def hello(name: str) -> str:
    return f"Hello, {name}"

def add(a: int, b: int) -> int:
    return a + b

def factorial(n: int) -> int:
    if n <= 1:
        return 1
    return n * factorial(n - 1)

def fib(n: int) -> int:
    a, b = 0, 1
    for _ in range(n):
        a, b = b, a + b
    return a

def format_items(items):
    return [f"item-{i}" for i in items]

def main() -> None:
    names = ["Alice", "Bob", "Charlie"]
    for name in names:
        print(hello(name))
    print(add(2, 3))
    print(factorial(5))
    print(fib(10))
    print(format_items(range(5)))

if __name__ == "__main__":
    main()
```

# 4.3 行内代码示例

请在终端运行 `python cyan_script.py` 并确认输出路径为 `output/软件名_版本号_软件说明书.docx`。

# 4.4 列表示例

有序列表：

1. 安装依赖
2. 配置模板路径
3. 生成文档

无序列表：

- 统一样式管理
- 模板占位符替换
- 自动更新目录

# 4.5 引用块示例

> 说明书中引用内容建议使用统一样式，便于审阅与排版一致性。
> 支持多行引用块内容。

# 4.6 提示/注意/警告示例

> 提示：请先配置模板路径再生成文档。
> 注意：模板样式变更后需要重新渲染。
> 警告：不要覆盖未备份的输出文件。

# 5. 表格示例

表 1 模块功能对照表

| 模块          | 功能说明                     | 负责人 |
| ------------- | ---------------------------- | ------ |
| 模板渲染      | docxtpl 替换占位符并插入正文 | 张三   |
| Markdown 解析 | 标题、正文、图片与表格转换   | 李四   |
| 输出保存      | 生成最终 docx 并自动更新目录 | 王五   |

# 5.1 第二个表格

表 2 运行环境对照表

| 依赖     | 版本  |
| -------- | ----- |
| Python   | 3.8+  |
| docxtpl  | 0.16+ |
| markdown | 3.4+  |

# 6. 结论

通过统一的输入与配置文件，生成符合标准的软著说明书。
