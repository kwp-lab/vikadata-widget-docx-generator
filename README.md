# 维格小程序 - Word文档生成器

根据提供的docx模板，一键批量填充字段并合成新的word文档

![package icon](package_icon.png)


## 🎨 介绍

本小程序可以将每一行数据填充到 Word 模板里面，从而形成一份新的 Word 文档。同时选中多行记录，即可实现批量导出 Word 文档。

例如一份《录取通知书》。在日常工作中，公司HR一天可能会发送多份《录取通知书》，里面的格式都是一样的，只是“岗位”，“部门”，“候选人姓名”，“通知日期”等等这些信息要素会有所不同，但HR却需要手工重复性地复制粘贴、复制粘贴...

使用本小程序后，只需要提前制作一次 Word 模板，往后的工作就只需要点一点手指头，小程序来帮你填充关键信息要素，并生成新的《录取通知书》！


## 💡 使用步骤

1. 提前准备好 Word 模板，在 Word 模板里面目标位置填写好维格表里的对应列名，写法跟智能公式里引用单元格值一样，在列名左右两边加上花括号，例如“{候选人姓名}”

2. 将修改好的 Word 模板以附件形式上传到当前维格表的附件列里，如下图示例

    ![示意图](https://s1.vika.cn/space/2021/12/02/22202756884f485dbfce5e257000644c)

3. 在本界面右侧的配置区域选择 Word 模板所在的附件列名

4. 点击右上角按钮，退出小程序的“展开模式”

5. 在维格视图中选择若干行，然后点击小程序的“导出 Word 文档”



## 🎯 更新日志
v0.1.4 - 2022年1月11日
- 【修复】修复因为useEffect的参数相同导致无法刷新record的问题

v0.1.3 - 2022年1月11日
- 【调整】鼠标点击空白地方后，仍会保留上次已选中的记录（如果没有选中记录，则无法导出word）

v0.1.2 - 2021年12月29日
- 【优化】更新icon远程图片
- 【调整】小程序的背景颜色

v0.1.0 - 2021年12月16日

- 【新增】发布首个版本，当用户在维格视图下选中任意行的时候可以快速导入行数据并依据字段名（```{字段名}```）一一对应替换，生成word文档。

