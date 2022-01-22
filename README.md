# 维格小程序 - Word文档生成器

根据提供的docx模板，一键批量填充字段并合成新的word文档

![cover](cover.jpg)


## 🎨 介绍

本小程序可以将每一行数据填充到 Word 模板里面，从而形成一份新的 Word 文档。同时选中多行记录，即可实现批量导出 Word 文档。

例如一份《录取通知书》。在日常工作中，公司HR一天可能会发送多份《录取通知书》，里面的格式都是一样的，只是“岗位”，“部门”，“候选人姓名”，“通知日期”等等这些信息要素会有所不同，但HR却需要手工重复性地复制粘贴、复制粘贴...

使用本小程序后，只需要提前制作一次 Word 模板，往后的工作就只需要点一点手指头，小程序来帮你填充关键信息要素，并生成新的《录取通知书》！


## 🚀 快速上手（现成模板）

为了让大家可以快速体验到这款小程序的用途，这里已经提前做好了一个维格表模板，包含两个例子，浏览器打开即可体验:

> 体验地址：https://vika.cn/share/shrws2voRW3hGRYffBCbc 

<br>

**如何修改模板**

“聘请函模板”是一个附件字段，将单元格里的模板文件下载到本地，然后用word打开并进行编辑，编辑完成后重新上传覆盖单元格里的旧模板即可。

下图是《入职邀请函》模板里的内容节选。红色高亮的花括号是表格里的字段名称，表示将表格里的对应字段值填充到当前位置。有用过维格表智能公式的用户应该比较好理解。

![模板里的字段](https://s1.vika.cn/space/2022/01/18/b99da6588ed04bafbeb61fb63c6a91e9)

<br/>

**读取「神奇关联」字段的值**

神奇关联需要用“开始标签”和“结束标签”组合起来读取。

开始标签：{#字段名字}

结束标签：{/字段的名字}

 

在开始标签和结束标标签中间，需要使用如下两个标签读取值：

循环读取关联记录的标题名称： {#字段名字}{title}{/字段的名字}

循环读取关联记录的id：{#字段名字}{recordId}{/字段的名字}

<br/>

**读取「神奇引用」字段的值**

在word模板里读取「神奇引用」字段的方式与「神奇关联」类似。但由于被引用的字段类型是多种多样的，具体如何适配，请通过console打印调试。


<br/>

**成员字段如何取值？**

成员字段的获取方式跟「神奇关联」类似：

循环读取成员字段的成员姓名　{#字段名字}{name}{/字段的名字}

<br/>

## 🙋‍♂️ 常见问题

**word模板修改完毕后需要重新上传，是每一行都要上传一次吗？**

是的。一行数据代表着独立的一份word文档，需要单独配置一个模板。tips：你可以拖动单元格右下角的“把手（小方块）”，进行快速的填充模板附件。

<br/>

**如何将「word文档生成器」小程序添加到自己空间站的其他表格里？**

「word文档生成器」已经上架到小程序中心，你可以直接安装。

<br/>

**使用Mac系统的Safari浏览器访问小程序，无法进行word文件的批量下载？**

safari的浏览器拦截了，暂不支持进行批量下载，只能一个一个下载。在Mac系统里维格表客户端同样存在这个问题。如果需要批量下载，请使用Chrome或者Edge浏览器。

<br/>

## 🥂 讨论交流

在日常使用中或者二次开发过程中有疑问或者新想法，欢迎前往官方社区的小程序主页留言评论给我~

👉 [点我跳转「Word文档生成器」的主页](https://bbs.vika.cn/article/111)

<br/>

## 🎯 更新日志
v0.1.5 - 2022年1月12日
- 【修复】loop index无法正常显示的问题

v0.1.4 - 2022年1月11日
- 【修复】修复因为useEffect的参数相同导致无法刷新record的问题

v0.1.3 - 2022年1月11日
- 【调整】鼠标点击空白地方后，仍会保留上次已选中的记录（如果没有选中记录，则无法导出word）

v0.1.2 - 2021年12月29日
- 【优化】更新icon远程图片
- 【调整】小程序的背景颜色

v0.1.0 - 2021年12月16日

- 【新增】发布首个版本，当用户在维格视图下选中任意行的时候可以快速导入行数据并依据字段名（```{字段名}```）一一对应替换，生成word文档。

<br/>

## 😍 更多有趣的维格小程序
如果你喜欢学习、折腾各种维格小程序，可以看看维格官方的宝藏库，里面收集有大量的小程序项目、vika API项目：

👉 [awesome-vikadata](https://github.com/vikadata/awesome-vikadata)
