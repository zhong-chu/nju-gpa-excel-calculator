# NJU GPA Excel 计算器
:rocket:*还在为南大app上卡顿的GPA计算器功能（~~好像消失了？~~）困扰吗？*

:rocket:*还在被课程种类成绩类型的计算而困扰吗？*

:rocket:*或者期待某门课要是得分多一点gpa会变成什么样子？*

**我做了这样一个工具：这是一个完全在本机运行的成绩Excel计算工具，不需要安装浏览器扩展。**

> 🎉 **v1.1.0 已发布**
>
> 新增按学期筛选、成绩模拟和 GPA 变化对比，并更新为南大紫界面。
>
> [查看更新并下载 v1.1.0](https://github.com/zhong-chu/nju-gpa-excel-calculator/releases/tag/v1.1.0)

## ✨ v1.1.0 更新内容

- 新增按学期筛选
- 新增成绩模拟
- 显示原始 GPA、模拟 GPA 和变化量
- 支持一键恢复原始成绩
- 更新为南大紫界面
- 
## :sparkles: 使用方法
1. 点击仓库的 **Code** 按钮，**Download ZIP** 下载完整项目，下载后解压文件夹。
2. 在南京大学教服平台点击成绩信息、成绩查询，然后在“我的成绩”页面选择你需要的学期成绩点击“导出”。
3. 双击 `打开GPA计算器.html`。
4. 选择或拖入刚刚导出的 Excel 文件。
5. 使用“课程性质”按钮整类选择平台、通修、通识等课程，GPA 和已选学分会自动更新。
6. 使用“学期范围”按钮查看全部学期或单独某个学期的课程与 GPA。
7. 点击“开启成绩模拟”，修改课程成绩即可比较原始 GPA 和模拟 GPA；模拟不会修改原始 Excel。
8. 你可以先在线试试看：**[点击打开 GPA 计算器](https://zhong-chu.github.io/nju-gpa-excel-calculator/)**
<p align="center">
  <img
    src="https://github.com/user-attachments/assets/6c6c7938-9955-4efa-9d53-791922429d47"
    alt="new2"
    height="250"
  >
  &nbsp;
  <img
    src="https://github.com/user-attachments/assets/7ea26bda-8c6d-4ad1-9b3b-82503dcb968d"
    alt="new3"
    height="250"
  >
</p>



## :sparkles: 数据规则

1. 优先匹配 Excel 中明确标为“课程性质”的列，因此分类与导出文件保持一致。
2. 支持“课程名 / 课程名称”“总成绩 / 总评成绩 / 成绩”等常见表头。
3. “通过”等非数值成绩会显示，但不参与 GPA。
4. 计算方式：`Σ（成绩 × 学分）÷ Σ学分 ÷ 20`。
5. 模拟成绩仅保存在当前页面内，重新选择文件或关闭页面后即清除。

## :sparkles: 隐私

1. Excel 文件仅由本机浏览器读取。工具没有服务器地址，不上传、不保存成绩数据；关闭页面后数据即从内存释放。
2. 请保留文件夹内的 `app.js`、`app.css` 和 `xlsx.full.min.js`，不要只单独移动 HTML 文件。
3. 支持 `.xlsx`、`.xls`、`.xlsb` 和 `.csv`。

## :sparkles: 致谢

本项目的最初想法受到 waterxjw 开发的 `NJU-GPA-Calculator` 启发。感谢原作者的开源分享。原项目是一款面向旧版南京大学教务系统的Chrome 扩展，因为版本等等原因我发现无法使用，

因此，我决定开发这个NJU GPA Excel 计算器项目，希望对大家也是对我自己能够有所帮助！

