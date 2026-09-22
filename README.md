# NJU GPA Excel 计算器
:rocket:*还在为南大app上卡顿的GPA计算器功能（~~好像消失了？~~）困扰吗？*

:rocket:*还在被课程种类成绩类型的计算而困扰吗？*

:rocket:*或者期待某门课要是得分多一点gpa会变成什么样子？*

**我做了这样一个工具：这是一个完全在本机运行的成绩Excel计算工具，不需要安装浏览器扩展。**

> 本项目是非官方辅助工具，与南京大学官方无隶属关系；计算结果仅供个人参考。


## :sparkles: 使用方法
1. 点击仓库的 **Code** 按钮，**Download ZIP** 下载完整项目，下载后解压文件夹。
2. 在南京大学教服平台点击成绩信息、成绩查询，然后在“我的成绩”页面选择你需要的学期成绩点击“导出”。
3. 双击 `打开GPA计算器.html`。
4. 选择或拖入刚刚导出的 Excel 文件。
5. 使用“课程性质”按钮整类选择平台、通修、通识等课程，GPA 和已选学分会自动更新。
6. 你可以先在线试试看：**(https://zhong-chu.github.io/nju-gpa-excel-calculator/)**
<p align="center">
  <img
    src="https://github.com/user-attachments/assets/f1c7a256-b73f-426d-8f28-6910939f7be1"
    alt="GitHub 1"
    height="250"
  >
  &nbsp;
  <img
    src="https://github.com/user-attachments/assets/c26a87ee-4ef3-47e0-9c5c-223c3b2a2345"
    alt="GitHub 2"
    height="250"
  >
</p>


## :sparkles: 数据规则

1. 优先匹配 Excel 中明确标为“课程性质”的列，因此分类与导出文件保持一致。
2. 支持“课程名 / 课程名称”“总成绩 / 总评成绩 / 成绩”等常见表头。
3. “通过”等非数值成绩会显示，但不参与 GPA。
4. 计算方式：`Σ（成绩 × 学分）÷ Σ学分 ÷ 20`。

## :sparkles: 隐私

1. Excel 文件仅由本机浏览器读取。工具没有服务器地址，不上传、不保存成绩数据；关闭页面后数据即从内存释放。
2. 请保留文件夹内的 `app.js`、`app.css` 和 `xlsx.full.min.js`，不要只单独移动 HTML 文件。
3. 支持 `.xlsx`、`.xls`、`.xlsb` 和 `.csv`。

## :sparkles: 致谢

本项目的最初想法受到 waterxjw 开发的 `NJU-GPA-Calculator` 启发。感谢原作者的开源分享。原项目是一款面向旧版南京大学教务系统的Chrome 扩展，因为版本等等原因我发现无法使用，

因此，我决定开发这个NJU GPA Excel 计算器项目，希望对大家也是对我自己能够有所帮助！


