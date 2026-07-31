# 导入文件
Step1 顶部菜单栏找到 Developer → Visual Basic

- 如果你顶部看不到 Developer（开发工具），按下面开启：

File → Options
Customize Ribbon
右侧主列表勾选 Developer → OK
顶部就出现 Developer 标签了。

Step2 打开 VBA 编辑器

- 左侧 Project 窗口，右键你的PERSONAL工作簿（需要先配置全局环境，否则只对单个excel生效）
Insert → Module
右侧空白代码框，粘贴完整代码

- 或者直接import 本地bas文件
- 关闭 VBA 窗口即可

# 配置全局环境
Step1 完全退出所有 Excel

- 关掉所有打开的表格
- 任务栏右键 Excel 图标 → Close all windows，确保 Excel 完全关闭

Step2 新建空白表格，录制临时宏生成 Personal

- 重新打开 Excel，新建空白 Book1
Developer → Record Macro
- 关键下拉：Store macro in → Personal Macro Workbook
Macro name 随便填：Test → OK 开始录制
- 随便点一下 A1 单元格（随便操作一下）
Developer → Stop Recording

Step3 查看 Personal 是否生成（验证）

- 顶部 View 标签 → Unhide
- 如果能看到 PERSONAL.XLSB，代表创建成功；点 Cancel 不用打开

# 设置快捷键
1. Alt+F8 选中 OpenLinkAndCopyAddress → Options

2. Shortcut key 输入 d，快捷键变成 Ctrl+D

3. 以后选中订单行，直接按 Ctrl+D 一键运行
