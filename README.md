# Windows窗口管理器

一个的Windows窗口管理工具，帮助用户快速管理和启动应用程序组合。

## 功能特性

### 📋 应用管理

- 自动检测系统中的应用程序
- 支持从开始菜单和注册表中检测应用
- 提供应用搜索和筛选功能
- 实时显示应用数量统计

### 🎯 组合管理

- 创建自定义应用组合
- 编辑现有组合
- 删除不需要的组合
- 一键运行整个组合中的所有应用
- 实时显示组合数量统计

### 🖥️ 系统托盘

- 最小化到系统托盘
- 从托盘快速打开主窗口
- 从托盘直接运行应用组合
- 自动更新托盘右键菜单，反映最新的组合列表

### 🎨 界面设计

- 现代化的GUI界面
- 响应式布局，支持窗口大小调整
- 美观的标签页设计
- 实时状态信息显示

## 系统要求

- Windows 10/11 64位系统
- Python 3.8+
- 至少 1GB 内存
- 至少 50MB 可用磁盘空间

## 安装方法

### 方法一：使用已打包的可执行文件

下载链接：

```
通过网盘分享的文件：TableManager.zip
链接: https://pan.baidu.com/s/13XTCR0hQYcNIPqiuRns6CQ?pwd=HENG 提取码: HENG 
```

双击运行 `窗口管理器.exe` 即可

软件使用方法见“软件使用说明书.txt”

### 方法二：源代码-运行/自定义修改

1. 克隆或下载本项目到本地
2. 安装依赖：

   ```bash
   pip install -r requirements.txt
   ```
3. 运行应用：

   ```bash
   python app.py
   ```
4. 打包成可执行文件

   如果需要将应用打包成独立的可执行文件，可使用 `pyinstaller` 工具。以下是打包命令示例：

   ```bash
   pip install pyinstaller
   pyinstaller --onefile --windowed --name=窗口管理器 app.py
   ```

   打包完成后，在 `dist` 目录下会生成 `窗口管理器.exe` 文件。

## 使用说明

### 1. 首次运行

- 应用会自动检测系统中的应用程序，可能需要几秒钟时间
- 检测完成后，应用列表会显示在"应用管理"标签页中

### 2. 创建应用组合

1. 切换到"组合管理"标签页
2. 点击"创建组合"按钮
3. 输入组合名称
4. 在应用列表中选择要包含在组合中的应用（可多选）
5. 点击"保存"按钮
6. 新创建的组合会自动出现在组合列表和系统托盘右键菜单中

### 3. 编辑应用组合

1. 在"组合管理"标签页中选择要编辑的组合
2. 点击"编辑组合"按钮
3. 修改组合名称和选择的应用
4. 点击"保存"按钮
5. 修改后的组合会自动更新到系统托盘右键菜单中

### 4. 删除应用组合

1. 在"组合管理"标签页中选择要删除的组合
2. 点击"删除组合"按钮
3. 确认删除操作
4. 删除后的组合会自动从系统托盘右键菜单中移除

### 5. 运行应用组合

#### 方法一：从主界面运行

1. 在"组合管理"标签页中选择要运行的组合
2. 点击"运行组合"按钮
3. 组合中的所有应用会依次启动

#### 方法二：从系统托盘运行

1. 点击系统托盘中的"Windows窗口管理器"图标
2. 在右键菜单中选择要运行的组合
3. 组合中的所有应用会依次启动

## 数据存储

应用数据存储在 `app_data.json` 文件中，包含：

- 检测到的应用程序列表
- 创建的应用组合列表

## 技术实现

- **前端框架**：tkinter + ttkbootstrap
- **系统集成**：winreg、subprocess
- **系统托盘**：pystray
- **数据存储**：JSON
- **多线程**：用于后台应用检测

## 开发环境

```bash
# 安装依赖
pip install ttkbootstrap pystray pillow psutil

# 打包成可执行文件
pip install pyinstaller
pyinstaller --onefile --windowed --name=窗口管理器 app.py
```

## 贡献指南

欢迎提交 Issue 和 Pull Request 来帮助改进这个项目！

### 开发流程

1. Fork 本项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开 Pull Request

## 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

## 联系方式

- 项目地址：[[https://github.com/Heng-DayDayUp/TableManager]](https://github.com/Heng-DayDayUp/TableManager) 
- 开发者：HENG-DAYDAYUP

---

**感谢使用 Windows 窗口管理器！** 🎉
