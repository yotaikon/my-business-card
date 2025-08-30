# 个人技术简历生成器

一个基于Python的自动化简历生成工具，可以从Excel文件中读取个人信息并生成美观的HTML简历页面。
个⼈主⻚了：https://yotaikon.github.io/my-business-card/

## 🌟 功能特性

### 📊 数据源支持
- 支持读取Excel文件（.xlsx格式）
- 自动提取个人信息、技能、项目经历等
- 当Excel文件不存在时，使用默认数据生成页面

### 🎨 美观的界面设计
- 现代化的渐变背景设计
- 响应式布局，支持移动端访问
- 卡片式布局，信息层次清晰
- 悬停动画效果，提升用户体验

### 📋 完整的信息展示
- **基本信息**：姓名、职位、联系方式、年龄、国籍等
- **教育背景**：学历信息、专业、学位等
- **语言能力**：多语言能力矩阵，包含阅读、写作、会话等级
- **专业技能**：分类展示操作系统、数据库、开发工具、编程语言等技能
- **项目经历**：详细的项目描述、技术栈、团队规模、角色等信息

### 🚀 自动化部署
- 集成GitHub Actions自动构建和部署
- 自动部署到GitHub Pages
- 支持持续集成/持续部署（CI/CD）

## 📁 项目结构

```
my-business-card/
├── build_page.py          # 主要的Python脚本
├── requirements.txt       # Python依赖包
├── index.html            # 生成的HTML简历页面
├── 技術者経歴書(Y.TK).xlsx # Excel数据源文件
├── .github/
│   └── workflows/
│       └── deploy.yml     # GitHub Actions工作流配置
└── README.md             # 项目说明文档
```

## 🛠️ 安装和设置

### 环境要求
- Python 3.7+
- pip包管理器

### 安装步骤

1. **克隆项目**
   ```bash
   git clone https://github.com/your-username/my-business-card.git
   cd my-business-card
   ```

2. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```

3. **准备数据文件**
   - 将您的Excel文件命名为 `技術者経歴書(Y.TK).xlsx`
   - 或者修改 `build_page.py` 中的文件名

## 📖 使用方法

### 本地运行

1. **生成HTML页面**
   ```bash
   python build_page.py
   ```

2. **查看结果**
   - 在浏览器中打开生成的 `index.html` 文件
   - 或者使用本地服务器：
     ```bash
     python -m http.server 8000
     # 然后访问 http://localhost:8000
     ```

### 自动部署

1. **推送到GitHub**
   ```bash
   git add .
   git commit -m "Update resume"
   git push origin main
   ```

2. **查看部署结果**
   - GitHub Actions会自动构建和部署
   - 访问您的GitHub Pages URL查看结果

## 📊 数据格式说明

### Excel文件结构
Excel文件应包含以下信息：

- **基本信息**：姓名、年龄、国籍、联系方式等
- **教育背景**：学校、专业、学位、时间等
- **技能信息**：按类别组织的技能和熟练程度
- **语言能力**：各语言的阅读、写作、会话能力
- **项目经历**：项目名称、时间、描述、技术栈等

### 技能等级说明
- **A级**：熟练 - 可以独立完成相关工作
- **B级**：可实践 - 有实践经验，可以参与项目
- **C级**：有经验 - 有一定经验，需要指导
- **D级**：基础知识 - 了解基本概念

### 语言等级说明
- **A级**：母语级 - 可以流利使用
- **B级**：商务级 - 可以用于商务交流
- **C级**：业务可用 - 可以用于基本业务交流
- **D级**：稍有困难 - 需要帮助才能交流

## 🎨 自定义样式

### 修改颜色主题
在 `build_page.py` 中找到CSS样式部分，可以修改：
- 背景渐变色
- 卡片颜色
- 技能等级颜色
- 字体和间距

### 添加新的信息模块
1. 在 `extract_personal_info` 函数中添加数据提取逻辑
2. 在 `generate_html` 函数中添加HTML生成代码
3. 添加相应的CSS样式

## 🔧 技术栈

- **后端**：Python 3.7+
- **数据处理**：pandas, openpyxl
- **前端**：HTML5, CSS3, JavaScript
- **部署**：GitHub Actions, GitHub Pages
- **版本控制**：Git

## 📝 配置说明

### GitHub Pages设置
1. 在GitHub仓库设置中启用Pages功能
2. 选择GitHub Actions作为部署源
3. 确保 `.github/workflows/deploy.yml` 文件存在

### 自定义域名（可选）
1. 在GitHub Pages设置中添加自定义域名
2. 在仓库根目录添加 `CNAME` 文件

## 🤝 贡献指南

欢迎提交Issue和Pull Request来改进这个项目！

### 开发流程
1. Fork项目
2. 创建功能分支
3. 提交更改
4. 创建Pull Request

### 代码规范
- 使用Python PEP 8代码规范
- 添加适当的注释
- 确保代码可读性

## 📄 许可证

本项目采用MIT许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 🙏 致谢

感谢所有为这个项目做出贡献的开发者和用户！

## 📞 联系方式

如果您有任何问题或建议，请通过以下方式联系：

- 创建GitHub Issue
- 发送邮件至：[your-email@example.com]

---

**注意**：请确保在部署前移除或修改个人信息中的敏感数据。
