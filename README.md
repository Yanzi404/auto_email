# 自动生成周报工具

这是一个自动从邮箱中获取日报并使用AI生成周报的工具。该工具会搜索您邮箱中本周的日报邮件，提取内容，然后使用AI服务生成一份完整的周报，并可以保存到邮箱草稿箱中。

## 功能特点

- 自动从邮箱中获取本周的日报邮件
- 支持从已发送邮件和草稿箱中获取日报
- 使用AI服务自动生成周报内容
- 可选择将生成的周报保存到草稿箱
- 完善的错误处理和日志记录

## 前置条件

- 已开启邮箱的IMAP和SMTP服务

## 安装步骤

1. 克隆或下载本项目到本地
2. 安装依赖包：
   ```
   pip install -r requirements.txt
   ```
3. 复制环境变量模板和邮件内容模版：
   ```
   cp .env.example .env
   ```
   ```
   cp template/template.html.example template/template.html
   ```
   然后编辑`.env`文件，填入您的邮箱和AI服务配置

## 使用方法

直接运行脚本：

```
python auto_weekly_report.py
```

程序会自动：
1. 连接到您的邮箱
2. 搜索本周的日报邮件
3. 使用AI生成周报内容
4. 显示生成的周报内容
5. 询问是否保存到草稿箱

## 配置说明

在`template`文件夹中配置周报格式模版

在`.env`文件中配置以下参数：

### 邮箱配置
- `IMAP_SERVER`: IMAP服务器地址
- `SMTP_SERVER`: SMTP服务器地址
- `EMAIL_USERNAME`: 邮箱用户名
- `EMAIL_PASSWORD`: 邮箱密码
- `SENT_FOLDER`: 已发送邮件文件夹名称
- `DRAFTS_FOLDER`: 草稿箱文件夹名称

### AI配置
- `AI_API_KEY`: AI服务API密钥
- `AI_API_URL`: AI服务API地址
- `AI_MODEL`: 使用的AI模型
- `AI_SYSTEM_PROMPT`: 系统提示词

### 周报配置
- `REPORT_TITLE_PREFIX`: 周报标题前缀
- `REPORT_TITLE_DATE_FORMAT`: 周报标题中日期的格式，默认为"%Y-%m-%d"
- `REPORT_DEFAULT_TO`: 默认收件人邮箱地址
- `REPORT_DEFAULT_CC`: 默认抄送人邮箱地址，多个邮箱用逗号分隔


## 注意事项

- 请确保您的邮箱允许IMAP和SMTP访问
- 请妥善保管您的邮箱密码和API密钥，不要将其提交到公共仓库
- 如果使用企业邮箱，可能需要管理员开启相关权限