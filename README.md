# OneDrive 直链解析工具

一个支持个人版、海外企业版、世纪互联版的 OneDrive/SharePoint 直链解析工具。

## 支持的链接类型

- **个人版**: `onedrive.live.com`, `1drv.ms`
- **海外企业版**: `*.sharepoint.com`, `*.my.sharepoint.com`  
- **世纪互联版**: `*.sharepoint.cn`, `*.my.sharepoint.cn`

## 特色功能

- ✅ **优化的世纪互联解析**: 直接生成 `/_layouts/52/download.aspx?share=` 格式链接，无需登录
- ✅ **智能链接识别**: 自动识别文件夹链接并提示
- ✅ **现代化网页界面**: 响应式设计，支持深色模式
- ✅ **命令行工具**: 支持批量处理和脚本集成

## 安装

### 环境要求
- Python 3.8+

### 安装依赖
```bash
pip install -r requirements.txt
```

## 使用方法

启动网站：
```bash
python OneDrive.py
```

打开浏览器访问 `http://127.0.0.1:5000`，粘贴分享链接即可获得直链。

#### 使用示例

在网页中输入分享链接，点击解析即可获得直链：

```
世纪互联链接：
https://xxxx-my.sharepoint.cn/:u:/g/personal/2_xxxx_partner_onmschina_cn/ID

转换后：
https://xxxx-my.sharepoint.cn/personal/2_xxxx_partner_onmschina_cn/_layouts/52/download.aspx?share=ID
```

## 工作原理

### 世纪互联和海外企业版
直接将分享链接格式从：
```
https://domain/:u:/g/personal/user/sharetoken
```
转换为：
```  
https://domain/personal/user/_layouts/52/download.aspx?share=sharetoken
```

这种格式的直链：
- ✅ 无需登录即可下载
- ✅ 长期有效（相对于签名链接）
- ✅ 支持断点续传

### 个人版 OneDrive
1. 展开短链（如 `1drv.ms`）
2. 添加 `download=1` 参数
3. 可选择跟随重定向获取最终签名 URL

## 注意事项

- **仅支持文件链接**: 文件夹分享链接无法生成直链
- **权限限制**: 某些企业租户可能禁用匿名下载
- **个人版时效性**: 个人版生成的签名直链具有时效性
- **网络超时**: 建议根据网络情况调整超时设置

## 项目结构

```
onedrive直链/
├── OneDrive.py         # 主程序（包含命令行和网页版）
├── templates/
│   └── index.html      # 网页模板
├── requirements.txt    # 依赖文件
└── README.md          # 说明文档
```

## 开发

### 本地开发
```bash
# 开发模式启动网站
python OneDrive.py
```

### 部署到生产环境
```bash
# 使用 Gunicorn
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:5000 "OneDrive:app"

# 或使用 Waitress (Windows 友好)
pip install waitress
waitress-serve --host=0.0.0.0 --port=5000 "OneDrive:app"

# 或直接修改代码中的 host 和 port
# 在 OneDrive.py 最后一行修改为：
# app.run(host="0.0.0.0", port=5000)
```

## 许可证

MIT License