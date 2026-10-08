# 表数通

一个可配置的桌面数据工具（Windows / macOS），用于从一个或多个 Google 表格链接中遍历子 Sheet、映射字段、汇总、查询、数据分析，以及按时间增量提取并去重。

## 运行

```powershell
python app.py
```

公开表格无需凭据（需要开启“知道链接的任何人可查看”）。私有表格在设置页添加一个或多个 Google 服务账号 JSON，并把目标表格共享给每一个服务账号邮箱。多个账号会按顺序轮询，遇到 429 自动切换。

时间提取支持输出到本地 Excel 或指定 Google 表格链接/子工作表。Google 私有表格读取会把多个子 Sheet 合并为批量请求；遇到 429、5xx 时自动指数退避重试。

查询结果按当前表的表头显示。识别为号码数据表时，会提供修正格式，并可复制号码列。时间提取写入已有 Google 子工作表时，会按字段别名识别同义表头，并遵循目标工作表现有列顺序。

应用内「检查并安装更新」会直接下载当前系统的安装包并启动安装，不会打开网页。

## 数据分析

侧栏「数据分析」使用你在「数据源」里已经配置好的表格。日期列、名字列可在下拉里切换。

- 队别列、日期列、名字列都可在下拉里切换。队别不选 = 全部队别；名字留空 = 该队全员。两者可单独筛选。
- 默认时间是当月、不对比。可选最近 7 天、最近 2 天或自定义；勾选「对比」后才显示对比期。
- 曲线上方用小方框勾选要画的列；太多时点「更多」。勾选分类列会按不同取值各画一条线。
- 名称含「场」的列只统计单元格里含 D 的条数，不按时长拆线、也不累计分钟。
- 鼠标移到曲线节点或饼图扇区上，会悬浮显示准确数字。
- 每日表在对比开启时同时列出本期和对比期；增减绿色为升、红色为降。
- 图表可切换曲线或饼图。饼图扇区上显示百分比，并用线连到名称。
- 排除关键词与数据查询相同，多个词用逗号分隔，命中任一则整行不进入统计。
- 默认按记录数看总数和日均。
- 查询和分析都先读本地库。点「同步本地库」会下载表格最新数据并替换本地库；内容完全相同则跳过写入。

## 子 Sheet 规则

- “指定 Sheet”留空：遍历全部，只跳过排除项。
- “指定 Sheet”有值：只读取指定项。
- 同时出现在指定和排除列表时，排除规则优先。

## 下载

从 [Releases](https://github.com/christiancagfr-alt/SheetDataHub/releases) 页面下载最新版本：

| 文件 | 说明 |
|------|------|
| `SheetDataHub-Setup-v1.4.3.exe` | Windows 安装程序（推荐） |
| `SheetDataHub-windows-v1.4.3.zip` | Windows 免安装便携包，解压后运行 `SheetDataHub/SheetDataHub.exe` |
| `SheetDataHub-macos-arm64-v1.4.3.dmg` | macOS 安装盘：打开后把 `SheetDataHub.app` 拖进「应用程序」 |
| `SheetDataHub-macos-arm64-v1.4.3.zip` | macOS 便携包 |

系统要求：Windows 10 / 11（64 位），或 macOS 12+（Apple Silicon）。配置 Developer ID 证书后，CI 会签名并公证安装盘；未配置时 Mac 需按住 Control 单击再打开。

## 验证软件来源

下载后，使用 GitHub CLI 验证文件确实由官方 CI 构建、且未被篡改：

```bash
gh attestation verify ./SheetDataHub-windows-v1.4.3.zip --repo christiancagfr-alt/SheetDataHub
gh attestation verify ./SheetDataHub-Setup-v1.4.3.exe --repo christiancagfr-alt/SheetDataHub
gh attestation verify ./SheetDataHub-macos-arm64-v1.4.3.dmg --repo christiancagfr-alt/SheetDataHub
```

验证成功表示该软件确实由官方 GitHub Actions 构建。

## 打包（本地）

Windows：

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1
```

macOS：

```bash
chmod +x build_macos.sh
./build_macos.sh
```

## 如何发布新版本

发布仓库：https://github.com/christiancagfr-alt/SheetDataHub  
应用内「检查并安装更新」读取该仓库的 GitHub Releases，并直接下载安装。

本项目使用 GitHub Actions 自动构建 Windows 与 macOS 包。每次发布新版本只需要创建一个 Git Tag 并推送到个人仓库。

### 发布步骤

#### 1. 确保代码已提交并推送

在发布之前，确保你的所有代码改动已经提交并推送到 GitHub：

```bash
# 查看当前状态
git status

# 添加所有改动
git add .

# 提交改动（把"你的改动说明"替换成实际的描述）
git commit -m "你的改动说明"

# 推送到 GitHub
git push origin main
```

#### 2. 创建版本 Tag

Git Tag 是一个版本标记，用于标识发布的版本号。版本号格式为 `v主版本.次版本.修订版本`，例如 `v1.0.0`、`v1.1.0`、`v2.0.0`。

```bash
# 创建一个新的版本 tag（将 v1.0.1 替换为你想要的版本号）
git tag -a v1.0.1 -m "Release version 1.0.1"
```

#### 3. 推送 Tag 触发自动构建

```bash
# 推送 tag 到 GitHub（这会自动触发 CI 构建）
git push origin v1.0.1
```

推送后，GitHub Actions 会自动执行以下操作：
1. 构建项目
2. 若已配置 Developer ID Secrets，则签名并公证 macOS 安装盘；否则打出未签名包
3. 生成安全签名（Attestation）
4. 创建 Release 并上传构建产物

### macOS 签名 Secrets

发布 macOS 包需要 Apple Developer Program 的 **Developer ID Application** 证书，以及 App Store Connect API 密钥。在个人仓库 `christiancagfr-alt/SheetDataHub` 的 Settings → Secrets and variables → Actions 中添加：

| Secret | 内容 |
|--------|------|
| `MACOS_CERTIFICATE_P12_BASE64` | Developer ID Application 的 `.p12` 做 Base64 |
| `MACOS_CERTIFICATE_PASSWORD` | 导出该 `.p12` 时设置的密码 |
| `MACOS_CODESIGN_IDENTITY` | 可选。例如 `Developer ID Application: 名称 (TEAMID)`。留空则自动识别 |
| `APPLE_API_KEY_P8_BASE64` | App Store Connect API 的 `.p8` 做 Base64 |
| `APPLE_API_KEY_ID` | API Key ID |
| `APPLE_API_ISSUER` | Issuer ID（UUID） |

Windows 上可以把文件转成 Base64 后写入 Secret：

```powershell
[Convert]::ToBase64String([IO.File]::ReadAllBytes("证书.p12"))
[Convert]::ToBase64String([IO.File]::ReadAllBytes("AuthKey_XXXXXX.p8"))
```

#### 4. 查看构建结果

- 构建进度：访问项目的 **Actions** 页面查看
- 发布结果：访问项目的 **Releases** 页面查看已发布的文件

### 版本号说明

| 版本号格式 | 什么时候用 | 对应形式 |
|-----------|-----------|------|
| `vX.0.0` | 重大更新、不兼容改动 | `v2.0.0` |
| `vX.Y.0` | 新增功能 | `v1.1.0` |
| `vX.Y.Z` | 修复 bug | `v1.0.1` |

### 如果构建失败怎么办

1. 访问项目的 **Actions** 页面查看错误日志
2. 修复代码问题
3. 删除失败的 tag 并重新创建：

```bash
# 删除本地 tag
git tag -d v1.0.1

# 删除远程 tag
git push origin :refs/tags/v1.0.1

# 修复问题后，重新创建并推送
git tag -a v1.0.1 -m "Release version 1.0.1"
git push origin v1.0.1
```
