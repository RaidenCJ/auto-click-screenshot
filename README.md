# Auto Click Screenshot 🖱️📸

自动点击截图工具 - 定时执行点击操作并截图发送到飞书

## 功能特性

- ✅ 自动点击屏幕指定位置
- ✅ 点击后自动截图
- ✅ 截图发送到飞书机器人
- ✅ 支持 Windows 任务计划触发
- ✅ 支持远程桌面断开后执行（向日葵/ToDesk 最小化可用）
- ✅ 支持 OpenClaw 定时任务调度
- ✅ 仅工作日执行（周一至周五）

## 文件说明

| 文件 | 说明 |
|-----|------|
| `auto_click.ps1` | 主脚本 - 执行点击、截图、发送飞书 |
| `trigger_click.ps1` | 触发脚本 - 通过 Windows 任务计划触发主脚本 |
| `set_coords.ps1` | 设置坐标 - 交互式设置 5 个点击位置 |
| `test_click.ps1` | 测试运行 - 10 秒后执行测试 |
| `auto_click_config.json` | 配置文件 - 存储坐标和飞书凭证 |

## 快速开始

### 1. 配置飞书机器人

1. 在飞书开放平台创建企业自建应用
2. 获取 `App ID`、`App Secret`
3. 获取接收消息的用户 `Open ID`
4. 编辑 `auto_click_config.json` 填入凭证：

```json
{
    "app_id": "cli_xxxxx",
    "app_secret": "xxxxxxxx",
    "user_id": "ou_xxxxx",
    "coordinates": [...]
}
```

### 2. 设置点击坐标

运行坐标设置脚本，依次移动鼠标到 5 个目标位置：

```powershell
powershell -ExecutionPolicy Bypass -File "set_coords.ps1"
```

每个位置有 4 秒时间移动鼠标，自动捕获坐标。

**位置说明**：
- 位置 1：主按钮（点击后等待 4 秒，不截图）
- 位置 2-4：点击后截图发送
- 位置 5：结束按钮（只点击，不截图）

### 3. 测试运行

```powershell
powershell -ExecutionPolicy Bypass -File "test_click.ps1"
```

10 秒后自动执行，检查飞书是否收到截图。

### 4. 设置定时任务

通过 OpenClaw 设置定时任务：

```
每天 09:00 执行自动点击截图任务
```

## 执行流程

```
位置 1 (主按钮) → 点击 → 等待 4 秒
位置 2 → 点击 → 等待 2 秒 → 截图 → 发送飞书
位置 3 → 点击 → 等待 2 秒 → 截图 → 发送飞书
位置 4 → 点击 → 等待 2 秒 → 截图 → 发送飞书
位置 5 → 点击 (不截图)
```

共发送 **3 张截图** 到飞书。

## 远程桌面支持

### 向日葵/ToDesk 最小化执行

本工具通过 Windows 任务计划触发，支持在远程软件最小化或断开后继续执行。

**原理**：
- OpenClaw Cron 触发 `trigger_click.ps1`
- 创建立即执行的 Windows 任务计划
- 任务计划在用户会话中运行主脚本
- 绕过远程软件的输入限制

### Windows RDP 断开

Windows 远程桌面断开后会话会挂起，需要使用以下方法保持会话活跃：

```powershell
# 方法：使用 tscon 断开会话但保持活跃
query session  # 查看会话 ID
tscon <session_id> /dest:console
```

## 配置文件说明

### auto_click_config.json

```json
{
    "app_id": "飞书应用 ID",
    "app_secret": "飞书应用密钥",
    "user_id": "接收消息的用户 Open ID",
    "webhook_url": "Webhook URL（可选）",
    "coordinates": [
        { "x": 186, "y": 124 },  // 位置 1
        { "x": 125, "y": 187 },  // 位置 2
        { "x": 96, "y": 273 },   // 位置 3
        { "x": 87, "y": 339 },   // 位置 4
        { "x": 725, "y": 1055 }  // 位置 5
    ],
    "click_interval": 3,
    "last_updated": "2026-03-29 18:28:17"
}
```

## 定时任务

默认配置的工作日定时任务（周一至周五）：

| 时间 | 时间 |
|-----|------|
| 09:00 | 14:15 |
| 09:30 | 14:45 |
| 10:00 | 15:00 |
| 10:45 | 21:00 |
| 11:15 | 21:30 |
| 13:45 | 22:00, 22:30, 23:00 |

## 截图保存位置

截图自动保存到桌面：

```
C:\Users\<用户名>\Desktop\auto_click_screenshots\
```

## 故障排除

### 点击不生效

1. 检查坐标是否正确：运行 `set_coords.ps1` 重新设置
2. 检查目标应用是否有管理员权限
3. 尝试手动执行脚本测试

### 截图黑屏

1. 确保显示器连接或远程会话活跃
2. 向日葵最小化后截图应该正常
3. RDP 断开后需要使用 tscon 保持会话

### 飞书发送失败

1. 检查 App ID 和 App Secret 是否正确
2. 检查用户 Open ID 是否正确
3. 检查飞书应用是否有发送消息权限

## 安全提示

- 请勿在公开仓库中提交包含敏感信息的配置文件
- 建议使用 `.gitignore` 排除 `auto_click_config.json`
- 飞书凭证泄露后请立即重置

## License

MIT
