# OSEM 插件 Python 脚本接口定义规范

本文档定义了 OSEM (Outlook Event Monitor) 插件与外部 Python 脚本之间的交互接口。

## 1. 调用机制

插件通过 Windows 命令行启动 Python 解释器来执行脚本。

- **执行命令**: `python.exe "<ScriptPath>" "<ContextJsonPath>"`
- **工作目录**: 脚本所在的目录。
- **参数**:
  1. `ScriptPath`: Python 脚本的绝对路径。
  2. `ContextJsonPath`: 包含当前事件上下文信息的临时 JSON 文件路径（绝对路径）。

## 2. 输入数据 (Context JSON)

插件生成的临时 JSON 文件采用 UTF-8 编码，根对象包含以下字段：

### 2.1 根对象结构

| 字段名 | 类型 | 说明 |
| :--- | :--- | :--- |
| `EventId` | string | 事件的唯一标识符 (GUID)。 |
| `EventTitle` | string | 事件的标题（通常是邮件主题或自定义标题）。 |
| `DashboardTemplateId` | string? | 当前使用的仪表盘模板 ID（可能为 null）。 |
| `DashboardValues` | object | 键值对字典，包含当前仪表盘已有的字段值。 |
| `Emails` | array | 相关邮件列表（按时间倒序或相关性排序）。 |
| `Attachments` | array | 相关附件列表。 |

### 2.2 EmailItem 对象结构 (在 `Emails` 数组中)

| 字段名 | 类型 | 说明 |
| :--- | :--- | :--- |
| `EntryId` | string | Outlook 邮件对象的唯一 EntryID (关键字段，用于 COM 调用)。 |
| `StoreId` | string | 邮件所在的存储区 ID。 |
| `Subject` | string | 邮件主题。 |
| `Sender` | string | 发件人地址/名称。 |
| `Participants` | array[string] | 所有参与者（收件人、抄送）的列表。 |
| `ReceivedOn` | string | 接收时间 (ISO 8601 格式)。 |
| `BodyFingerprint` | string | 邮件正文指纹（用于去重）。 |
| `InternetMessageId` | string | 邮件的 Internet Message ID。 |
| `ConversationId` | string | 会话 ID。 |

### 2.3 AttachmentItem 对象结构 (在 `Attachments` 数组中)

| 字段名 | 类型 | 说明 |
| :--- | :--- | :--- |
| `Id` | string | 附件的唯一标识。 |
| `FileName` | string | 文件名（包含扩展名）。 |
| `FileType` | string | 文件扩展名（如 `.pdf`）。 |
| `FileSizeBytes` | number | 文件大小（字节）。 |
| `SourceMailEntryId` | string | 该附件所属邮件的 EntryID。 |

### 2.4 输入示例

```json
{
  "EventId": "550e8400-e29b-41d4-a716-446655440000",
  "EventTitle": "Shipment #12345 Update",
  "DashboardTemplateId": "logistics_template_v1",
  "DashboardValues": {
    "HAWB": "H123456789",
    "Status": "In Transit"
  },
  "Emails": [
    {
      "EntryId": "000000001A4473...",
      "Subject": "Re: Shipment #12345",
      "Sender": "agent@logistics.com",
      "ReceivedOn": "2023-10-27T10:30:00"
    }
  ],
  "Attachments": [
    {
      "FileName": "Invoice.pdf",
      "SourceMailEntryId": "000000001A4473..."
    }
  ]
}
```

## 3. 全局脚本模式 (Global Scripts)

当脚本被配置为“全局脚本”并在主界面执行时，输入数据的结构会有所不同。

### 3.1 识别全局模式
通过检查 `DashboardValues` 中的 `IsGlobalExecution` 字段来判断是否为全局执行模式。

```json
{
  "EventId": "GLOBAL",
  "EventTitle": "Global Execution",
  "DashboardValues": {
    "IsGlobalExecution": "true",
    "GlobalDataPath": "C:\\Users\\User\\AppData\\Local\\Temp\\osem-global-data-xxxx.json"
  },
  "Emails": [],
  "Attachments": []
}
```

### 3.2 获取全局数据
全局模式下，实际的事件列表存储在 `GlobalDataPath` 指向的另一个 JSON 文件中。该文件包含一个 `EventRecord` 对象的数组。

**重要提示**: 全局数据文件直接序列化了 C# 的 `EventRecord` 对象，因此其结构与单事件模式下的 `Context` 略有不同，包含的信息更加完整。

**Global Data JSON 结构 (Array of EventRecord):**
```json
[
  {
    "EventId": "550e8400-...",
    "EventTitle": "Shipment #12345",
    "Status": "Open",
    "DashboardTemplateId": "logistics_v1",
    "CreatedOn": "2023-10-01T10:00:00Z",
    "LastUpdatedOn": "2023-10-05T14:30:00Z",
    
    // 注意：这里是 DashboardItems 数组，而不是 DashboardValues 字典
    "DashboardItems": [
      { "Key": "HAWB", "Value": "H123", "Confidence": 1.0 },
      { "Key": "ETD", "Value": "2023-11-01", "Confidence": 0.8 }
    ],
    
    "Emails": [
      { "EntryId": "...", "Subject": "...", "Sender": "...", "ReceivedOn": "..." }
    ],
    
    "Attachments": [
      { "FileName": "Inv.pdf", "SourceMailEntryId": "..." }
    ],
    
    // 额外字段
    "ConversationIds": [ "..." ],
    "RelatedSubjects": [ "..." ],
    "Participants": [ "..." ]
  },
  ...
]
```

### 3.3 全局脚本读取示例

```python
import sys
import json
import os

def get_dashboard_value(event, key):
    """Helper to get value from EventRecord structure"""
    for item in event.get("DashboardItems", []):
        if item.get("Key") == key:
            return item.get("Value")
    return None

def load_data():
    # 1. 读取基础 Context
    context_path = sys.argv[1]
    with open(context_path, 'r', encoding='utf-8') as f:
        context = json.load(f)

    # 2. 检查是否为全局模式
    dashboard_values = context.get("DashboardValues", {})
    if dashboard_values.get("IsGlobalExecution") == "true":
        # 3. 读取全局数据文件
        global_data_path = dashboard_values.get("GlobalDataPath")
        if global_data_path and os.path.exists(global_data_path):
            with open(global_data_path, 'r', encoding='utf-8') as df:
                all_events = json.load(df)
                return "GLOBAL", all_events
    
    # 普通单事件模式
    return "SINGLE", context

def main():
    mode, data = load_data()
    
    if mode == "GLOBAL":
        print(f"Processing {len(data)} events in GLOBAL mode.")
        for event in data:
            # 访问完整数据
            title = event.get('EventTitle')
            hawb = get_dashboard_value(event, 'HAWB')
            email_count = len(event.get('Emails', []))
            print(f" - [{hawb}] {title} ({email_count} emails)")
            
    else:
        # 单事件模式 (Context 结构)
        print(f"Processing single event: {data.get('EventTitle')}")
        # 单事件模式下直接访问 DashboardValues 字典
        vals = data.get('DashboardValues', {})
        print(f" - HAWB: {vals.get('HAWB')}")

if __name__ == "__main__":
    main()
```

## 4. 脚本行为规范

### 4.1 数据读取
脚本应读取命令行参数传入的文件路径，并解析 JSON 内容。

```python
import sys
import json

def load_context():
    if len(sys.argv) < 2:
        raise ValueError("Missing context file path argument")
    
    context_path = sys.argv[1]
    with open(context_path, 'r', encoding='utf-8') as f:
        return json.load(f)
```

### 4.2 与 Outlook 交互 (可选)
如果 JSON 中的元数据不足（例如需要读取邮件正文或下载附件），可以使用 `pywin32` 库通过 `EntryId` 连接 Outlook。

#### 获取邮件正文
```python
import win32com.client

def get_mail_body(entry_id):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    item = outlook.GetItemFromID(entry_id)
    return item.Body
```

#### 获取附件文件
由于附件内容不直接包含在 JSON 中，需要通过 `SourceMailEntryId` 找到原邮件并下载附件。

```python
import os

def save_attachments(context, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # 注意：如果是全局模式，context 是一个列表，需要双重循环
    # 这里假设 context 是单个事件对象
    for att_info in context.get("Attachments", []):
        entry_id = att_info.get("SourceMailEntryId")
        file_name = att_info.get("FileName")
        
        try:
            mail = outlook.GetItemFromID(entry_id)
            for attachment in mail.Attachments:
                if attachment.FileName == file_name:
                    # SaveAsFile 需要绝对路径
                    save_path = os.path.abspath(os.path.join(output_dir, file_name))
                    attachment.SaveAsFile(save_path)
                    break
        except Exception as e:
            # 建议将日志打印到 stderr 以免干扰 stdout 的 JSON 输出
            sys.stderr.write(f"Error saving {file_name}: {e}\n")
```

### 4.3 输出数据 (Standard Output)
脚本应将处理结果以 **JSON 格式** 打印到标准输出 (stdout)。
*建议*: 仅在最后一行打印 JSON 结果，其他日志信息打印到标准错误 (stderr)。

**推荐的输出结构**:
插件应设计为接收以下格式的更新指令（需插件端代码配合支持）：

```json
{
  "updates": {
    "DashboardValues": {
      "ETD": "2023-11-01",
      "Vessel": "EVER GIVEN"
    }
  },
  "logs": [
    "Extracted ETD successfully",
    "Confidence score: 0.98"
  ]
}
```

### 4.4 退出代码 (Exit Codes)
- `0`: 执行成功。
- `非0`: 执行失败（插件可能会捕获此状态并报错）。

## 5. 开发注意事项

1. **编码**: 始终使用 UTF-8 编码读写文件和处理字符串，以避免中文乱码。
2. **依赖**: 尽量减少第三方库依赖，或者确保运行环境已安装所需库（如 `pywin32`, `ollama` 等）。
3. **异常处理**: 脚本应包含全局异常捕获，将错误信息打印到 stderr 或包含在返回的 JSON 错误字段中，防止静默失败。
