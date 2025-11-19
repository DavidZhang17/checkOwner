# 查业主数据导出工具 - 配置说明

## 📁 文件结构

```
查业主/
├── convertToExcelFixed.js  # 主程序
├── queries.json            # 查询条件配置文件
├── files/                  # Excel文件输出目录（自动创建）
│   └── *.xlsx             # 生成的Excel文件
└── README_配置说明.md      # 本说明文件
```

## 🚀 使用方法

### 1. 配置查询条件

编辑 `queries.json` 文件，添加或修改查询条件：

```json
[
  {
    "fileName": "文件名称_不含xlsx后缀",
    "query": {
      "pageNum": 1, // 起始页码（断点续传时修改此值）
      "pageSize": 5000 // 每页大小（建议1000-5000）
      // ... 其他查询条件
    }
  }
]
```

### 2. 运行程序

```bash
node convertToExcelFixed.js
```

### 3. 查看结果

生成的 Excel 文件将保存在 `files/` 文件夹中。

## ⚙️ 重要配置说明

### 查询条件字段说明

- `fileName`: Excel 文件名（不含.xlsx 后缀）
- `pageNum`: 起始页码
  - 首次运行：设置为 `1`
  - 断点续传：设置为中断的下一页
- `pageSize`: 每页记录数
  - 建议值：1000-5000
  - **每个查询最多导出 5000 条记录**

### 常用查询条件

- `changeTimeMax` / `changeTimeMin`: 变更时间范围
- `endAge`: 最大年龄
- `registStatus`: 注册状态（2=有效）
- `arrayCertTypes`: 证书类型
- `totalAchieveCount`: 最少业绩数量

## 📝 多查询配置示例

```json
[
  {
    "fileName": "一级建造师_公路工程_2025Q1",
    "query": {
      "pageNum": 1,
      "pageSize": 5000,
      "changeTimeMax": "2025-03-31",
      "changeTimeMin": "2025-01-01",
      "arrayCertTypes": [{ "code": "961", "type": 1 }]
      // ... 其他条件
    }
  },
  {
    "fileName": "一级建造师_市政工程_2025Q1",
    "query": {
      "pageNum": 1,
      "pageSize": 5000,
      "changeTimeMax": "2025-03-31",
      "changeTimeMin": "2025-01-01",
      "arrayCertTypes": [{ "code": "962", "type": 1 }]
      // ... 其他条件
    }
  }
]
```

## 🔄 断点续传

如果程序中途中断，修改对应查询的 `pageNum` 继续运行：

```json
{
  "fileName": "一级建造师_公路工程_2025Q1",
  "query": {
    "pageNum": 3, // 从第3页开始继续
    "pageSize": 5000
    // ... 其他条件
  }
}
```

## ⚠️ 注意事项

1. **记录数限制**：每个查询最多导出 5000 条记录
2. **文件覆盖**：同名文件会被覆盖（断点续传除外）
3. **Token 有效期**：需要定期更新 `convertToExcelFixed.js` 中的 `authorization` token
4. **查询间隔**：多个查询之间会自动休息 5 秒，避免请求过快

## 📊 Excel 文件内容

每个 Excel 文件包含以下列：

| 列名     | 说明                               |
| -------- | ---------------------------------- |
| 姓名     | 人员姓名                           |
| 年龄     | 年龄（从身份证计算）               |
| 身份证号 | 身份证号码                         |
| 业绩     | A/B/C/D 级业绩统计                 |
| 专业     | 证书信息（时间、状态、公司、专业） |
| 项目名称 | 最近 5 个项目名称                  |

## 🛠️ 故障排除

### 配置文件加载失败

- 检查 `queries.json` 格式是否正确
- 确保 JSON 语法正确（逗号、引号等）

### Token 失效

- 更新 `convertToExcelFixed.js` 第 9 行的 `authorization` 值

### 文件夹权限问题

- 确保程序有权限在当前目录创建 `files` 文件夹

## 📞 技术支持

如有问题，请检查控制台输出的错误信息。
