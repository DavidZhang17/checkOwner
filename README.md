# JSON 转 Excel 工具

这是一个用于将 JSON 数据转换为 Excel 文件的 Node.js 工具，特别适用于处理复杂的嵌套 JSON 数据结构。

## 功能特点

- ✅ 支持复杂嵌套的 JSON 数据结构
- ✅ 智能展平对象和数组
- ✅ 自动处理空值和 null 值
- ✅ 创建主数据工作表和统计信息工作表
- ✅ 自动调整 Excel 列宽
- ✅ 支持命令行参数
- ✅ 详细的转换日志

## 安装依赖

首先确保你已经安装了 Node.js，然后安装必要的依赖：

```bash
npm install
```

## 使用方法

### 基本用法

```bash
# 使用默认的data.json文件，输出为data_converted.xlsx
npm start
```

或者

```bash
node convertToExcel.js
```

### 指定输入和输出文件

```bash
# 指定输入文件
node convertToExcel.js your-data.json

# 指定输入和输出文件
node convertToExcel.js your-data.json output.xlsx
```

## 数据处理规则

### 1. 基本字段

所有基本类型字段（字符串、数字、布尔值）会直接保留。

### 2. 数组处理

- **基本类型数组**: 用分号(`;`)连接所有值
- **对象数组**: 展平第一个对象的所有字段，如果有多个对象：
  - 添加`_count`字段显示数组长度
  - 合并关键字段（如证书号、到期日期等）到`_all_字段名`

### 3. 嵌套对象

所有嵌套对象会被展平，字段名用下划线连接，如：

- `company.name` → `company_name`
- `safety[0].certificateNumber` → `safety_certificateNumber`

### 4. 空值处理

`null`、`undefined`或空值会转换为空字符串。

## 输出文件结构

生成的 Excel 文件包含两个工作表：

### 1. 人员数据工作表

- 包含所有处理后的人员数据
- 自动添加行号
- 自动调整列宽

### 2. 统计信息工作表

- 总记录数统计
- 各类数据完整性统计
- 按地区分布的 TOP10 统计

## 快速开始示例

```bash
# 1. 确保依赖已安装
npm install

# 2. 运行转换（使用data.json）
npm start

# 3. 查看生成的data_converted.xlsx文件
```

## 高级用法

```bash
# 处理自定义JSON文件
node convertToExcel.js my-data.json

# 指定输出文件名
node convertToExcel.js data.json 人员信息表.xlsx
```

## 错误排除

如果遇到"Cannot find module"错误：

1. 确保在正确的目录下运行命令
2. 确保已经运行了 `npm install`
3. 检查`convertToExcel.js`文件是否存在

## 技术栈

- **Node.js**: JavaScript 运行环境
- **xlsx**: Excel 文件处理库

## 许可证

MIT License
