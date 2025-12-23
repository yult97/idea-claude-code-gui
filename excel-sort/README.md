# Excel需求排序工具

一个智能化的Excel需求管理工具，能够根据二级模块的顺序自动调整需求名称的排列顺序，提高需求管理效率。

## 核心功能

- **智能排序**: 根据二级模块的顺序，自动调整需求名称的排列顺序
- **一一对应**: 确保每个需求名称与其对应的二级模块保持一致
- **异常标注**: 自动识别无法匹配的需求项，并在结果中明确标注
- **Excel导入导出**: 支持Excel文件的上传和处理结果的导出

## 业务价值

- **效率提升**: 将手动排序的小时级工作缩短到分钟级
- **准确性保证**: 消除人工操作的错误风险
- **可追溯性**: 自动标注异常项，便于后续处理

## 技术架构

### 前端技术栈
- React 18+ - 现代化组件开发框架
- Ant Design 5.x - 企业级UI组件库
- Axios - HTTP请求库

### 后端技术栈
- Python 3.9+ - 简洁高效的数据处理语言
- Flask 2.x - 轻量级Web框架
- Pandas - 强大的数据处理库
- OpenPyXL - Excel文件读写库
- difflib - 字符串相似度计算

## 快速开始

### 环境要求
- Python 3.9+
- Node.js 14+
- npm 或 yarn

### 安装与运行

1. **克隆项目**
```bash
git clone <repository-url>
cd excel-sort
```

2. **一键启动**
```bash
./run.sh
```

3. **手动启动**

后端服务：
```bash
cd backend
python3 -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

前端服务：
```bash
cd frontend
npm install
npm start
```

### 使用说明

1. 打开浏览器访问 `http://localhost:3000`
2. 点击或拖拽上传Excel文件
3. 点击"开始处理"按钮
4. 等待处理完成，自动下载结果文件
5. 查看处理结果统计和详情

## Excel文件格式要求

- **A列**: 二级模块（参考顺序）
- **D列**: 需求名称（待排序）
- 支持格式：`.xlsx`, `.xls`
- 文件大小：不超过50MB

## 匹配算法

### 精确匹配
- 模块名完全包含在需求名中
- 相似度 ≥ 80%

### 模糊匹配
- 基于字符串相似度计算
- 相似度阈值 ≥ 30%
- 使用difflib.SequenceMatcher算法

### 异常处理
- 无法匹配的需求标记为"未匹配"
- 在结果文件中添加"匹配状态"和"匹配模块"列
- 提供详细的处理报告

## API接口

### POST /api/sort-excel
处理Excel文件排序

**请求参数:**
- `file`: Excel文件（multipart/form-data）

**响应:**
- 处理后的Excel文件下载

### GET /api/process-result
获取最近一次处理的结果统计

**响应:**
```json
{
  "totalRequirements": 100,
  "matchedRequirements": 95,
  "unmatchedRequirements": 5,
  "matchRate": 95.0,
  "details": [...],
  "warnings": [...]
}
```

### GET /api/health
健康检查接口

## 项目结构

```
excel-sort/
├── frontend/                 # React前端
│   ├── public/
│   ├── src/
│   │   ├── App.js           # 主应用组件
│   │   ├── App.css          # 样式文件
│   │   └── index.js         # 入口文件
│   └── package.json
├── backend/                  # Python后端
│   ├── app.py               # Flask应用主文件
│   └── requirements.txt     # Python依赖
├── run.sh                   # 一键启动脚本
└── README.md               # 项目文档
```

## 开发说明

### 前端开发
```bash
cd frontend
npm start          # 开发模式
npm run build      # 生产构建
npm test           # 运行测试
```

### 后端开发
```bash
cd backend
source venv/bin/activate
python app.py      # 开发模式
```

## 注意事项

1. 确保Excel文件格式正确，A列为二级模块，D列为需求名称
2. 文件大小不超过50MB
3. 处理大文件时请耐心等待
4. 建议在处理前备份原始文件

## 故障排除

### 常见问题

1. **文件上传失败**
   - 检查文件格式是否为.xlsx或.xls
   - 检查文件大小是否超过50MB

2. **处理失败**
   - 检查Excel文件结构是否正确
   - 确保A列和D列数据完整

3. **服务启动失败**
   - 检查Python和Node.js版本
   - 确保端口5000和3000未被占用

## 贡献指南

欢迎提交Issue和Pull Request来改进这个项目。

## 许可证

MIT License