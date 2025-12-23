import React, { useState } from 'react';
import {
  Layout,
  Typography,
  Card,
  Upload,
  Button,
  message,
  Table,
  Space,
  Statistic,
  Row,
  Col,
  Alert,
  Tag,
  Progress,
  Spin,
  Pagination
} from 'antd';
import {
  InboxOutlined,
  DownloadOutlined,
  FileExcelOutlined,
  CheckCircleOutlined,
  ExclamationCircleOutlined,
  InfoCircleOutlined
} from '@ant-design/icons';
import axios from 'axios';
import './App.css';

// 未匹配内容展示组件
const UnmatchedContentDisplay = ({ items, title, type }) => {
  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 10;

  const paginatedItems = items.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  const columns = [
    {
      title: '序号',
      dataIndex: 'index',
      key: 'index',
      width: 60,
      render: (text, record, index) => (currentPage - 1) * pageSize + index + 1
    },
    {
      title: type === 'module' ? '模块名称' : '需求名称',
      dataIndex: 'content',
      key: 'content',
      render: (text) => (
        <code style={{
          backgroundColor: '#f6f8fa',
          padding: '2px 6px',
          borderRadius: '4px',
          fontSize: '12px'
        }}>
          {text}
        </code>
      )
    }
  ];

  const dataSource = paginatedItems.map((item, index) => ({
    key: index,
    content: item
  }));

  return (
    <div>
      <Table
        columns={columns}
        dataSource={dataSource}
        pagination={false}
        size="small"
        scroll={{ y: 300 }}
        locale={{ emptyText: `暂无${title}` }}
      />
      {items.length > pageSize && (
        <div style={{ textAlign: 'center', marginTop: '16px' }}>
          <Pagination
            current={currentPage}
            total={items.length}
            pageSize={pageSize}
            showSizeChanger={false}
            showQuickJumper
            showTotal={(total, range) =>
              `第 ${range[0]}-${range[1]} 条，共 ${total} 条${title}`
            }
            onChange={setCurrentPage}
          />
        </div>
      )}
    </div>
  );
};

const { Header, Content } = Layout;
const { Title, Text, Paragraph } = Typography;
const { Dragger } = Upload;

function App() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [processing, setProcessing] = useState(false);

  const uploadProps = {
    name: 'file',
    multiple: false,
    accept: '.xlsx,.xls',
    beforeUpload: (file) => {
      const isExcel = file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                     file.type === 'application/vnd.ms-excel';
      if (!isExcel) {
        message.error('请上传Excel文件!');
        return false;
      }
      const isLt50M = file.size / 1024 / 1024 < 50;
      if (!isLt50M) {
        message.error('文件大小不能超过50MB!');
        return false;
      }
      setFile(file);
      return false; // 阻止自动上传
    },
    onRemove: () => {
      setFile(null);
      setResult(null);
    }
  };

  const handleProcess = async () => {
    if (!file) {
      message.error('请先选择文件');
      return;
    }

    const formData = new FormData();
    formData.append('file', file);

    setLoading(true);
    setProcessing(true);
    setResult(null);

    try {
      const response = await axios.post('/api/sort-excel', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
        responseType: 'blob'
      });

      // 检查响应状态
      if (!response || response.status !== 200) {
        throw new Error('服务器响应异常');
      }

      // 检查响应数据是否为JSON格式（错误信息）
      const contentType = response.headers['content-type'];
      if (contentType && contentType.includes('application/json')) {
        // 如果是JSON，说明返回的是错误信息
        const reader = new FileReader();
        reader.onload = () => {
          try {
            const errorData = JSON.parse(reader.result);
            message.error(errorData.error || '文件处理失败');
          } catch (e) {
            message.error('文件处理失败');
          }
        };
        reader.readAsText(response.data);
        return;
      }

      // 创建下载链接
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `sorted_${file.name}`);
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);

      message.success('文件处理完成！');

      // 获取处理结果信息
      try {
        const resultResponse = await axios.get('/api/process-result');
        setResult(resultResponse.data);
      } catch (resultError) {
        console.warn('无法获取处理结果:', resultError);
        // 不影响文件下载，只记录警告
      }

    } catch (error) {
      console.error('处理失败:', error);

      // 处理不同类型的错误
      if (error.response) {
        // 服务器返回了错误状态码
        if (error.response.data instanceof Blob) {
          // 如果错误信息是Blob格式，读取其内容
          const reader = new FileReader();
          reader.onload = () => {
            try {
              const errorData = JSON.parse(reader.result);
              message.error(errorData.error || '文件处理失败');
            } catch (e) {
              message.error('文件处理失败');
            }
          };
          reader.readAsText(error.response.data);
        } else {
          message.error(error.response.data?.error || '文件处理失败');
        }
      } else if (error.request) {
        // 网络错误
        message.error('网络连接失败，请检查后端服务是否正常启动');
      } else {
        // 其他错误
        message.error(`处理失败: ${error.message}`);
      }
    } finally {
      setLoading(false);
      setProcessing(false);
    }
  };

  const resultColumns = [
    {
      title: '状态',
      dataIndex: 'status',
      key: 'status',
      render: (status) => {
        const statusConfig = {
          success: { color: 'success', icon: <CheckCircleOutlined />, text: '成功' },
          warning: { color: 'warning', icon: <ExclamationCircleOutlined />, text: '警告' },
          info: { color: 'default', icon: <InfoCircleOutlined />, text: '信息' }
        };
        const config = statusConfig[status] || statusConfig.info;
        return (
          <Tag color={config.color} icon={config.icon}>
            {config.text}
          </Tag>
        );
      }
    },
    {
      title: '描述',
      dataIndex: 'message',
      key: 'message'
    },
    {
      title: '数量',
      dataIndex: 'count',
      key: 'count',
      render: (count) => count || '-'
    }
  ];

  return (
    <Layout className="min-h-screen">
      <Header style={{ background: '#fff', boxShadow: '0 2px 8px rgba(0,0,0,0.1)' }}>
        <div style={{ display: 'flex', alignItems: 'center', height: '100%' }}>
          <FileExcelOutlined style={{ fontSize: '24px', color: '#1890ff', marginRight: '12px' }} />
          <Title level={3} style={{ margin: 0, color: '#1890ff' }}>Excel需求排序工具</Title>
        </div>
      </Header>

      <Content style={{ padding: '24px', background: '#f0f2f5', minHeight: 'calc(100vh - 64px)' }}>
        <div className="container">
          <Card>
            <div style={{ marginBottom: '24px' }}>
              <Title level={4}>Excel数据智能匹配工具</Title>
              <Paragraph>
                本工具基于精确匹配算法，处理Excel文件中"二级模块"和"需求名称"列的数据对应关系。
                以二级模块列为基准，在需求名称列中查找完全相同的文本进行匹配。
              </Paragraph>

            <Card title="🎯 匹配规则说明" size="small" style={{ marginBottom: '16px', backgroundColor: '#f6ffed' }}>
              <Paragraph style={{ marginBottom: '8px' }}>
                <strong>✅ 精确匹配（成功）：</strong>
              </Paragraph>
              <ul style={{ marginBottom: '16px' }}>
                <li>二级模块：<code>"应急演练公告"</code>，需求名称：<code>"应急演练公告"</code> → 匹配成功</li>
                <li>二级模块：<code>"皮肤管理"</code>，需求名称：<code>"皮肤管理"</code> → 匹配成功</li>
                <li>自动去除首尾空格：<code>"  应急演练公告  "</code> = <code>"应急演练公告"</code></li>
              </ul>

              <Paragraph style={{ marginBottom: '8px' }}>
                <strong>❌ 匹配失败（不匹配）：</strong>
              </Paragraph>
              <ul style={{ marginBottom: '0' }}>
                <li>二级模块：<code>"应急演练公告"</code>，需求名称：<code>"应急演练公告v2.0"</code> → 不匹配</li>
                <li>二级模块：<code>"皮肤管理"</code>，需求名称：<code>"皮肤管理模块"</code> → 不匹配</li>
                <li>文本必须<strong>完全相同</strong>，不能多字、少字或顺序不同</li>
              </ul>
            </Card>

            <Paragraph>
              <strong>📋 输入要求：</strong>A列为二级模块，D列为需求名称，其他列将被忽略
            </Paragraph>
            <Paragraph>
              <strong>📄 输出格式：</strong>生成包含4列的新Excel文件，移除所有"Unnamed"列
            </Paragraph>
            </div>

            <Card title="文件上传" style={{ marginBottom: '24px' }}>
              <Dragger {...uploadProps}>
                <p className="ant-upload-drag-icon">
                  <InboxOutlined />
                </p>
                <p className="ant-upload-text">点击或拖拽文件到此区域上传</p>
                <p className="ant-upload-hint">
                  支持单个Excel文件上传(.xlsx, .xls)，文件大小不超过50MB
                </p>
              </Dragger>

              {file && (
                <div style={{ marginTop: '16px', textAlign: 'center' }}>
                  <Space>
                    <Text>已选择文件: <strong>{file.name}</strong></Text>
                    <Button
                      type="primary"
                      onClick={handleProcess}
                      loading={loading}
                      icon={<FileExcelOutlined />}
                    >
                      开始处理
                    </Button>
                  </Space>
                </div>
              )}
            </Card>

            {processing && (
              <Card style={{ textAlign: 'center', marginBottom: '24px' }}>
                <Spin size="large" />
                <div style={{ marginTop: '16px' }}>
                  <Text>正在处理文件，请稍候...</Text>
                  <Progress percent={100} status="active" showInfo={false} style={{ marginTop: '8px' }} />
                </div>
              </Card>
            )}

            {result && (
              <div className="result-area">
                <Row gutter={16} style={{ marginBottom: '24px' }}>
                  <Col span={6}>
                    <Card>
                      <Statistic
                        title="总需求数"
                        value={result.totalRequirements}
                        valueStyle={{ color: '#1890ff' }}
                      />
                    </Card>
                  </Col>
                  <Col span={6}>
                    <Card>
                      <Statistic
                        title="成功匹配"
                        value={result.matchedRequirements}
                        valueStyle={{ color: '#52c41a' }}
                      />
                    </Card>
                  </Col>
                  <Col span={6}>
                    <Card>
                      <Statistic
                        title="未匹配"
                        value={result.unmatchedRequirements.length}
                        valueStyle={{ color: '#faad14' }}
                      />
                    </Card>
                  </Col>
                  <Col span={6}>
                    <Card>
                      <Statistic
                        title="匹配率"
                        value={result.matchRate}
                        suffix="%"
                        valueStyle={{ color: '#52c41a' }}
                      />
                    </Card>
                  </Col>
                </Row>

                <Card title="处理结果详情">
                  <Table
                    columns={resultColumns}
                    dataSource={result.details}
                    pagination={false}
                    size="small"
                  />
                </Card>

                {/* 未匹配内容展示区域 */}
                {(result.unmatchedModules && result.unmatchedModules.length > 0) && (
                  <Card title="📋 未匹配的模块" style={{ marginTop: '16px' }}>
                    <UnmatchedContentDisplay
                      items={result.unmatchedModules}
                      title="未匹配模块"
                      type="module"
                    />
                  </Card>
                )}

                {(result.unmatchedRequirements && result.unmatchedRequirements.length > 0) && (
                  <Card title="📋 未匹配的需求名称" style={{ marginTop: '16px' }}>
                    <UnmatchedContentDisplay
                      items={result.unmatchedRequirements}
                      title="未匹配需求"
                      type="requirement"
                    />
                  </Card>
                )}

                <div style={{ textAlign: 'center', marginTop: '24px' }}>
                  <Button
                    type="primary"
                    size="large"
                    icon={<DownloadOutlined />}
                    onClick={() => {
                      const formData = new FormData();
                      formData.append('file', file);
                      axios.post('/api/sort-excel', formData, {
                        headers: { 'Content-Type': 'multipart/form-data' },
                        responseType: 'blob'
                      }).then(response => {
                        const url = window.URL.createObjectURL(new Blob([response.data]));
                        const link = document.createElement('a');
                        link.href = url;
                        link.setAttribute('download', `sorted_${file.name}`);
                        document.body.appendChild(link);
                        link.click();
                        link.remove();
                        window.URL.revokeObjectURL(url);
                      });
                    }}
                  >
                    重新下载结果文件
                  </Button>
                </div>
              </div>
            )}
          </Card>
        </div>
      </Content>
    </Layout>
  );
}

export default App;