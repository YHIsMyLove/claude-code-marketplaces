# Mermaid 图表测试文档

这是一个包含 Mermaid 图表的测试文档，用于验证 Markdown 到 Word 的转换功能。

## 1. 流程图

下面是一个简单的流程图示例：

```mermaid
graph TD
    A[开始] --> B{是否有数据?}
    B -->|是| C[处理数据]
    B -->|否| D[等待数据]
    C --> E[保存结果]
    D --> B
    E --> F[结束]
```

## 2. 序列图

用户认证的序列图：

```mermaid
sequenceDiagram
    participant U as 用户
    participant C as 客户端
    participant S as 服务器

    U->>C: 输入用户名密码
    C->>S: 发送登录请求
    S->>S: 验证凭据
    S-->>C: 返回认证结果
    C-->>U: 显示登录状态
```

## 3. 饼图

数据统计分布图：

```mermaid
pie
    title 销售额分布
    "产品A" : 35
    "产品B" : 25
    "产品C" : 20
    "产品D" : 15
    "其他" : 5
```

## 4. 甘特图

项目时间线：

```mermaid
gantt
    title 项目开发计划
    dateFormat  YYYY-MM-DD
    section 需求分析
    需求调研           :a1, 2024-01-01, 7d
    需求文档编写       :a2, after a1, 5d
    section 设计
    系统设计           :b1, after a2, 10d
    数据库设计         :b2, after b1, 5d
    section 开发
    前端开发           :c1, after b2, 15d
    后端开发           :c2, after b2, 20d
    section 测试
    集成测试           :d1, after c1, 10d
    用户验收测试       :d2, after c2, 5d
```

## 5. 类图

系统架构类图：

```mermaid
classDiagram
    class User {
        +String username
        +String email
        +login()
        +logout()
    }

    class Order {
        +String orderId
        +Date createTime
        +Float amount
        +create()
        +cancel()
    }

    class Product {
        +String productId
        +String name
        +Float price
        +getDetails()
    }

    User "1" --> "0..*" Order : places
    Order "1" --> "1..*" Product : contains
```

## 结论

这个文档包含了多种类型的 Mermaid 图表，用于测试转换工具的兼容性。如果所有图表都能正确转换为 Word 文档中的图片，说明转换工具工作正常。