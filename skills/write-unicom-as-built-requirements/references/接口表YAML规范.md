# 接口表 YAML 规范

接口表规范属于本 Skill，不依赖任何后端 `api-tables` 目录。一个接口、一个只读契约或一张物理表各用一个 YAML 文件。

## 公共字段

```yaml
id: template.list
kind: http
title: 查询考核模板
nesma: TF-08
section: 8.1
table: 8-1
caption: 查询考核模板接口
```

- `id`：全局唯一，建议使用 `domain.action`；
- `kind`：`http`、`elf` 或 `db`；
- `nesma`：阶段一编号，可为字符串或列表；
- `section`、`table`、`caption`：Word 中的归属位置和图表编号。

未独立计数的接口写 `nesma: uncounted`，不得伪造编号或计入 UFP。

## HTTP

```yaml
id: template.list
kind: http
title: 查询考核模板
method: GET
path: /ea/template/templates
controller: TemplateController#list
auth: admin-or-ea-admin
request:
  - name: templateName
    type: string
    required: false
response:
  - name: rows
    type: array
rules:
  - 按模板名称和状态筛选。
nesma: TF-08
section: 8.1
table: 8-1
caption: 查询考核模板接口
```

HTTP 接口必须与固定代码基线中的控制器路径、方法、DTO/VO 和鉴权一致。只描述成功响应和业务字段；不复制旧交付件中的错误码、示例路径或未实现字段。

## ELF

```yaml
id: master-data.employee
kind: elf
title: 主系统人员
nesma: DF-12
source: 主系统
direction: 主系统 -> 本应用
fields:
  - employeeId
  - employeeName
section: 8.12
table: 8-46
caption: 主系统人员只读契约
```

ELF 只描述本应用持续引用、由边界外应用维护的数据。一次性导入文件不是 ELF。

## DB

```yaml
id: db.ea-template
kind: db
title: 考核模板物理表
tableName: ea_template
columns:
  - name: template_name
    type: varchar
    nullable: false
purpose: 承载模板基本信息
section: 7.1
table: 7.1-1
caption: 考核模板物理表
```

物理表用于补充业务模型，不按表数或列数推导 NESMA。技术表可以描述，但只有阶段一确认的数据功能才引用 NESMA 编号。

## 投影与校验

```text
YAML -> TeX 卡片 -> PNG/SVG -> Word
```

- YAML 是编辑面；生成的 TeX 卡片和图片不得手改。
- 卡片样式使用本 Skill 的 `assets/api-card.sty`，不引用旧仓库路径。
- 每个 YAML 必须生成至少一张图片；每个 `nesma` 引用必须在阶段一 JSON 中存在。
- HTTP 端点集合必须与固定代码基线的控制器一致；辅助 ELF/DB 卡片单独校验。
- Word 中每张图片必须有图题、表号和正文引用。
