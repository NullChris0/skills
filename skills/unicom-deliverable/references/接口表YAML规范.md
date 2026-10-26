# 接口表 YAML 规范

接口表规范属于本 Skill，不依赖任何旧接口表目录。一个接口、一个只读契约或一张物理表各用一个 YAML 文件。下面的 YAML 是形状示例。

## 公共字段

```yaml
id: item.list
kind: http
title: 查询记录
nesma: TF-01
section: "7.1"
table: "7-1"
caption: 查询记录接口
```

- `id`：全局唯一，建议使用 `domain.action`；
- `kind`：`http`、`elf` 或 `db`；
- `nesma`：阶段一编号，可为字符串或列表；
- `section`、`table`、`caption`：Word 中的归属位置和图表编号。

未独立计数的接口写 `nesma: uncounted`，不得伪造编号或计入 UFP。

## HTTP

```yaml
id: item.list
kind: http
title: 查询记录
method: GET
path: /items
controller: ItemController#list
auth: 管理端
request:
  - name: name
    type: string
    required: false
response:
  - name: rows
    type: array
rules:
  - 按名称和状态筛选。
nesma: TF-01
section: "7.1"
table: "7-1"
caption: 查询记录接口
```

HTTP 接口必须与固定代码基线中的路径、方法、DTO/VO 和鉴权一致。`controller` 是代码指针，生成卡片时不印出，也不写入调用规则。`auth` 是卡片上原样印出的鉴权文案，不是固定角色词表。`rules` 是调用规则，只写业务句子。只描述成功响应和业务字段；不复制旧交付件中的错误码、示例路径或未实现字段。

## ELF

```yaml
id: master.party
kind: elf
title: 外部主体
nesma: DF-01
source: 外部应用
direction: 外部应用 -> 本应用
fields:
  - partyId
  - partyName
section: "7.2"
table: "7-2"
caption: 外部主体只读契约
```

ELF 只描述本应用持续引用、由边界外应用维护的数据。一次性导入文件不是 ELF。

## DB

```yaml
id: db.item
kind: db
title: 记录物理表
tableName: item
columns:
  - name: item_name
    type: varchar
    nullable: false
purpose: 承载记录基本信息
section: "7.1"
table: "7.1-1"
caption: 记录物理表
```

物理表用于补充业务模型，不按表数或列数推导 NESMA。技术表可以描述，但只有阶段一确认的数据功能才引用 NESMA 编号。

## 投影与校验

```text
YAML -> TeX 卡片 -> PNG/SVG -> Word
```

- YAML 是编辑面；生成的 TeX 卡片和图片不得手改。
- 卡片样式默认使用本 Skill 的 `assets/api-card.sty`。该文件是版式范本：鉴权文案由 YAML 供给；成功信封和宏前缀属于范本，换项目时复制后改写，并用交付模型的 `card_style` 指向改写文件。
- 每个 YAML 必须生成至少一张图片；每个 `nesma` 引用必须在阶段一 JSON 中存在。
- HTTP 端点集合必须与固定代码基线一致；辅助 ELF/DB 卡片单独校验。
- `nesma` 仅用于内部追溯和映射表；ELF/DB 数据模型卡片不把 `DF-xx` 印到图片上。
- Word 中每张图片必须有图题、表号和正文引用。
- 调用规则等于 YAML 的 `rules`，不追加类名、方法名或文件路径。
