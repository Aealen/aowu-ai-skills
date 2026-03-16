---
name: dao-unit-test-guide
description: Use when generating DAO layer unit tests for MyBatis-based projects, especially when Mapper XML SQL definitions must exist for methods to work. Use when creating test methods for add/getById/update/delById operations, handling getBySqlKey vs getOne method selection, or debugging "Mapped Statements collection does not contain value" errors.
---

# DAO 单元测试生成指南

## 核心原则

> **所有SQL必须在Mapper.xml中定义才能执行**
>
> 即使Dao继承了BaseDao的预定义方法（add、getById、delById），如果Mapper.xml中没有对应SQL，方法也无法执行。

**调用链路**：Service → Dao → BaseDao/GenericDao → Mapper.xml

## 生成前必做检查

1. **SQL存在性**：查看Mapper.xml确认SQL语句是否存在
2. **Dao方法存在性**：读取Dao类源代码，很多Dao是空的只继承BaseDao
3. **调用链路**：`grep -rn "sqlKey" --include="*Dao.java" --include="*Service.java" ./`

## 类结构模板

```java
@Slf4j
@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(locations = {"classpath:app-context.xml"})
@Transactional
@Rollback
@Component
public class XxxDaoTest {
    @Resource
    private XxxDao xxxDao;
}
```

## BaseDao 方法选择规则（核心）

| SQL类型 | 返回类型 | 正确方法 | 错误方法 |
|--------|---------|---------|---------|
| INSERT | `int` | `dao.add(entity)` 或 `dao.add("sqlKey", param)` | `dao.getBySqlKey("insert", entity)` |
| SELECT单个 | `Object` | `dao.getOne("sqlKey", params)` 需强转 | `dao.getBySqlKey("getById", id)` |
| SELECT List | `List<E>` | `dao.getBySqlKey("sqlKey", params)` | - |
| SELECT分页 | `PageBean<E>` | `dao.getBySqlKey("sqlKey", queryFilter)` | 用`List<E>`接收 |
| UPDATE | `int` | `dao.update(entity)` 或 `dao.update("sqlKey", param)` | - |
| DELETE | `int` | `dao.delBySqlKey("sqlKey", param)` | - |
| 逻辑删除 | `int` | `dao.update("sqlKey", param)` (实际是UPDATE) | `dao.delBySqlKey()` |

### GenericDao内置方法优先

| Mapper.xml SQL ID | 正确调用 |
|------------------|---------|
| `add` | `dao.add(entity)` |
| `getById` | `dao.getById(id)` |
| `update` | `dao.update(entity)` |
| `delById` | `dao.delById(id)` |

### 自定义SQL ID调用

| SQL ID | 正确调用 |
|--------|---------|
| `insert` | `dao.add("insert", entity)` |
| `selectById` | `(Entity)dao.getOne("selectById", id)` |
| `updateById` | `dao.update("updateById", entity)` |
| `deleteById` | `dao.delBySqlKey("deleteById", id)` |

## 造数据规则

| 字段类型 | Mock方式 |
|---------|---------|
| 主键 | `UUID.randomUUID().toString()` |
| 公司ID | `"336666"` |
| 编号字段(xxxCode) | `"B001"`（6字符以内） |
| 状态(int) | `entity.setStatus(0)` |
| 状态(String) | `entity.setBizStatus("1")` |
| BigDecimal | `new BigDecimal("10000")` |

## 常见错误

### 错误1：INSERT用getBySqlKey
```java
// ❌ getBySqlKey返回List<E>，不是int
int result = dao.getBySqlKey("insert", entity);
// ✅ 使用add方法
int result = dao.add("insert", entity);
```

### 错误2：类型不匹配
```java
// ❌ status是int类型
entity.setStatus("0");
assertEquals(Integer.valueOf(1), entity.getStatus());
// ✅
entity.setStatus(0);
assertEquals(1, entity.getStatus());
```

### 错误3：使用var
```java
// ❌ Java 8不支持var
var list = dao.getBySqlKey("getDataList", params);
// ✅
List<XxxEntity> list = dao.getBySqlKey("getDataList", params);
```

## 特殊情况

### 无调用链路
```java
@Test
public void testCustomQuery() {
    log.warn("无调用链路！SQL ID: customQuery");
    // 仍需编写完整测试逻辑
}
```

### Lombok @Data的父类字段问题
使用`@Data`注解的实体类，**只会为当前类生成getter/setter**，不包含父类字段。

### Mapper XML bind标签
必须为所有`<bind>`涉及的字段提供值，否则NPE。

### 删除操作判断
根据Mapper.xml中**实际SQL内容**选择方法：
- `delete from ...` → `dao.delBySqlKey()`
- `update ... set status = 1` → `dao.update()`

## 检查清单

- [ ] Mapper.xml中存在对应SQL ID
- [ ] Dao类中有对应方法，或使用BaseDao通用方法
- [ ] INSERT使用add方法
- [ ] SELECT单个对象使用getOne并强转
- [ ] QueryFilter参数返回PageBean（不是List）
- [ ] 参数类型与字段类型一致
- [ ] 未使用var关键字
- [ ] 实体类中存在使用的setter/getter
- [ ] 逻辑删除使用update方法
