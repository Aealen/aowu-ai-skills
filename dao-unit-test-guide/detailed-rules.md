# DAO 单元测试详细规则

## 一、三者关联关系

| 组件 | 位置 | 关联标识 |
|------|------|----------|
| Mapper.xml | `src/main/resources/.../*Mapper.xml` | `namespace="com.xxx.model.Entity"` |
| 实体类 | `*-api/.../model/Entity.java` | `package com.xxx.model; class Entity` |
| Dao类 | `*-svc/.../dao/XxxDao.java` | `extends BaseDao<Entity>` |

## 二、测试方法模板

### 插入测试
```java
@Test
public void testAdd() {
    XxxEntity entity = getEntity();
    int result = xxxDao.add(entity);
    assertEquals(1, result);
    XxxEntity fetched = xxxDao.getById(entity.getId());
    assertNotNull(fetched);
}
```

### 删除测试
```java
@Test
public void testDelById() {
    XxxEntity entity = getEntity();
    xxxDao.add(entity);
    int result = xxxDao.delById(entity.getId());
    assertEquals(1, result);
    assertNull(xxxDao.getById(entity.getId()));
}
```

### 更新测试
```java
@Test
public void testUpdate() {
    XxxEntity entity = getEntity();
    xxxDao.add(entity);
    entity.setField1("UPDATED");
    int result = xxxDao.update(entity);
    assertEquals(1, result);
    assertEquals("UPDATED", xxxDao.getById(entity.getId()).getField1());
}
```

### 自定义查询测试
```java
@Test
public void testCustomQuery() {
    XxxEntity entity = getEntity();
    xxxDao.add(entity);
    Map<String, Object> params = new HashMap<>();
    params.put("id", entity.getId());
    List<XxxEntity> list = xxxDao.getBySqlKey("customQuery", params);
    assertNotNull(list);
    assertFalse(list.isEmpty());
}
```

## 三、getEntity() 方法模板

### 3.1 基础模板
```java
private XxxEntity getEntity() {
    XxxEntity entity = new XxxEntity();
    entity.setId(UUID.randomUUID().toString());
    entity.setCompanyId("336666");
    // 只设置实体类中存在的setter方法
    return entity;
}
```

### 3.2 实体类验证流程（生成测试前必做）

**步骤1：打开实体类文件**
```bash
# 实体类通常位于以下位置：
# - mkt-api/src/main/java/com/paas/mkt/model/
# - mkt-api/src/main/java/com/paas/mkt/model/vo/
```

**步骤2：查找所有字段定义**
```bash
# 方法1：查看所有 private 字段
grep -n "private.*\s\w\+;" XxxEntity.java

# 方法2：查看所有 setter 方法
grep -n "public void set" XxxEntity.java
```

**步骤3：创建字段清单**
```markdown
实体类： BusinessInformationVO
可用字段：
- businessId (String) → setBusinessId(), getBusinessId()
- businessCode (String) → setBusinessCode(), getBusinessCode()
- companyId (String) → setCompanyId(), getCompanyId()

不可用字段（不存在）：
- ❌ orgId
- ❌ pmId (实际是 principalId)
- ❌ createId (@Data不会生成)
- ❌ status (@Data不会生成)
```

**步骤4：根据字段清单编写 getEntity() 方法**

### 3.3 字段名推断错误表（重要）

| 推断的字段名 | 实际字段名 | 错误原因 |
|------------|-----------|----------|
| `setOrgId()` | ❌ 不存在 | 根据命名约定推断，实际 VO 中无此字段 |
| `setPmId()` | ❌ 不存在 | 实际是 `setPrincipalId()` |
| `setEstimateAmount()` | ❌ 不存在 | 实际是 `setEstimateInvest()` |
| `setProfession()` | ❌ 不存在 | 实际 VO 中无此字段 |
| `setCustomerId()` | ❌ 不存在 | 实际主键可能是其他字段 |
| `setBusinessId()` | ❌ 不存在 | BusinessProject 只有 resultId 和 projId |

**原则**：永远不要推断字段名，必须实际查看实体类定义！

### 3.4 大量参数转换技巧

当参数过多时，使用 `CommonStatic.transBean2Map()` 转换：

```java
// ❌ 错误：手动逐个设置参数
Map<String, Object> params = new HashMap<>();
params.put("field1", entity.getField1());
params.put("field2", entity.getField2());
// ... 太多参数

// ✅ 正确：使用工具方法批量转换
Map<String, Object> params = com.paas.model.fm.CommonStatic.transBean2Map(entity);
List<XxxEntity> list = xxxDao.getBySqlKey("getByCondition", params);
```

## 四、类型校验规则

### 4.1 Setter参数类型匹配
```java
// ❌ 错误：类型不匹配
entity.setStatus("0");        // status是int类型
entity.setBizStatus(1);       // bizStatus是String类型

// ✅ 正确
entity.setStatus(0);          // int类型
entity.setBizStatus("1");     // String类型
```

### 4.2 assertEquals参数类型匹配
```java
// BaseModel.status 是 int 类型
// ❌ 错误
assertEquals(Integer.valueOf(1), entity.getStatus());  // Integer vs int
// ✅ 正确
assertEquals(1, entity.getStatus());                    // int vs int
```

### 4.3 常见字段类型对照
| 字段名 | 通常类型 | 正确用法 |
|--------|---------|---------|
| `status`（BaseModel） | `int` | `setStatus(0)`、`assertEquals(1, getStatus())` |
| `bizStatus` | `String` | `setBizStatus("1")`、`assertEquals("1", getBizStatus())` |
| `createDate` | `Date` | `setCreateDate(new Date())` |
| `id` | `String` | `setId(UUID.randomUUID().toString())` |

### 4.4 BigDecimal比较
```java
// ❌ 错误
assertEquals(new BigDecimal("100"), entity.getAmount());
// ✅ 正确
assertTrue(entity.getAmount().compareTo(new BigDecimal("100")) == 0);
```

### 4.5 getOne返回值处理
```java
// getOne 返回 Object，需要类型转换
Integer count = (Integer) dao.getOne("isBusinessExistsModify", entity);
String value = (String) dao.getOne("getNameById", id);
Entity entity = (Entity) dao.getOne("getById", id);
```

### 4.6 非实体类参数对象检查（重要）

所有参数对象（VO、DTO、Config等）必须验证字段名称：

```java
// ❌ 错误：使用不存在的字段名
CheckRepeatConfig config = new CheckRepeatConfig();
config.setFieldName("businessCode");  // 不存在

// ✅ 正确：使用实际的字段名
config.setClassField("businessCode"); // 正确
```

**检查清单**：
- [ ] 查找参数对象类的源代码
- [ ] 确认字段名称和类型
- [ ] 验证使用的 setter/getter 方法是否存在
- [ ] 修复所有不存在的字段调用

## 五、特殊情况处理

### 5.1 已知SQL问题
```java
@Test
@Ignore("原sql报错：ORA-00933: SQL 命令未正确结束")
public void testAddBatch() { }
```

### 5.2 无调用链路
```java
@Test
public void testCustomQuery() {
    log.warn("无调用链路！SQL ID: customQuery 该方法未被业务代码调用");
    // 仍需编写完整测试逻辑
}
```

### 5.3 Lombok @Data 的父类字段问题

**重要**：使用`@Data`注解的实体类，**只会为当前类生成getter/setter，不会包含父类字段**。

```java
@Data
public class BusinessInformationVO extends BaseModel {
    private String businessId;  // 只有当前类的字段有getter/setter
}

// ❌ 错误：BaseModel字段不会自动生成
entity.setStatus(1);
entity.setCreateId("xxx");

// ✅ 正确：只使用当前类定义的字段
entity.setBusinessId("xxx");
```

### 5.4 Mapper XML bind标签参数完整性

```xml
<bind name="cstNameBlur" value="'%' + _parameter.get('cstName') + '%'" />
<bind name="cstCodeBlur" value="'%' + _parameter.get('cstCode') + '%'" />
```

```java
// ❌ 错误：cstCode为null导致NPE
params.put("cstName", entity.getCstName());
// params.put("cstCode", ...);  // 缺少

// ✅ 正确：为所有bind涉及字段提供值
params.put("cstName", entity.getCstName());
params.put("cstCode", "");      // 提供空字符串
```

### 5.5 父类方法入参类型

`getById`/`delById`入参类型为`String`，如果实体主键是`Integer`/`Long`，需要转换：

```java
entity.setId(new Random().nextInt(1000000) + 1000000);
Entity fetched = dao.getById(String.valueOf(entity.getId()));
```

### 5.6 getBySqlKey返回类型区分

| 方法签名 | 返回类型 |
|---------|---------|
| `getBySqlKey(sqlKey, params)` | `List<E>` |
| `getBySqlKey(sqlKey, QueryFilter)` | **`PageBean<E>`** |

```java
// ❌ 错误
List<Entity> list = dao.getBySqlKey("getAllPage", queryFilter);
// ✅ 正确
PageBean<Entity> pageBean = dao.getBySqlKey("getAllPage", queryFilter);
List<Entity> list = pageBean.getData();
```

### 5.7 删除操作判断规则

根据Mapper.xml中**实际SQL内容**选择方法：

| 实际SQL | 使用方法 |
|---------|---------|
| `delete from ...` | `dao.delBySqlKey("sqlId", param)` |
| `update ... set status = 1` | `dao.update("sqlId", param)` |

**注意**：标签可能是`<delete>`但实际SQL是`update`，需要查看SQL内容。

### 5.8 禁止使用var关键字

```java
// ❌ 错误：Java 8不支持var
var list = dao.getBySqlKey("getDataList", entity);
var entityList = new ArrayList<Entity>();

// ✅ 正确
List<Entity> list = dao.getBySqlKey("getDataList", entity);
List<Entity> entityList = new ArrayList<Entity>();
```

### 5.9 优先使用业务代码调用方式

```java
// 业务代码调用方式
List<ContactsInfo> data = contactsInfoDao.getBySqlKey("getAllForMainSel", map, pageBean);

// 测试代码应使用相同方式
List<ContactsInfo> data = contactsInfoDao.getBySqlKey("getAllForMainSel", params, pageBean);
```

### 5.10 getBySqlKey方法重载选择

| 场景 | 使用方法 |
|------|---------|
| SQL不需要参数 | `dao.getBySqlKey(sqlKey)` |
| SQL需要参数 | `dao.getBySqlKey(sqlKey, params)` |
| 分页查询 | `dao.getBySqlKey(sqlKey, queryFilter)` 返回`PageBean<E>` |

**避免使用null作为参数**，可能导致编译器选择错误的重载。

### 5.11 SQL ID与测试方法一一对应

```java
// Mapper.xml: <delete id="deleteById">
// ❌ 错误：使用未定义的delById
dao.delById(id);
// ✅ 正确：使用实际定义的deleteById
dao.delBySqlKey("deleteById", id);
```

## 六、BaseDao 方法详细规则

### 6.1 根据返回结果类型选择方法（核心）

**核心原则**：选择 `getBySqlKey` 还是 `getOne`，**完全取决于 SQL 返回的结果类型**，与 SQL ID 是否为"标准"名称无关。

| SQL 返回结果 | 应使用的方法 | 返回类型 | 示例 |
|------------|-------------|---------|------|
| **单个对象** | `dao.getOne("sqlKey", params)` | `Object`（需强转） | `(Entity) dao.getOne("getById", id)` |
| **单个对象（标准ID）** | `dao.getById(id)` | `E`（直接返回） | `dao.getById(id)` |
| **List 集合** | `dao.getBySqlKey("sqlKey", params)` | `List<E>` | `dao.getBySqlKey("getByParams", params)` |

**重要说明**：
1. **SQL ID 名称不重要**：不管是 `getById`、`selectById` 还是 `findById`，只要返回单个对象就应该用 `getOne()`
2. **只看返回结果类型**：检查 Mapper.xml 中的 `resultType` 或实际 SQL 语句
3. **getBySqlKey 总是返回 List**：即使 SQL 只返回一条记录，`getBySqlKey` 也返回 `List<E>`

### 6.2 INSERT 语句必须使用 add 方法

| GenericDao add 方法 | 用途 | 返回类型 | 示例 |
|---------------------|------|----------|------|
| `dao.add(entity)` | 单条插入（标准 SQL ID 为 `add`） | `int` | `dao.add(entity)` |
| `dao.add(String sqlKey, Object param)` | 单条插入（自定义 SQL ID） | `int` | `dao.add("insert", entity)` |
| `dao.add(String sqlKey, List list)` | 批量插入（自定义 SQL ID） | `int` | `dao.add("addBatch", list)` |

**选择规则**：

| Mapper.xml SQL ID | 正确的调用方式 |
|------------------|---------------|
| `add` | `dao.add(entity)` **（优先使用）** |
| `insert` | `dao.add("insert", entity)` |
| `addBatch` | `dao.add("addBatch", list)` |
| `addUnitInfo` | `dao.add("addUnitInfo", entity)` |

### 6.3 GenericDao 内置方法优先原则

当 Mapper.xml 中定义了标准 SQL ID 时，**必须直接使用内置方法**：

| GenericDao 方法 | Mapper.xml SQL ID | 返回类型 | 说明 |
|----------------|------------------|----------|------|
| `dao.add(entity)` | `<insert id="add">` | `int` | **必须使用此方法** |
| `dao.update(entity)` | `<update id="update">` | `int` | **必须使用此方法** |
| `dao.delById(id)` | `<delete id="delById">` | `int` | **必须使用此方法** |
| `dao.getById(id)` | `<select id="getById">` | `E` | **必须使用此方法** |

### 6.4 Dao中可能没有与SQL ID同名的方法

**核心原则**：Mapper.xml 中的 SQL ID 不一定在 Dao 类中有对应的方法！

```java
// ❌ 错误：Dao 类可能为空，没有自定义方法
dao.getByParams(params);           // 方法可能不存在
dao.updateBizStatusByParams(params); // 方法可能不存在

// ✅ 正确：使用 BaseDao 通用方法
dao.getBySqlKey("getByParams", params);
dao.update("updateBizStatusByParams", params);
```

**检查方法**：编写测试前，必须先读取 Dao 类源代码确认方法是否存在。

### 6.5 调用方式对照表

| Mapper.xml SQL ID | 正确的调用方式 | 错误的调用方式 |
|------------------|---------------|---------------|
| **标准 SQL ID** |
| `add` | `dao.add(entity)` | `dao.getBySqlKey("add", entity)` ❌ |
| `getById` | `dao.getById(id)` 或 `(Entity) dao.getOne("getById", id)` | `dao.getBySqlKey("getById", id)` ❌ |
| `update` | `dao.update(entity)` | `dao.getBySqlKey("update", entity)` ❌ |
| `delById` | `dao.delById(id)` | - |
| **自定义 SQL ID** |
| `insert` | `dao.add("insert", entity)` | `dao.getBySqlKey("insert", entity)` ❌；`dao.insert(entity)` ❌ |
| `selectById` | `(Entity) dao.getOne("selectById", id)` | `dao.getBySqlKey("selectById", id)` ❌；`dao.selectById(id)` ❌ |
| `updateById` | `dao.update("updateById", entity)` | `dao.updateById(entity)` ❌ |
| `deleteById` | `dao.delBySqlKey("deleteById", id)` | `dao.deleteById(id)` ❌ |
| `addBatch` | `dao.add("addBatch", list)` | `dao.getBySqlKey("addBatch", list)` ❌；`dao.addBatch(list)` ❌ |

### 高优先级（必须校验）

- [ ] **SQL存在性**：Mapper.xml中是否有对应SQL定义
- [ ] **Dao方法存在性**：Dao类中是否定义了对应方法（没有则用BaseDao通用方法）
- [ ] **调用链路**：在Dao和Service中搜索SQL Key
- [ ] **实体类字段**：只使用实体类中实际存在的setter/getter
- [ ] **禁止var**：所有变量使用明确类型声明
- [ ] **INSERT用add**：`dao.add()`，不是`getBySqlKey()`
- [ ] **SELECT返回单个用getOne**：返回`Object`需强转
- [ ] **SELECT返回List用getBySqlKey**：返回`List<E>`
- [ ] **QueryFilter参数返回PageBean**：不是`List<E>`
- [ ] **Setter参数类型匹配**：String传String，int传int
- [ ] **assertEquals类型匹配**：基本类型用字面量
- [ ] **删除操作判断**：根据实际SQL内容选择方法

### 常规检查

- [ ] 所有import正确
- [ ] `@Rollback`、`@Transactional`、`@Component`注解存在
- [ ] 每个测试方法独立，无相互依赖
- [ ] 断言覆盖重要字段

## 七、常用import

```java
import com.paas.common.web.query.QueryFilter;
import com.paas.model.fm.XxxEntity;
import org.junit.Test;
import org.junit.Ignore;
import org.junit.runner.RunWith;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.test.annotation.Rollback;
import org.springframework.stereotype.Component;
import lombok.extern.slf4j.Slf4j;
import javax.annotation.Resource;
import java.util.*;
import static org.junit.Assert.*;
```

## 八、BaseDao方法速查

| 方法 | 返回类型 | 说明 |
|------|---------|------|
| `add(entity)` | `int` | 单条插入（标准add） |
| `add(sqlKey, entity)` | `int` | 单条插入（自定义SQL ID） |
| `add(sqlKey, list)` | `int` | 批量插入 |
| `getById(id)` | `E` | 按ID查询（标准getById） |
| `getOne(sqlKey, params)` | `Object` | 查询单个对象（需强转） |
| `getBySqlKey(sqlKey, params)` | `List<E>` | 查询列表 |
| `getBySqlKey(sqlKey, QueryFilter)` | `PageBean<E>` | 分页查询 |
| `getAll(queryFilter)` | `List<E>` | 条件查询 |
| `update(entity)` | `int` | 更新（标准update） |
| `update(sqlKey, params)` | `int` | 自定义更新 |
| `delById(id)` | `int` | 按ID删除（标准delById） |
| `delBySqlKey(sqlKey, params)` | `int` | 自定义删除 |
