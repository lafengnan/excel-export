# Introduction

## Version

Current version is 1.0.9-SNAPSHOT.

## Change Log
* 2016-06-15 Fix summary overwrite issue.
* 2016-07-04 Fill in empty data if there is no data content.
* 2016-09-06 1.0.4 Add view group annotation.
* 2016-10-28 1.0.5 Add new API to support multiple data source exportation.
* 2016-11-23 1.0.6 Set column to auto size.(**Deprecated**)
* 2016-11-24 1.0.7 fix auto size issue introduced in 1.0.7
* 2017-01-17 1.0.8-SNAPSHOT
  * Extends maximum row from 65535 to 1,048,576 as office 2013 spec
  * Fix issue of export with super class
  * Refactor the source code of exportDataList
  * Add two methods to export data list:
    * public boolean exportDataList(List<T> dataList, String sheetName, OutputStream os, boolean superFlag) {
    * public boolean exportDataList(List<T> dataList, String sheetName, OutputStream os, String dateFmt, boolean superFlag)
  * Add one method to export multiple data list:
    * public boolean exportMultipleDataList(String sheetName, boolean superFlag, OutputStream os, List<?>... dataList)
* 2017-01-18 1.0.8 Update Java document and change version number
* 2017-01-20 1.0.9-SNAPSHOT
  * Fix issue of locating for duplicate specified columns
  * Fix issue of mark font display issue of exportDataList with multiple data source

## HOW TO
Two annotations are provided to identify which filed would be exported to excel
file.

* **ExcelAttribute**
  This annotation is common used to annotate field. With several attributes provided.
  + title(String) - The name of cell header
  + isMark(Boolean, default *false*) - The cell would be marked with red color
  + isExport(Boolean, default *true*) - The cell would be exported
  + prompt(String, default *""*) - The prompt information while mouse over the value
  + isSum(Boolean, default *false*) - The value of this field would be summed at last row of the sheet
  + combo(Array, default *[]*) - The cell would be chosen not input
  + translate(Annotation Array, default *[]*) - Details refer @Translate annotation
  + groups(Array, default *[]*) - If set only annotated view would be exported
  + format(String, default *yyyy-MM-dd HH:mm:ss*) - The date format
* **Translate**
  This annotation provides two string to map one specified value to another literal
  value for human read, eg. "0" translated to "未支付".
  + key - The raw value of data in model
  + value - The literal string to display in excel

### 0X01 Dependency

``` xml
<dependency>
  <groupId>com.allinmoney.platform</groupId>
  <artifactId>excel-export</artifactId>
  <version>1.0.9-SNAPSHOT</version>
</dependency>
```

### 0x02 Example

**Annotate the fields in model definition**

``` java
import com.allinmoney.platform.annotation.ExcelAttribute;
import com.allinmoney.platform.annotation.Translate;

public class Person {

    @ExcelAttribute(title = "ID", isMark = true)
    private Integer id;

    @ExcelAttribute(title= "name", isMark = false)
    private String name;

    @ExcelAttribute(title= "年龄", isMark = false, isSum = true)
    private BigDecimal age;

    @ExcelAttribute(title = "性别", translate= {@Translate(key = "0", value = "女"),
    @Translate(key = "1", value = "男")})
    private Integer gender;

    @ExcelAttribute(title= "备注", prompt = "辅助信息")
    private String remark;

    public String getRemark() {
        return remark;
    }

    public void setRemark(String remark) {
        this.remark = remark;
    }

    public BigDecimal getAge() {
        return age;
    }

    public void setAge(BigDecimal age) {
        this.age = age;
    }

    public Integer getGender() {
        return gender;
    }

    public void setGender(Integer gender) {
        this.gender = gender;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}

public class Employee extends Person {

    @ExcelAttribute(title = "级别", isMark = true, column = "b")
    private Integer level;

    public Integer getLevel() {
        return level;
    }

    public void setLevel(Integer level) {
        this.level = level;
    }
}
```

**Using ExcelUtil class to export data to one OutPutStream**
``` java
@Test
public class TestExcelExport {
    private List<Employee> employees;
    private List<Person> persons;
    private static final String FMT = "yyyy-MM-dd";

    @BeforeClass
    public void setUp() {
        employees = new LinkedList<>();
        persons = new LinkedList<>();
        for (int i = 0; i < 100; i++) {
            Person person = new Person();
            Employee employee = new Employee();
            person.setId(i * 1000000);
            employee.setId(i * 1000000);
            person.setName("西门吹雪叶孤城陆小凤东邪西毒南帝北丐中神通");
            employee.setName("秦皇汉武唐宗宋祖");
            person.setGender(i % 2);
            employee.setGender(i % 2);
            person.setBirthDay(new Date());
            employee.setBirthDay(new Date());
            person.setAge(new BigDecimal("100.12" + i));
            employee.setAge(new BigDecimal("200.12" + i));
            person.setRemark("醒醒,该上班了. 你没看见川普都当总统了，大清朝要亡了吗？");
            employee.setRemark("一年又过去了，奖金呢？？？");
            employee.setLevel(i);
            persons.add(person);
            employees.add(employee);
        }
    }

    public void testExport() {
        ExcelUtil<Person> util = new ExcelUtil<>(Person.class);
        ExcelUtil<Employee> util1 = new ExcelUtil<>(Employee.class);
        FileOutputStream os = null;
        try {
            os = new FileOutputStream("/tmp/test.xls");
            util.exportDataList(persons, "person", os, false);
            os = new FileOutputStream("/tmp/test-2.xls");
            util1.exportDataList(employees, "person", os, FMT, true);
            os = new FileOutputStream("/tmp/test-3.xls");
            util.exportMultipleDataList("multiple", true, os, employees, persons);
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }

        Assert.assertNotNull(os);
        try {
            Assert.assertEquals(os.getChannel().size(),
                    new BufferedInputStream(new FileInputStream("/tmp/test.xls")).available());
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }
}
```
