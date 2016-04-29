# README

## HOW TO
Two annotations are provided to identify which filed would be exported to excel
file.

* **ExcelAttribute**
  This annotation is common used to annotate field. With several attributes provided.
  + title(String) - The name of cell header
  + isMark(Boolean) - The cell would be marked with red color
  + prompt(String) - The prompt information while mouse over the value
  + isSum(Boolean) - The value of this field would be summed at last row of the sheet
  + combo(Array) - The cell would be chosen not input
  + translate(Annotation) - Details refer @Translate annotation
* **Translate**
  This annotation provides two string value for map one specified value to another literal
  value for easy read, ex 0 translated to 未支付.
  + key - The raw value of data in model
  + value - The value to display in excel
### 0X01 Dependency
``` xml
<dependency>
  <groupId>com.allinmoney.platform</groupId>
  <artifactId>excel-export</artifactId>
  <version>1.0.0-SNAPSHOT</version>
</dependency>
```

### 0x02 Example

1. Annotate the fields in model definition.

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
```
2. Using ExcelUtil class to export data to one OutPutStream
``` java
@Test
public class TestExcelExport {
    private List<Person> persons;
    private static final String FMT = "yyyy-MM-dd";

    @BeforeClass
    public void setUp() {
        persons = new LinkedList<>();
        for (int i = 0; i < 100; i++) {
            Person person = new Person();
            person.setId(i);
            person.setName("西门吹雪 " + i);
            person.setGender(i % 2);
            person.setAge(new BigDecimal("100.12" + i));
            person.setRemark("醒醒,该上班了.");
            persons.add(person);
        }
    }

    public void testExport() {
        ExcelUtil<Person> util = new ExcelUtil<>(Person.class);
        FileOutputStream os = null;
        try {
            os = new FileOutputStream("/tmp/test.xls");
            util.exportDataList(persons, "person", os, FMT);
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
