import com.allinmoney.platform.excel.ExcelUtil;
import org.testng.Assert;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.io.*;
import java.math.BigDecimal;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

/**
 * Created by chris on 16/4/28.
 */

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
            os = new FileOutputStream("/tmp/test-4.xls");
            util.exportDataList("multiple", os, persons, employees);
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
