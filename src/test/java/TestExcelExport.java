import com.allinmoney.platform.excel.ExcelUtil;
import org.testng.Assert;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.io.*;
import java.math.BigDecimal;
import java.util.LinkedList;
import java.util.List;

/**
 * Created by chris on 16/4/28.
 */

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
            person.setName(i % 2 == 0?("西门吹雪 " + i):null);
            person.setGender(i % 2);
            person.setAge(new BigDecimal("100.12" + i));
            person.setRemark("醒醒,该上班了. 你没看见川普都当总统了，大清朝要亡了吗？");
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
