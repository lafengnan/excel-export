import com.allinmoney.platform.excel.ExcelUtil;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.LinkedList;
import java.util.List;

/**
 * Created by chris on 16/4/28.
 */
public class TestPerson {
    public static void main(String[] args) {
        List<Person> persons = new LinkedList<>();
        for (int i = 0; i < 100; i++) {
            Person person = new Person();
            person.setId(i);
            person.setName("西门吹雪 " + i);
            person.setGender(i % 2);
            person.setAge(new BigDecimal("100.12" + i));
            person.setRemark("醒醒,该上班了.");
            persons.add(person);
        }

        try {

            FileOutputStream os = new FileOutputStream("/tmp/test.xls");
            ExcelUtil<Person> util = new ExcelUtil<>(Person.class);
            util.exportDataList(persons, "person", os, "yyyy-MM-dd");
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }

    }
}
