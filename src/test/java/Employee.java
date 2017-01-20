import com.allinmoney.platform.annotation.ExcelAttribute;

/**
 * Created by chris on 2017/1/17.
 */
public class Employee extends Person {

    @ExcelAttribute(title = "级别", isMark = true, column = "a")
    private Integer level;

    public Integer getLevel() {
        return level;
    }

    public void setLevel(Integer level) {
        this.level = level;
    }
}
