import com.allinmoney.platform.annotation.ExcelAttribute;
import com.allinmoney.platform.annotation.Translate;

import java.math.BigDecimal;

/**
 * Created by chris on 16/4/28.
 */
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
