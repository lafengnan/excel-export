import com.allinmoney.platform.annotation.ExcelAttribute;
import lombok.Data;

/**
 * Created by e077173 on 1/10/2018.
 */
@Data
public class Defect {
    @Data
    private static class Header {
        @ExcelAttribute(title = "No. 系统名-ID")
        private String title;
        private String moduleName;
        private String inspectDate;
        private String codeVersion;
        private int fileCount;
        private int lineOfCode;
        private int defectCount;
        private int defectPerHour;
        private double inspectVelocity;
        private int defectsPerKiloLine;
        private String shortDesc;
    }

    enum Severity {
        High, Medium, Low
    }

    enum Status {
        OPEN, WORKING, VERIFY, CLOSED
    }

    @ExcelAttribute(title = "NO.")
    private String no;
    @ExcelAttribute(title = "问题位置\nPosition")
    private String position;
    @ExcelAttribute(title = "问题描述\nDescription")
    private String desc;
    @ExcelAttribute(title = "提出者\nReviewer")
    private String reviewer;
    @ExcelAttribute(title = "问题类型\nType")
    private String defectType;
    @ExcelAttribute(title = "问题级别\nSeverity")
    private String severity;
    @ExcelAttribute(title = "解决措施\nMethods")
    private String methods;
    @ExcelAttribute(title = "责任人\nResponsibility")
    private String responsibility;
    @ExcelAttribute(title = "预计完成时间\nScheduled End\nDate")
    private String schEndDate;
    @ExcelAttribute(title = "发现时间\nReport Date")
    private String reportDate;
    @ExcelAttribute(title = "实际完成时间\nActual End Date")
    private String actualEndDate;
    @ExcelAttribute(title = "修改工作量\nModifying\nEffort")
    private int effort;
    @ExcelAttribute(title = "状态\nStatus")
    private String status;
    @ExcelAttribute(title = "备注\nRemark")
    private String remark;

}
