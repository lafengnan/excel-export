import com.allinmoney.platform.excel.ExcelUtil;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.time.Instant;
import java.util.Arrays;
import java.util.List;

/**
 * Created by Chris Pan on 1/11/2018.
 * @author Chris Pan
 */

@Setter
public class CodeReviewReportController {
    private static final Logger logger = LoggerFactory.getLogger(CodeReviewReportController.class);
    private String sourceDir;
    private String sourceFileName;
    private String sheetName;

    @Setter
    @Getter
    @ToString
    @AllArgsConstructor
    private static class DefectFilter {
        Field filed;
        Object value;
    }

    public List<Defect> importDefectFromExcel() {
        ExcelUtil<Defect> util = new ExcelUtil<>(Defect.class);
        List<Defect> defects = util.importData(sourceDir, sourceFileName, sheetName, "MM/dd/yyyy");
        defects.forEach(defect -> logger.info(defect.toString()));
        return defects;
    }

    public void exportDefectList(List<Defect> defects, List<DefectFilter> filters) {
        ExcelUtil<Defect> util = new ExcelUtil<>(Defect.class);
        if (filters != null && !filters.isEmpty()) {
            filters.forEach(f -> {
                logger.debug("filtering defect with filter: " + f.toString());
                f.getFiled().setAccessible(true);
                defects.removeIf(defect -> {
                    try {
                        return !f.getFiled().get(defect).equals(f.getValue());
                    } catch (IllegalAccessException e) {
                        logger.debug(e.getMessage());
                        return false;
                    }
                });
            });
        }
        String path = sourceDir + "\\" + sourceFileName + Instant.now().getEpochSecond() + ".xls";
        FileOutputStream outputStream = null;
        try {
            File file = new File(path);
            outputStream = new FileOutputStream(file);
            util.exportDataList(defects, sheetName, outputStream, "yyyy-MM-dd");
        } catch (IOException e) {
            logger.info(e.getMessage());
        } finally {
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void main(String... args) throws Exception{
        CodeReviewReportController controller = new CodeReviewReportController();
        controller.setSheetName("File Server");
        controller.setSourceDir("C:\\Users\\chris\\Documents\\Work\\Code review");
        controller.setSourceFileName("code_review_track_20171206_fileserver.xlsx");
        controller.exportDefectList(
                controller.importDefectFromExcel(),
                Arrays.asList(
                        new DefectFilter(Defect.class.getDeclaredField("severity"), "High"),
                        new DefectFilter(Defect.class.getDeclaredField("status"), "Open")
                )
        );
    }

}
