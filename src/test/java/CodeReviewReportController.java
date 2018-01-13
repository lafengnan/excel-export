import com.allinmoney.platform.ScaUtil;
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
import java.util.*;
import java.util.stream.Collectors;


/**
 * Created by Chris Pan on 1/11/2018.
 * @author Chris Pan
 */
@Setter
@Getter
public class CodeReviewReportController {
    private static final Logger logger = LoggerFactory.getLogger(CodeReviewReportController.class);
    private String sourceDir;
    private String sourceFileName;
    private String sheetName;
    /* Default splitter is windows pattern
     */
    private String splitter = "\\";

    private interface Conf{
        String V_DIR = "V:\\Operations & Technology\\MVP\\MVP Software Launch\\08 Code Review\\Daily Update";
        String[] MODULES = {
                "BPS",
                "TPP",
                "TSP",
                "Risk",
                "Batch",
                "Online",
                "mngportal",
                "Integrator",
                "File Server"
        };
    }

    @Setter
    @Getter
    @ToString
    @AllArgsConstructor
    private static class DefectFilter {
        Field filed;
        Object[] values;
    }

    public List<Defect> importDefectFromExcel(String path) {
        ExcelUtil<Defect> util = new ExcelUtil<>(Defect.class);
        List<Defect> defects = util.importData(path, sheetName, "MM/dd/yyyy");
        defects.forEach(defect -> logger.info(defect.toString()));
        return defects.stream().filter(defect -> defect.getNo() != null).collect(Collectors.toList());
    }

    public void exportHighDefectList() {
        List<File> files = ScaUtil.listFiles(new File(Conf.V_DIR), true);
        Map<String, List<File>> groupFiles = new HashMap<>(Conf.MODULES.length);
        files.forEach(file -> {
            Arrays.asList(Conf.MODULES).forEach(module -> {
                if (file.getAbsolutePath().contains(module)) {
                    List<File> moduleFiles = groupFiles.getOrDefault(module, new LinkedList<>());
                    moduleFiles.add(file);
                    groupFiles.put(module, moduleFiles);
                }
            });
        });

        groupFiles.forEach((module, reports) -> {
            logger.info("generating report for: " + module);
            setSourceDir(Conf.V_DIR + splitter + module);
            List<File> excelFiles = reports.stream()
                    .filter(report -> report.getName().contains("code_review"))
                    .collect(Collectors.toList());
            setSourceFileName(excelFiles.get(excelFiles.size() - 1).getName());
            /* Retrieving sheet 0 in default
             */
            setSheetName("");
            String path = sourceDir + splitter + sourceFileName;

            try {
                exportDefectList(
                        importDefectFromExcel(path),
                        Arrays.asList(
                                new DefectFilter(Defect.class.getDeclaredField("severity"), new String[]{Defect.Severity.High.name()}),
                                new DefectFilter(Defect.class.getDeclaredField("status"), Defect.Status.unClosedStatusStrings().toArray())
                        )
                );
            } catch (NoSuchFieldException e) {
                e.printStackTrace();
            } catch (Exception e) {
                e.printStackTrace();
                logger.debug(e.getMessage());
            }
        });

    }

    public void exportDefectList(List<Defect> defects, List<DefectFilter> filters) {
        ExcelUtil<Defect> util = new ExcelUtil<>(Defect.class);
        if (filters != null && !filters.isEmpty()) {
            filters.forEach(f -> {
                logger.debug("filtering defect with filter: " + f.toString());
                f.getFiled().setAccessible(true);
                defects.removeIf(defect -> {
                    boolean hitFlag = false;
                    try {
                        Object[] values = f.getValues();
                        for (int i = 0; !hitFlag && i < values.length; i++) {
                            hitFlag = f.getFiled().get(defect).equals(values[i]);
                        }
                    } catch (IllegalAccessException e) {
                        logger.debug(e.getMessage());
                    } catch (NullPointerException e) {
                        hitFlag = true;
                        logger.debug(e.getMessage());
                    }
                    return !hitFlag;
                });
            });
        }
        String path = sourceDir + splitter + sourceFileName + Instant.now().getEpochSecond() + ".xls";
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
        controller.exportHighDefectList();
    }

}
