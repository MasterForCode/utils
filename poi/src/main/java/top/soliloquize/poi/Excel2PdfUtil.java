package top.soliloquize.poi;

import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Table;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * @author wb
 * @date 2019/12/23
 */
public class Excel2PdfUtil {

    /**
     * 初始化Workbook
     *
     * @param filePath excel路径
     * @return Workbook
     * @throws IOException IOException
     */
    public static Workbook initWorkbook(String filePath) throws IOException {
        Objects.requireNonNull(filePath);
        if (!ExcelUtil.validatorSuffix(FilenameUtils.getExtension(filePath))) {
            throw new RuntimeException("不支持的格式");
        }
        return WorkbookFactory.create(new BufferedInputStream(new FileInputStream(filePath)));
    }

    /**
     * 初始化Workbook
     *
     * @param file excel 文件
     * @return Workbook
     * @throws IOException IOException
     */
    public static Workbook initWorkbook(File file) throws IOException {
        Objects.requireNonNull(file);
        if (!ExcelUtil.validatorSuffix(FilenameUtils.getExtension(file.getAbsolutePath()))) {
            throw new RuntimeException("不支持的格式");
        }
        return WorkbookFactory.create(new BufferedInputStream(new FileInputStream(file)));
    }

    /**
     * 通过sheet名获取sheet
     *
     * @param workbook  workbook
     * @param sheetName sheet名
     * @return Sheet
     */
    public static Sheet getSheetByName(Workbook workbook, String sheetName) {
        Objects.requireNonNull(workbook);
        Objects.requireNonNull(sheetName);
        return workbook.getSheet(sheetName);
    }

    /**
     * 通过下标获取sheet
     *
     * @param workbook workbook
     * @param index    index
     * @return Sheet
     */
    public static Sheet getSheetByIndex(Workbook workbook, int index) {
        Objects.requireNonNull(workbook);
        if (index <= 0) {
            throw new RuntimeException("sheet下标不能小于0");
        }
        return workbook.getSheetAt(index);
    }

    public static void main(String[] args) throws IOException {
        excel2Pdf("/test.xlsx", "test.pdf", 4);
    }

    /**
     * excel转pdf
     *
     * @param excelFilePath excel路径
     * @param pdfFilePath   pdf路径
     * @param maxColumnSize 最大列数
     * @return pdf文件
     * @throws IOException IOException
     */
    public static File excel2Pdf(String excelFilePath, String pdfFilePath, int maxColumnSize) throws IOException {
        Objects.requireNonNull(excelFilePath);
        Objects.requireNonNull(pdfFilePath);
        if (maxColumnSize <= 0) {
            throw new RuntimeException("列数不能等于0");
        }
        Workbook workbook = initWorkbook(excelFilePath);
        File file = PdfUtil.initFile(pdfFilePath);
        Document document = PdfUtil.initPdfPage(PageSize.A4, false, file);
        int sheetNum = workbook.getNumberOfSheets();
        for (int i = 0; i < sheetNum; i++) {
            Table table = PdfUtil.initTable(maxColumnSize);
            action(getSheetByIndex(workbook, i), maxColumnSize).forEach(each -> PdfUtil.fillData(table, each));
            PdfUtil.fillDoc(document, table);
        }
        document.close();
        return file;
    }

    /**
     * excel转pdf
     *
     * @param excelFile excel文件
     * @param pdfFile   pdf文件
     * @param maxColumnSize 最大列数
     * @return pdf文件
     * @throws IOException IOException
     */
    public static File excel2Pdf(File excelFile, File pdfFile, int maxColumnSize) throws IOException {
        Objects.requireNonNull(excelFile);
        Objects.requireNonNull(pdfFile);
        if (maxColumnSize <= 0) {
            throw new RuntimeException("列数不能等于0");
        }
        Workbook workbook = initWorkbook(excelFile);
        Document document = PdfUtil.initPdfPage(PageSize.A4, false, pdfFile);
        int sheetNum = workbook.getNumberOfSheets();
        for (int i = 0; i < sheetNum; i++) {
            Table table = PdfUtil.initTable(maxColumnSize);
            action(getSheetByIndex(workbook, i), maxColumnSize).forEach(each -> PdfUtil.fillData(table, each));
            PdfUtil.fillDoc(document, table);
        }
        document.close();
        return pdfFile;
    }

    /**
     * 获取sheet中的内容
     *
     * @param sheet         sheet
     * @param maxColumnSize 该sheet中最大的列数
     * @return List<List < String>>
     */
    public static List<List<String>> action(Sheet sheet, int maxColumnSize) {
        Objects.requireNonNull(sheet);
        if (maxColumnSize <= 0) {
            throw new RuntimeException("列数不能等于0");
        }
        int firstRowNum = sheet.getFirstRowNum();
        Integer lastRowNum = sheet.getLastRowNum();
        List<List<String>> data = new ArrayList<>();
        for (Integer i = firstRowNum; i < lastRowNum; i++) {
            List<String> detail = new ArrayList<>();
            Row row = sheet.getRow(i);
            // 去除空行
            if (row != null) {
                short firstCellNum = row.getFirstCellNum();
                short lastCellNum = row.getLastCellNum();
                for (int j = firstCellNum; j < lastCellNum; j++) {
                    if (row.getCell(j) != null) {
                        detail.add(row.getCell(j).getStringCellValue());
                    } else {
                        detail.add("");
                    }
                }
                // 空白cell置为空
                if (detail.size() < maxColumnSize) {
                    for (int j = detail.size(); j < maxColumnSize; j++) {
                        detail.add("");
                    }
                }
                data.add(detail);
            }
        }
        return data;
    }

}
