package top.soliloquize.poi;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import top.soliloquze.base.Iterables;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * @author wb
 * @date 2019/12/18
 */
public class PoiUtils {
    public static void main(String[] args) throws JsonProcessingException {
        XSSFWorkbook workbook = init();

        // 字体样式
        XSSFFont xssfFont = workbook.createFont();
        // 加粗
        xssfFont.setBold(true);
        // 字体名称
        xssfFont.setFontName("楷体");
        // 字体大小
        xssfFont.setFontHeight(12);

        // 表头样式
        XSSFCellStyle headStyle = workbook.createCellStyle();
        // 设置字体css
        headStyle.setFont(xssfFont);
        // 竖向居中
        headStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 横向居中
        headStyle.setAlignment(HorizontalAlignment.CENTER);
        // 边框
        headStyle.setBorderBottom(BorderStyle.THIN);
        headStyle.setBorderLeft(BorderStyle.THIN);
        headStyle.setBorderRight(BorderStyle.THIN);
        headStyle.setBorderTop(BorderStyle.THIN);

        // 内容字体样式
        XSSFFont contFont = workbook.createFont();
        // 加粗
        contFont.setBold(false);
        // 字体名称
        contFont.setFontName("楷体");
        // 字体大小
        contFont.setFontHeight(11);
        // 内容样式
        XSSFCellStyle contentStyle = workbook.createCellStyle();
        // 设置字体css
        contentStyle.setFont(contFont);
        // 竖向居中
        contentStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 横向居中
        //contentStyle.setAlignment(HorizontalAlignment.CENTER);
        // 边框
        contentStyle.setBorderBottom(BorderStyle.THIN);
        contentStyle.setBorderLeft(BorderStyle.THIN);
        contentStyle.setBorderRight(BorderStyle.THIN);
        contentStyle.setBorderTop(BorderStyle.THIN);

        // 自动换行
        contentStyle.setWrapText(true);

        // 数字样式
        XSSFCellStyle numStyle = workbook.createCellStyle();
        // 设置字体css
        numStyle.setFont(contFont);
        // 竖向居中
        numStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 横向居中
        numStyle.setAlignment(HorizontalAlignment.CENTER);
        // 边框
        numStyle.setBorderBottom(BorderStyle.THIN);
        numStyle.setBorderLeft(BorderStyle.THIN);
        numStyle.setBorderRight(BorderStyle.THIN);
        numStyle.setBorderTop(BorderStyle.THIN);

        // 标题字体样式
        XSSFFont titleFont = workbook.createFont();
        // 加粗
        titleFont.setBold(false);
        // 字体名称
        titleFont.setFontName("宋体");
        // 字体大小
        titleFont.setFontHeight(16);

        // 标题样式
        XSSFCellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFont(titleFont);
        // 竖向居中
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 横向居中
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        // 边框
        titleStyle.setBorderBottom(BorderStyle.THIN);
        titleStyle.setBorderLeft(BorderStyle.THIN);
        titleStyle.setBorderRight(BorderStyle.THIN);
        titleStyle.setBorderTop(BorderStyle.THIN);
        XSSFSheet sheet1 = initSheetWithStyle(workbook, "sheet1", Arrays.asList("姓名", "年龄", "身高"), contentStyle);
        XSSFSheet sheet2 = initSheet(workbook, "sheet2", Arrays.asList("姓名", "年龄", "身高"));
        ObjectMapper objectMapper = new ObjectMapper();
        List<User> userList = objectMapper.readValue("[{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"}]", objectMapper.getTypeFactory().constructParametricType(List.class, User.class));
        fillDataWithStyle(sheet1, userList, contentStyle,1);
        List<User> userList1 = objectMapper.readValue("[{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"},{\"name\":\"张三\",\"age\":20,\"height\":\"175cm\"}]", objectMapper.getTypeFactory().constructParametricType(List.class, User.class));
        fillData(sheet2, userList1, 1 + userList.size());
        createExcel("test", workbook);
    }

    private static XSSFWorkbook init() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        return workbook;
    }

    private static XSSFSheet initSheet(XSSFWorkbook workbook, String title, List<String> columns) {
        XSSFSheet sheet = workbook.createSheet(title);

        Row row = sheet.createRow(0);
        Iterables.forEach(columns, (i, each) -> {
            Cell cell = row.createCell(i);
            cell.setCellValue(columns.get(i));
        });
        return sheet;
    }

    private static XSSFSheet initSheetWithStyle(XSSFWorkbook workbook, String title, List<String> columns, XSSFCellStyle contentStyle) {
        XSSFSheet sheet = workbook.createSheet(title);

        Row row = sheet.createRow(0);
        Iterables.forEach(columns, (i, each) -> {
            Cell cell = row.createCell(i);
            cell.setCellStyle(contentStyle);
            cell.setCellValue(columns.get(i));
        });
        return sheet;
    }

    private static void fillData(XSSFSheet sheet, List<?> data, int rowNum) {
        List<Map<String, Object>> fixedData = top.soliloquze.base.ObjectUtils.objects2Maps(data);
        Iterables.forEach(fixedData, (i, each) -> {
            Row rowBody = sheet.createRow(rowNum + i);
            Iterables.forEach(fixedData.get(i), (index, k, v) -> {
                Cell cellBody = rowBody.createCell(index);
                cellBody.setCellValue(ObjectUtils.defaultIfNull(v, "").toString());
            });
        });
    }

    private static void fillDataWithStyle(XSSFSheet sheet, List<?> data, XSSFCellStyle contentStyle, int rowNum) {
        List<Map<String, Object>> fixedData = top.soliloquze.base.ObjectUtils.objects2Maps(data);
        Iterables.forEach(fixedData, (i, each) -> {
            Row rowBody = sheet.createRow(rowNum + i);
            Iterables.forEach(fixedData.get(i), (index, k, v) -> {
                Cell cell = rowBody.createCell(index);
                cell.setCellStyle(contentStyle);
                cell.setCellValue(ObjectUtils.defaultIfNull(v, "").toString());
            });
        });
    }

    private static File createExcel(String fileName, XSSFWorkbook workbook) {
        fileName += ".xlsx";
        File file = new File("", fileName);
        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(file);
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                Objects.requireNonNull(outputStream).close();
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return file;
    }
}
