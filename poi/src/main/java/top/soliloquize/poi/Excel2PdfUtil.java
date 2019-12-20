package top.soliloquize.poi;


import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Text;
import com.itextpdf.layout.property.TextAlignment;
import com.itextpdf.layout.property.VerticalAlignment;
import top.soliloquze.base.ObjectUtils;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/**
 * @author wb
 * @date 2019/12/20
 */
public class Excel2PdfUtil {
    private static PdfFont pdfFont;

    private static float fontSize = 15;

    static {
        try {
            pdfFont = PdfFontFactory.createFont("STSong-Light", "UniGB-UCS2-H", true);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws IOException {
        Document document = initPdfPage(PageSize.A4, true, "a.pdf");
        Table table = initTable(5);
        fillDoc(document, table);

        setHeader(document, "h", 1, document.getPdfDocument().getPage(1).getPageSize().getWidth() / 2,
                 20, TextAlignment.CENTER, VerticalAlignment.MIDDLE);

        setFooter(document, "ha", 1, document.getPdfDocument().getPage(1).getPageSize().getWidth() / 2,
                 20, TextAlignment.CENTER, VerticalAlignment.MIDDLE);
        document.close();

    }


    public static void setHeader(Document document, String content, int pageIndex, float x, float margin, TextAlignment textAlignment, VerticalAlignment verticalAlignment) {
        Objects.requireNonNull(document);
        Objects.requireNonNull(content);
        Objects.requireNonNull(textAlignment);
        Objects.requireNonNull(verticalAlignment);
        document.showTextAligned(new Paragraph(content)
                .setFont(Excel2PdfUtil.pdfFont).setFontSize(Excel2PdfUtil.fontSize), x, document.getPdfDocument().getPage(1).getPageSize().getTop() - margin, pageIndex, textAlignment, verticalAlignment, 0);
    }
    public static void setFooter(Document document, String content, int pageIndex, float x, float margin, TextAlignment textAlignment, VerticalAlignment verticalAlignment) {
        Objects.requireNonNull(document);
        Objects.requireNonNull(content);
        Objects.requireNonNull(textAlignment);
        Objects.requireNonNull(verticalAlignment);
        document.showTextAligned(new Paragraph(content)
                .setFont(Excel2PdfUtil.pdfFont).setFontSize(Excel2PdfUtil.fontSize), x, document.getPdfDocument().getPage(1).getPageSize().getBottom() + margin, pageIndex, textAlignment, verticalAlignment, 0);
    }

    public static Document initPdfPage(PageSize pageSize, boolean rotate, String filePath) throws IOException {
        if (rotate) {
            pageSize = pageSize.rotate();
        }
        PdfDocument pdfDoc = new PdfDocument(new PdfWriter(new File(filePath)));
        return new Document(pdfDoc, pageSize);
    }

    public static Table initTable(int columnSize) {
        if (columnSize == 0) {
            throw new RuntimeException("列数不能等于0");
        }
        Table table = new Table(columnSize);
        return table;
    }

    public static void fillData(Table table, List<String> data) {
        Objects.requireNonNull(table);
        ObjectUtils.listRequireNonNull(data);
        data.stream().map(each -> new Paragraph(new Text(each).setFont(Excel2PdfUtil.pdfFont))).forEach(table::addCell);


    }

    public static void fillData(Table table, List<String> data, PdfFont font) {
        Objects.requireNonNull(table);
        ObjectUtils.listRequireNonNull(data);
        Objects.requireNonNull(font);

        data.stream().map(each -> new Paragraph(new Text(each).setFont(font))).forEach(table::addCell);

    }

    public static void fillDoc(Document document, Table table) {
        document.add(table);
    }
}
