package top.soliloquze.base;

import java.text.DecimalFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.Objects;

/**
 * @author wb
 * @date 2020/1/2
 */
public class FormatUtils {

    public static String DATE_FORMAT_NORMAL = "yyyy-MM-dd";

    public static DateTimeFormatter normalDateTimeFormatter = DateTimeFormatter.ofPattern(FormatUtils.DATE_FORMAT_NORMAL);

    public static DecimalFormat df = new DecimalFormat("####.##");

    public static LocalDate date2LocalDate(Date date) {
        if (date == null) {
            return null;
        }
        return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }

    public static String date2String(Date date) {
        if (date == null) {
            return "";
        }
        return FormatUtils.date2LocalDate(date).format(normalDateTimeFormatter);
    }

    public static Date localDate2Date(LocalDate localDate) {
        if (localDate == null) {
            return null;
        }
        ZonedDateTime zdt = localDate.atStartOfDay(ZoneId.systemDefault());

        return Date.from(zdt.toInstant());
    }

    public static LocalDate string2LocalDate(String value) {
        if ("".equals(value) || value == null) {
            return null;
        }
        return LocalDate.parse(value, normalDateTimeFormatter);
    }

    public static LocalDate string2LocalDate(String value, String format) {
        if (value == null || "".equals(value) || format == null) {
            return null;
        }
        return LocalDate.parse(value, DateTimeFormatter.ofPattern(format));
    }
}
