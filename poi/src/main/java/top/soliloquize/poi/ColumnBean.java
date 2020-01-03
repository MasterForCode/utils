package top.soliloquize.poi;

import lombok.Builder;
import lombok.Data;

import java.lang.reflect.Method;

/**
 * @author wb
 * @date 2020/1/2
 */
@Data
@Builder
public class ColumnBean {
    /**
     * get时是返回值类型
     * set时是参数类型
     */
    String type;
    Integer columnIndex;
    /**
     * get或set方法
     */
    Method method;
    /**
     * 字段名
     */
    String fieldName;
}
