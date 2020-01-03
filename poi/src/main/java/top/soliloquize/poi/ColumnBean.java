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
    String paramType;
    Integer columnIndex;
    Method method;
}
