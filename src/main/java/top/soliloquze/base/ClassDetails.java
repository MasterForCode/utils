package top.soliloquze.base;

import lombok.Builder;
import lombok.Data;

import java.lang.reflect.Method;
import java.lang.reflect.Type;

/**
 * @author wb
 * @date 2020/1/10
 */
@Data
@Builder
public class ClassDetails {
    private String fieldName;
    private Type type;
    private Method getMethod;
    private Method setMethod;
}
