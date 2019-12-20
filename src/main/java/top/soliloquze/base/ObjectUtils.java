package top.soliloquze.base;

import com.fasterxml.jackson.databind.ObjectMapper;

import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @author wb
 * @date 2019/12/18
 */
public class ObjectUtils {
    /**
     * 对象列表转map列表
     * @param objects 对象列表
     * @param <T> T
     * @return map列表
     */
    public static <T> List<Map<String, Object>> objects2Maps(List<T> objects) {
        return objects.parallelStream().map(ObjectUtils::object2Map).collect(Collectors.toList());
    }

    /**
     * 对象转map
     * @param obj 对象
     * @return map
     */
    public static Map<String, Object> object2Map(Object obj) {
        ObjectMapper objectMapper = new ObjectMapper();
        return objectMapper.convertValue(obj,  objectMapper.getTypeFactory().constructParametricType(Map.class, String.class, Object.class));
    }

    public static <T> List<T> listRequireNonNull(List<T> data) {
        if (data == null || data.size() == 0) {
            throw new NullPointerException();
        } else {
            return data;
        }
    }
}
