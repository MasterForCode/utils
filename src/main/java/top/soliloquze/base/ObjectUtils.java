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
    public static <T> List<Map<String, Object>> objects2Maps(List<T> objects) {
        return objects.parallelStream().map(ObjectUtils::object2Map).collect(Collectors.toList());
    }

    public static Map<String, Object> object2Map(Object obj) {
        ObjectMapper objectMapper = new ObjectMapper();
        return objectMapper.convertValue(obj,  objectMapper.getTypeFactory().constructParametricType(Map.class, String.class, Object.class));
    }
}
