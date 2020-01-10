package top.soliloquze.base;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.AccessibleObject;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.Stream;

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

    private static <T> List<ClassDetails> analysisClz(Class<T> clz) {

        List<Field> fieldList = Arrays.stream(clz.getDeclaredFields()).filter(field -> !field.isAccessible()).collect(Collectors.toList());
        List<Method> methodList = Arrays.stream(clz.getDeclaredMethods()).filter(method -> Modifier.isPublic(method.getModifiers())).collect(Collectors.toList());
        List<ClassDetails> classDetailsList = new ArrayList<>();
        fieldList.forEach(field -> {
            Method setMethod = Iterables.findFirst(methodList, method -> method.getName().equalsIgnoreCase("set" + StringUtils.capitalize(field.getName())));
            Method getMethod = Iterables.findFirst(methodList, method -> method.getName().equalsIgnoreCase("get" + StringUtils.capitalize(field.getName())));
            ClassDetails classDetails = ClassDetails.builder().fieldName(field.getName()).type(field.getGenericType()).setMethod(setMethod).getMethod(getMethod).build();
            classDetailsList.add(classDetails);
        });
       return classDetailsList;
    }

    public static void main(String[] args) {
        System.out.println(analysisClz(A.class));
    }
}
