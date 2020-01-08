package top.soliloquze.base;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * @author wb
 * @date 2020/1/8
 */
public class JsonUtil {
    private static final ObjectMapper MAPPER;

    static {
        MAPPER = new ObjectMapper();
        //解析器支持解析单引号
        MAPPER.configure(JsonParser.Feature.ALLOW_SINGLE_QUOTES,true);
        //解析器支持解析结束符
        MAPPER.configure(JsonParser.Feature.ALLOW_UNQUOTED_CONTROL_CHARS,true);
    }

    public static <T> T getBeanFromJson(String jsonStr, Class<T> clz) throws JsonProcessingException {
        return MAPPER.readValue(jsonStr, clz);
    }

    public static <T> T getBeanFromJsonEx(String jsonStr, Class<T> clz) {
        try {
            return MAPPER.readValue(jsonStr, clz);
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> String json(T obj) throws JsonProcessingException {
        return MAPPER.writeValueAsString(obj);
    }

    public static <T> String jsonEx(T obj) {
        try {
            return MAPPER.writeValueAsString(obj);
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e);
        }
    }
}
