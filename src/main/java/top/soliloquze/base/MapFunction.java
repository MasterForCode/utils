package top.soliloquze.base;

/**
 * @author wb
 * @date 2019/12/18
 */
@FunctionalInterface
public interface MapFunction<I, K, V> {
    /**
     * 遍历map
     *
     * @param index map的下标
     * @param key   map的key
     * @param value map的value
     */
    void accept(I index, K key, V value);
}
