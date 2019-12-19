package top.soliloquze.base;

import java.util.ConcurrentModificationException;
import java.util.Map;
import java.util.Objects;
import java.util.function.BiConsumer;

/**
 * @author wb
 * @date 2019/12/18
 */
public class Iterables {
    /**
     * 为List的forEach添加下标
     * @param elements
     * @param action
     * @param <E>
     */
    public static <E> void forEach(Iterable<? extends E> elements, BiConsumer<Integer, ? super E> action) {
        Objects.requireNonNull(elements);
        Objects.requireNonNull(action);

        int index = 0;
        for (E element : elements) {
            action.accept(index++, element);
        }
    }

    /**
     * 为Map的forEach添加下标
     * @param elements
     * @param action
     * @param <K>
     * @param <V>
     */
    public static <K, V> void forEach(Map<K, V> elements, MapFunction<Integer, ? super K, ? super V> action) {
        Objects.requireNonNull(elements);
        Objects.requireNonNull(action);

        int index = 0;
        for (Map.Entry<K, V> entry : elements.entrySet()) {
            K k;
            V v;
            try {
                k = entry.getKey();
                v = entry.getValue();
            } catch (IllegalStateException ise) {
                // this usually means the entry is no longer in the map.
                throw new ConcurrentModificationException(ise);
            }
            action.accept(index++, k, v);
        }
    }
}
