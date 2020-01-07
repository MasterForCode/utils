package top.soliloquize.password;

/**
 * @author wb
 * @date 2020/1/7
 */
public enum Level {
    /**
     * 简单
     */
    EASY(0),
    /**
     * 中等
     */
    MIDDLE(1),
    /**
     * 强
     */
    STRONG(2),
    /**
     * 非常强
     */
    VERY_STRONG(3),
    /**
     * 特别强
     */
    EXTREMELY_STRONG(4);
    Integer value;

    Level(Integer value) {
        this.value = value;
    }
}
