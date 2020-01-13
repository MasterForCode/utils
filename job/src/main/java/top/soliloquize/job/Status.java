package top.soliloquize.job;

/**
 * @author wb
 * @date 2020/1/13
 */
public enum Status {
     FINISHED(1), RUNNING(2), AVAILABLE(3);

    private int value;

    Status(int value) {
        this.value = value;
    }

    public int getValue() {
        return value;
    }
}
