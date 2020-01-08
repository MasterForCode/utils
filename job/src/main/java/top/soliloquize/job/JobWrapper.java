package top.soliloquize.job;

import lombok.Builder;
import lombok.Data;

/**
 * @author wb
 * @date 2020/1/8
 */
@Data
@Builder
public class JobWrapper {
    @Builder.Default
    private String defaultGroupId = "group";
    @Builder.Default
    private String defaultJobId = "job";
    @Builder.Default
    private String defaultTriggerId = "trigger";
    private String cronExpression;
    private String jsonParams;
}
