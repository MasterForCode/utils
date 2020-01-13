package top.soliloquize.job;

import org.quartz.SchedulerException;
import top.soliloquze.base.JsonUtil;

import java.util.HashMap;
import java.util.Map;

/**
 * @author wb
 * @date 2020/1/7
 */
public class Test {

    public static void main(String[] args) throws SchedulerException, InterruptedException {
        Map<String, String> map = new HashMap<>();
        map.put("test", "hi");
        SchedulerWrapper.INSTANCE.put(MyJob.class, JobWrapper.builder().jsonParams(JsonUtil.jsonEx(map)).defaultJobId("job1").defaultTriggerId("trigger1").cronExpression("0/2 * * * * ? ").build())
                ;
        Thread.sleep(1000);
        SchedulerWrapper.INSTANCE.put(MyJob.class, JobWrapper.builder().defaultJobId("job2").defaultTriggerId("trigger2").cronExpression("0/5 * * * * ? ").build());
//        schedulerWrapper.pause(JobWrapper.builder().build());
        Thread.sleep(10000);
        System.out.println("====");
//        schedulerWrapper.resume(JobWrapper.builder().build());
    }
}
