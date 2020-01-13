package top.soliloquize.job;

import org.quartz.Job;
import org.quartz.JobExecutionContext;

import java.util.Calendar;

/**
 * @author wb
 * @date 2020/1/7
 */
public class MyJob implements Job {
    @Override
    public void execute(JobExecutionContext context) {
        System.out.println(context.getJobDetail().getKey().getName() + "任务正在执行，执行时间: " + Calendar.getInstance().getTime().getSeconds());
    }
}
