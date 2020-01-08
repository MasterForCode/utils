package top.soliloquize.job;

import org.quartz.*;
import org.quartz.impl.StdSchedulerFactory;
import top.soliloquze.base.JsonUtil;

import java.util.HashMap;
import java.util.Objects;

/**
 * @author wb
 * @date 2020/1/7
 */
public enum  SchedulerWrapper {

    /**
     * 全局唯一实例
     */
    INSTANCE;

    private static Scheduler scheduler;

    static {
        try {
            scheduler = StdSchedulerFactory.getDefaultScheduler();
        } catch (SchedulerException e) {
            throw new RuntimeException(e);
        }
    }

    public SchedulerWrapper put(Class<? extends Job> clz, JobWrapper jobWrapper) {
        Objects.requireNonNull(clz);
        if (jobWrapper == null) {
            jobWrapper = JobWrapper.builder().build();
        }
        JobBuilder jobBuilder = JobBuilder.newJob(clz).withIdentity(jobWrapper.getDefaultJobId(), jobWrapper.getDefaultGroupId());
        JobDetail job;
        if (jobWrapper.getJsonParams() == null) {
            job = JobBuilder.newJob(clz).withIdentity(jobWrapper.getDefaultJobId(), jobWrapper.getDefaultGroupId()).build();
        } else {
            JobDataMap jobDataMap = new JobDataMap(JsonUtil.getBeanFromJsonEx(jobWrapper.getJsonParams(), HashMap.class));
            job = jobBuilder.setJobData(jobDataMap).build();
        }
        Trigger trigger;
        if (jobWrapper.getCronExpression() != null && jobWrapper.getCronExpression().length() > 0) {
            trigger = TriggerBuilder.newTrigger().withIdentity(jobWrapper.getDefaultTriggerId(), jobWrapper.getDefaultGroupId()).withSchedule(CronScheduleBuilder.cronSchedule(jobWrapper.getCronExpression())).startNow().build();
        } else {
            trigger = TriggerBuilder.newTrigger().withIdentity(jobWrapper.getDefaultTriggerId(), jobWrapper.getDefaultGroupId()).startNow().build();
        }
        try {
            scheduler.scheduleJob(job, trigger);
            scheduler.start();
        } catch (SchedulerException e) {
            throw new RuntimeException(e);
        }
        return this;
    }

    public SchedulerWrapper pauseAll() {
        try {
            scheduler.pauseAll();
        } catch (SchedulerException e) {
            throw new RuntimeException();
        }
        return this;
    }

    public SchedulerWrapper resumeAll() {
        try {
            scheduler.resumeAll();
        } catch (SchedulerException e) {
            throw new RuntimeException();
        }
        return this;
    }

    public SchedulerWrapper pause(JobWrapper jobWrapper) {
        Objects.requireNonNull(jobWrapper);
        try {
            scheduler.pauseJob(JobKey.jobKey(jobWrapper.getDefaultJobId(), jobWrapper.getDefaultGroupId()));
        } catch (SchedulerException e) {
            throw new RuntimeException(e);
        }
        return this;
    }

    public SchedulerWrapper resume(JobWrapper jobWrapper) {
        Objects.requireNonNull(jobWrapper);
        try {
            scheduler.resumeJob(JobKey.jobKey(jobWrapper.getDefaultJobId(), jobWrapper.getDefaultGroupId()));
        } catch (SchedulerException e) {
            throw new RuntimeException(e);
        }
        return this;
    }

}
