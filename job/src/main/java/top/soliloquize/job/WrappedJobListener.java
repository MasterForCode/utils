package top.soliloquize.job;

import org.quartz.Job;
import org.quartz.JobExecutionContext;
import org.quartz.JobExecutionException;
import org.quartz.JobListener;

import java.util.Map;

/**
 * @author wb
 * @date 2020/1/13
 */
public class WrappedJobListener implements JobListener {
    public static final String LISTENER_NAME = "simpleJobListener";

    private Map<String, Status> jobMap;

    public WrappedJobListener(Map<String, Status> jobMap) {
        this.jobMap = jobMap;
    }

    @Override
    public String getName() {
        return WrappedJobListener.LISTENER_NAME;
    }

    @Override
    public void jobToBeExecuted(JobExecutionContext context) {
        jobMap.put(context.getJobDetail().getKey().getGroup() + "-"  + context.getJobDetail().getKey().getName(), Status.AVAILABLE);
    }

    @Override
    public void jobExecutionVetoed(JobExecutionContext context) {
        jobMap.put(context.getJobDetail().getKey().getGroup() + "-" + context.getJobDetail().getKey().getName() , Status.RUNNING);
    }

    @Override
    public void jobWasExecuted(JobExecutionContext context, JobExecutionException jobException) {
        jobMap.put(context.getJobDetail().getKey().getGroup() + "-" + context.getJobDetail().getKey().getName(), Status.FINISHED);
    }
}
