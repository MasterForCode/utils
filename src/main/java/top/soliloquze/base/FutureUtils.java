package top.soliloquze.base;

import java.util.Arrays;
import java.util.List;
import java.util.Random;
import java.util.concurrent.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @author wb
 * @date 2019/12/31
 */
public class FutureUtils {

    /**
     * 出现异常立马结束所有future
     *
     * @param futures futures
     */
    public static void fastDownWhenException(List<CompletableFuture<Void>> futures) {
        CompletableFuture<Void> allWithFailFast = CompletableFuture.allOf(futures.toArray(new CompletableFuture[0]));
        futures.forEach(future -> future.exceptionally(e -> {
            allWithFailFast.completeExceptionally(e);
            return null;
        }));
    }

    /**
     * 设置超时时间
     *
     * @param future  future
     * @param timeOut 超时时间
     * @param unit    时间单位
     * @param <T>     T
     * @return CompletableFuture
     */
    public static <T> CompletableFuture<T> completeOnTimeout(CompletableFuture<T> future, long timeOut, TimeUnit unit) {
        ScheduledExecutorService e = Executors.newSingleThreadScheduledExecutor();
        ScheduledFuture<Boolean> job = e.schedule(() -> future.completeExceptionally(new TimeoutException()), timeOut, unit);
        return future.whenComplete((x, y) -> {
            job.cancel(false);
            e.shutdown();
        });
    }

    /**
     * 等待一定时间,按提交顺序获取数据
     *
     * @param futures futures
     * @param timeOut 超时时间
     * @param unit    时间单位
     * @param <T>     T
     * @return List
     */
    public static <T> List<T> allOf(List<CompletableFuture<T>> futures, long timeOut, TimeUnit unit) {
        return completeOnTimeout(
                CompletableFuture.allOf(futures.toArray(new CompletableFuture[0])), timeOut, unit)
                .thenApply(v -> futures.stream()
                        .filter(CompletableFuture::isDone)
                        .map(CompletableFuture::join)
                        .collect(Collectors.toList())).join();
    }

    public static void main(String[] args) {
        // Generate data
        int arraySize = 32768;
        int[] data = new int[arraySize];

        Random rnd = new Random(0);
        for (int c = 0; c < arraySize; ++c) {
            data[c] = rnd.nextInt() % 256;
        }

        // !!! With this, the next loop runs faster
//        Arrays.sort(data);

        // Test
        long start = System.nanoTime();
        long sum = 0;

//        for (int i = 0; i < 100000; ++i)
//        {
//            // Primary loop
//            for (int c = 0; c < arraySize; ++c)
//            {
////                if (data[c] >= 128) {
//                    sum += data[c];
////                }
//            }
//        }
//        for (int i = 0; i < 100000; ++i)
//        {
//            for (int j = 0; j < arraySize; ++j)
//            {
//                if (data[j] >= 128) {
//                    sum += data[j];
//                }
//            }
//        }
//        for (int j = 0; j < arraySize; ++j)
//        {
//            for (int i = 0; i < 100000; ++i)
//            {
//                if (data[j] >= 128) {
//                    sum += data[j];
//                }
//            }
//        }

        System.out.println((System.nanoTime() - start) / 1000000000.0);
        System.out.println("sum = " + sum);
    }
}
