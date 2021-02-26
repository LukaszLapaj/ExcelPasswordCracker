import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.IOException;
import java.time.Duration;
import java.time.Instant;
import java.time.LocalDateTime;
import java.util.concurrent.*;

public class Main {

    final static File inputFile = new File("input.xlsx");

    public static void main(String[] args) {
        CompletableFuture<Integer> cf = new CompletableFuture<>();
        BlockingQueue<Integer> queue = new ArrayBlockingQueue<>(16);

        POIFSFileSystem fileSystem;
        EncryptionInfo info;
        Decryptor decryptor = null;

        try {
            fileSystem = new POIFSFileSystem(inputFile);
            info = new EncryptionInfo(fileSystem);
            decryptor = Decryptor.getInstance(info);
        } catch (IOException e) {
            e.printStackTrace();
        }

        Runnable producer = new Runnable() {
            final static int MAX_SIZE = Integer.MAX_VALUE;

            @Override
            public void run() {
                for (int i = 0; i < MAX_SIZE; ++i) {
                    if (!Thread.interrupted()) {
                        try {
//                        queue.offer(i, 10000, TimeUnit.SECONDS);
                            queue.put(i);
                        } catch (InterruptedException e) {
                            break;
//                            e.printStackTrace();
                        }
                    }
                    print("Adding number: " + i);

                }
            }
        };

        Decryptor finalDecryptor = decryptor;
        Runnable consumer = new Runnable() {
            final Decryptor d = finalDecryptor;

            @Override
            public void run() {
                while (!Thread.interrupted()) {
                    try {
                        int number = queue.take();
                        print("Read number: " + number);
                        Instant start = Instant.now();
                        d.verifyPassword(String.valueOf(number));
                        Instant finish = Instant.now();
                        long timeElapsed = Duration.between(start, finish).toMillis();
                        System.out.println(timeElapsed);
                        if (d.verifyPassword(String.valueOf(number))) {
                            cf.complete(number);
                            break;
                        }
                    } catch (Exception e) {
                        break;
                    }
                }

            }
        };

        int cores = Runtime.getRuntime().availableProcessors();

        ExecutorService service = new ThreadPoolExecutor(cores / 2, cores,
                0L, TimeUnit.MILLISECONDS,
                new LinkedBlockingQueue<Runnable>());

        service.execute(producer);

        for (int i = 0; i < cores - 1; ++i) {
            service.execute(consumer);

        }
//        service.execute(consumer);

        Integer result = 0;
        try {
            Instant start = Instant.now();
            result = cf.get();
            Instant finish = Instant.now();
            long timeElapsed = Duration.between(start, finish).toMillis();
            System.out.println("Łączny czas: " + timeElapsed);
            service.shutdownNow();
        } catch (InterruptedException | ExecutionException e) {
            e.printStackTrace();
        }
        System.out.println(result);

//        Set<Thread> threadSet = Thread.getAllStackTraces().keySet();
//        System.out.println(threadSet);

//        System.out.println(Runtime.getRuntime().availableProcessors());
//        shutdownAndAwaitTermination(service);
    }

    static void print(Object output) {
        System.out.println(
                String.format("%s: %s - %s", LocalDateTime.now(), Thread.currentThread().getName(), output)
        );
    }
}
