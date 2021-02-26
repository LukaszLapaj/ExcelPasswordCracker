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
        int cores = Runtime.getRuntime().availableProcessors();

        ExecutorService service = new ThreadPoolExecutor(cores / 2, cores,
                0L, TimeUnit.MILLISECONDS,
                new LinkedBlockingQueue<Runnable>());

        CompletableFuture<Integer> cf = new CompletableFuture<>();
        BlockingQueue<Integer> queue = new ArrayBlockingQueue<>(cores * 2);

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

        Decryptor finalDecryptor = decryptor;

        Runnable producer = new Runnable() {
            final static int MAX_SIZE = Integer.MAX_VALUE;

            @Override
            public void run() {
                for (int i = 0; i < MAX_SIZE; ++i) {
                    if (!Thread.interrupted()) {
                        try {
                            queue.put(i);
                        } catch (InterruptedException e) {
                            break;
                        }
                    }
                    print("Adding number: " + i);

                }
            }
        };

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

        service.execute(producer);
        for (int i = 0; i < cores - 1; ++i) {
            service.execute(consumer);

        }

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
    }

    static void print(Object output) {
        System.out.printf("%s: %s - %s%n", LocalDateTime.now(), Thread.currentThread().getName(), output);
    }
}
