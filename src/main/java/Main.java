import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.IOException;
import java.time.Duration;
import java.time.Instant;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.concurrent.*;
import java.util.stream.IntStream;

public class Main {

    final static File inputFile = new File("input.xlsx");

    public static void main(String[] args) {
        int cores = Runtime.getRuntime().availableProcessors();

        ExecutorService service = new ThreadPoolExecutor(cores, cores,
                0L, TimeUnit.MILLISECONDS,
                new LinkedBlockingQueue<Runnable>());

        CompletableFuture<String> cf = new CompletableFuture<>();
        BlockingQueue<String> queue = new ArrayBlockingQueue<>(cores * 2);

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

        ArrayList<Character> characters = new ArrayList<>();
        // Numbers
//        IntStream.range(48, 57).forEach(i -> characters.add((char) i));
        // Lower case
        IntStream.range(65, 90).forEach(i -> characters.add((char) i));
        // Upper case
        IntStream.range(97, 122).forEach(i -> characters.add((char) i));
        
        Character[] charSet = getCharSet(characters);
        int charSetSize = charSet.length;

        final int min_len = 2;
        final int max_len = Integer.MAX_VALUE;

        Runnable producer = new Runnable() {
            @Override
            public void run() {
//                for (int i = 0; i < charSet.length && !Thread.interrupted(); ++i) {
                outerloop:
                for (int i = min_len; i < max_len; ++i) {
                    ArrayList<String> t = new ArrayList<>();
                    t.addAll(generate(charSet, i, "", charSetSize));
                    for (int j = 0; j < t.size(); j++) {
                        String pass = t.get(j);
                        try {
                            queue.offer(pass, 10000, TimeUnit.SECONDS);
                        } catch (InterruptedException ignored) {
                            break outerloop;
                        }
                    }
                }
            }
        };

        Runnable consumer = new Runnable() {
            final Decryptor d = finalDecryptor;

            @Override
            public void run() {
                while (!Thread.interrupted()) {
                    try {
                        String password = queue.take();
                        print("Read number: " + password);
                        Instant start = Instant.now();
                        d.verifyPassword(password);
                        Instant finish = Instant.now();
                        long timeElapsed = Duration.between(start, finish).toMillis();
//                        System.out.println(timeElapsed);
                        if (d.verifyPassword(password)) {
                            cf.complete(password);
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

        String result = "";
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

    private static Character[] getCharSet(ArrayList<Character> letters) {
        Character[] chars = new Character[letters.size()];
        for (int i = 0; i < chars.length; ++i) {
            chars[i] = letters.get(i);
        }
        return chars;
    }

    static ArrayList<String> generate(Character[] arr, int i, String s, int len) {
        ArrayList<String> passwords = new ArrayList<>();
        if (i == 0) {
            passwords.add(s);
            return passwords;
        }
        for (int j = 0; j < len; j++) {
            String appended = s + arr[j];
            passwords.addAll(generate(arr, i - 1, appended, len));
        }
        return passwords;
    }

    static void print(Object output) {
        System.out.printf("%s: %s - %s%n", LocalDateTime.now(), Thread.currentThread().getName(), output);
    }
}
