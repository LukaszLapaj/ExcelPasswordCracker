import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.time.Duration;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.concurrent.*;
import java.util.stream.IntStream;

public class Main {
    final static File inputFile = new File("input.xlsx");

    public static void main(String[] args) {
        // Character set for cracking
        ArrayList<Character> characters = new ArrayList<>();
        // Numbers
        IntStream.range(48, 58).forEach(i -> characters.add((char) i));
        // Lower case
        IntStream.range(65, 91).forEach(i -> characters.add((char) i));
        // Upper case
        IntStream.range(97, 123).forEach(i -> characters.add((char) i));
        Character[] charSet = getCharSet(characters);

        // Parameters
        final int minPasswordLength = 2;
        final int maxPasswordLength = Integer.MAX_VALUE;

        // Thread pool
        int threadCount = Runtime.getRuntime().availableProcessors();
        ExecutorService threadPoolExecutor = new ThreadPoolExecutor(threadCount, threadCount, 0L, TimeUnit.MILLISECONDS, new LinkedBlockingQueue<>());
        CompletableFuture<String> cf = new CompletableFuture<>();
        BlockingQueue<String> passwordQueue = new LinkedBlockingQueue<>(threadCount * 2);

        // Runnables to execute
        final Decryptor excelDecryptor = getDecryptor();
        Runnable producer = passwordProvider(charSet, passwordQueue, minPasswordLength, maxPasswordLength);
        Runnable consumer = passwordCracker(cf, passwordQueue, excelDecryptor);

        executeRunnableDesiredTimes(1, threadPoolExecutor, producer);
        executeRunnableDesiredTimes(threadCount - 1, threadPoolExecutor, consumer);

        String result = crackPassword(threadPoolExecutor, cf);
        System.out.println("Password found: " + result);
    }

    private static String crackPassword(ExecutorService service, CompletableFuture<String> cf) {
        String result = "";
        try {
            Instant start = Instant.now();
            result = cf.get();
            Instant finish = Instant.now();
            long timeElapsed = Duration.between(start, finish).toMillis();
            System.out.println("Total time elapsed: " + timeElapsed);
            service.shutdownNow();
        } catch (InterruptedException | ExecutionException e) {
            e.printStackTrace();
        }
        return result;
    }

    private static void executeRunnableDesiredTimes(int times, ExecutorService service, Runnable runnable) {
        for (int i = 0; i < times; ++i) {
            service.execute(runnable);
        }
    }

    private static Runnable passwordProvider(Character[] charSet, BlockingQueue<String> passwordQueue, int minLen, int maxLen) {
        int charSetSize = charSet.length;
        return () -> {
            generatorLoop:
            for (int i = minLen; i < maxLen; ++i) {
                ArrayList<String> passwords = new ArrayList<>();
                passwords.addAll(generatePasswordsForDesiredLength(charSet, i, "", charSetSize));
                for (int j = 0; j < passwords.size(); j++) {
                    String pass = passwords.get(j);
                    try {
                        passwordQueue.offer(pass, 12, TimeUnit.HOURS);
                    } catch (InterruptedException ignored) {
                        break generatorLoop;
                    }
                }
            }
        };
    }

    private static Runnable passwordCracker(CompletableFuture<String> cf, BlockingQueue<String> passwordQueue, Decryptor excelDecryptor) {
        return () -> {
            while (!Thread.interrupted()) {
                try {
                    String password = passwordQueue.take();
                    print("Testing password: " + password);
                    boolean decryptResult = excelDecryptor.verifyPassword(password);
                    if (decryptResult) {
                        cf.complete(password);
                        break;
                    }
                } catch (InterruptedException | GeneralSecurityException ignored) {
                    break;
                }
            }
        };
    }

    private static Decryptor getDecryptor() {
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
        return finalDecryptor;
    }

    private static Character[] getCharSet(ArrayList<Character> letters) {
        Character[] chars = new Character[letters.size()];
        for (int i = 0; i < chars.length; ++i) {
            chars[i] = letters.get(i);
        }
        return chars;
    }

    static ArrayList<String> generatePasswordsForDesiredLength(Character[] arr, int i, String s, int length) {
        ArrayList<String> passwords = new ArrayList<>();
        if (i == 0) {
            passwords.add(s);
            return passwords;
        }
        for (int j = 0; j < length; j++) {
            String appended = s + arr[j];
            passwords.addAll(generatePasswordsForDesiredLength(arr, i - 1, appended, length));
        }
        return passwords;
    }

    static void print(Object output) {
        String formattedTime = DateTimeFormatter.ofPattern("MM-dd HH:mm:ss.SSSSSSS")
                .withZone(ZoneId.systemDefault())
                .format(Instant.now());
        System.out.printf("%s: %ss - %s%n", formattedTime, Thread.currentThread().getName(), output);
    }
}
