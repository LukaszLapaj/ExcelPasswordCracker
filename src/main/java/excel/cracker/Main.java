package excel.cracker;

import java.io.File;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;
import java.util.stream.IntStream;

import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Main {

	private static final Logger log = LoggerFactory.getLogger(Main.class);

	static File inputFile;

	public static void main(String[] args) throws IOException {

		if (args.length <= 0) {
			log.error("File can't be null");
			return;
		}
		inputFile = new File(args[0]);

		// Character set for cracking
		final ArrayList<Character> characters = new ArrayList<>();
		// Numbers
		IntStream.range(48, 58).forEach(i -> characters.add((char) i));
		// Lower case
		IntStream.range(65, 91).forEach(i -> characters.add((char) i));
		// Upper case
		IntStream.range(97, 123).forEach(i -> characters.add((char) i));

		final Character[] charSet = characters.toArray(new Character[characters.size()]);

		// Parameters
		final int minPasswordLength = 4;
		final int maxPasswordLength = Integer.MAX_VALUE;

		// Thread pool
		final int threadCount = Runtime.getRuntime().availableProcessors();
		final ExecutorService threadPoolExecutor = new ThreadPoolExecutor(threadCount, threadCount, 0L,
				TimeUnit.MILLISECONDS, new LinkedBlockingQueue<>());
		final CompletableFuture<String> cf = new CompletableFuture<>();
		final BlockingQueue<String> passwordQueue = new LinkedBlockingQueue<>(threadCount * 2);

		// Runnables to execute
		final Decryptor excelDecryptor = Decryptor.getInstance(new EncryptionInfo(new POIFSFileSystem(inputFile)));
		final Runnable producer = passwordProvider(charSet, passwordQueue, minPasswordLength, maxPasswordLength);
		final Runnable consumer = passwordCracker(cf, passwordQueue, excelDecryptor);

		executeRunnableDesiredTimes(1, threadPoolExecutor, producer);
		executeRunnableDesiredTimes(threadCount - 1, threadPoolExecutor, consumer);

		final String result = crackPassword(threadPoolExecutor, cf);
		log.info("Password found: {}", result);
		System.out.println("Password found: " + result);
	}

	private static String crackPassword(ExecutorService service, CompletableFuture<String> cf) {
		String result = "";
		try {
			final Instant start = Instant.now();
			result = cf.get();
			final Instant finish = Instant.now();
			final long timeElapsed = Duration.between(start, finish).toMillis();
			log.info("Total time elapsed: {}", timeElapsed);
			service.shutdownNow();
		} catch (InterruptedException | ExecutionException e) {
			log.error("crackPassword : {}", e.getMessage(), e);
		}
		return result;
	}

	private static void executeRunnableDesiredTimes(int times, ExecutorService service, Runnable runnable) {
		IntStream.range(0, times).forEach(value -> service.execute(runnable));
	}

	private static Runnable passwordProvider(Character[] charSet, BlockingQueue<String> passwordQueue, int minLen,
			int maxLen) {
		final int charSetSize = charSet.length;
		return () -> {
			generatorLoop: for (int i = minLen; i < maxLen; ++i) {
				final ArrayList<String> passwords = new ArrayList<>();
				passwords.addAll(generatePasswordsForDesiredLength(charSet, i, "", charSetSize));
				for (int j = 0; j < passwords.size(); j++) {
					final String pass = passwords.get(j);
					try {
						passwordQueue.offer(pass, 12, TimeUnit.HOURS);
					} catch (final InterruptedException e) {
						log.error("passwordProvider : {}", e.getMessage(), e);
						break generatorLoop;
					}
				}
			}
		};
	}

	private static Runnable passwordCracker(CompletableFuture<String> cf, BlockingQueue<String> passwordQueue,
			Decryptor excelDecryptor) {
		return () -> {
			while (!Thread.interrupted()) {
				try {
					final String password = passwordQueue.take();
					log.info("Testing password: {}", password);
					final boolean decryptResult = excelDecryptor.verifyPassword(password);
					if (decryptResult) {
						cf.complete(password);
						break;
					}
				} catch (InterruptedException | GeneralSecurityException e) {
					log.error("passwordCracker : {}", e.getMessage(), e);
					break;
				}
			}
		};
	}

	static ArrayList<String> generatePasswordsForDesiredLength(Character[] arr, int i, String s, int length) {
		final ArrayList<String> passwords = new ArrayList<>();
		if (i == 0) {
			passwords.add(s);
			return passwords;
		}
		for (int j = 0; j < length; j++) {
			final String appended = s + arr[j];
			passwords.addAll(generatePasswordsForDesiredLength(arr, i - 1, appended, length));
		}
		return passwords;
	}

}
