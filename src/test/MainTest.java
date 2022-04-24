import org.junit.Assert;
import org.junit.Test;

import java.io.File;


public class MainTest {
    @Test
    public void givenSystemOutRedirection_whenInvokePrintln_thenOutputCaptorSuccess() {
        final File inputFile = new File("src/test/resources/test_file.xlsx");
        String crackedPassword = Main.crackPassword(inputFile);
        Assert.assertEquals("abc", crackedPassword);
    }
}