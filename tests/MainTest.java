import org.junit.Assert;
import com.aksndr.BCPrint;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * Created by Aksndr on 28.02.2016.
 */
public class MainTest {

    @org.junit.Test
    public void test1() throws IOException {
        List<String> barCodes = buildTestBCodes();

        BCPrint bcPrint = new BCPrint();

        Map<String, Object> result = bcPrint.getSheet(barCodes);
        Assert.assertNotNull(result);
        //Assert.assertFalse(result.isEmpty());
        Assert.assertTrue((Boolean) result.get("ok"));

        byte[] value = (byte[])result.get("value");

        Assert.assertNotNull(value);
        Assert.assertTrue(value.length > 0);

        FileOutputStream fos = new FileOutputStream(String.format("D:\\temp\\bcprint\\%s.docx", new Date().getTime()));
        fos.write(value);
        fos.close();
    }

    @org.junit.Test
    public void test2() throws IOException {
        List<String> barCodes = buildTestBCodes( new Random().nextInt(999));

        BCPrint bcPrint = new BCPrint();

        Map<String, Object> result = bcPrint.getSheet(barCodes);
        Assert.assertNotNull(result);
        //Assert.assertFalse(result.isEmpty());
        Assert.assertTrue((Boolean) result.get("ok"));

        byte[] value = (byte[])result.get("value");

        Assert.assertNotNull(value);
        Assert.assertTrue(value.length > 0);

        FileOutputStream fos = new FileOutputStream(String.format("D:\\temp\\bcprint\\%s.docx", new Date().getTime()));
        fos.write(value);
        fos.close();
    }

    @org.junit.Test
    public void test3() throws IOException {
        test2();
        test2();
        test2();
        test2();
    }

    private List<String> buildTestBCodes() {
        List<String> barCodes = new ArrayList<>();

        for (int i = 1; i < 43; i++) {
            String barcode = "70000005" + String.format("%012d", new Random().nextInt(999999999));
            barCodes.add(barcode);
        }
        return barCodes;
    }

    private List<String> buildTestBCodes(int qty) {
        List<String> barCodes = new ArrayList<>();

        for (int i = 1; i < qty; i++) {
            String barcode = "70000005" + String.format("%012d", i);
            barCodes.add(barcode);
        }
        return barCodes;
    }
}
