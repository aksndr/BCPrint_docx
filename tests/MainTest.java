import org.junit.Assert;
import ru.aksndr.BCPrint;

import java.io.FileNotFoundException;
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
        bcPrint.init();

        Map<String, Object> result = bcPrint.createBarcodeDocument(barCodes);
        Assert.assertNotNull(result);
        //Assert.assertFalse(result.isEmpty());
        Assert.assertTrue((Boolean) result.get("ok"));

        byte[] value = (byte[])result.get("value");

        Assert.assertNotNull(value);
        Assert.assertTrue(value.length > 0);

        FileOutputStream fos = new FileOutputStream(String.format("D:\\temp\\bcprint_docx\\%s.docx",new Date().getTime()));
        fos.write(value);
        fos.close();
    }

    private List<String> buildTestBCodes() {
        List<String> barCodes = new ArrayList<>();

        for (int i = 1; i < 43; i++) {
            String barcode = "70000005" + String.format("%012d", new Random().nextInt(999999999));
            barCodes.add(barcode);
        }
        return barCodes;
    }
}
