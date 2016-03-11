package com.capella.mailmerge.aspose;

import com.capella.mailmerge.poi.PoiDocxMergeTest;
import com.google.common.collect.ImmutableMap;
import org.junit.Test;

import java.io.InputStream;
import java.util.Map;
import java.util.stream.IntStream;

/**
 * Created on : 3/9/16
 *
 * @author Ramesh Rajendran
 */
public class AsposeMergeTest {
    private static final String TEMPLATE_DOC = "poi_2010.docx";

    @Test
    public void testMerge() throws Exception {
        System.out.println("**********ASPOSE-MAIL MERGE***************");
        String[] fieldNames = new String[]{"LastName", "FirstName", "ImageField"};
        Object[] fieldValues = new Object[]{"Josh", "Renée", "Mail Merge Demo"};
        IntStream.range(1, 10).forEach((index) -> {
            try {
                long currentMilliseconds = System.currentTimeMillis();
                InputStream inputStream = PoiDocxMergeTest.class.getClassLoader().getResourceAsStream("aspose.docx");
                AsposeMailMerge.merge(inputStream, fieldNames, fieldValues);
                System.out.println("time taken : " + (System.currentTimeMillis() - currentMilliseconds) + " ms");
            } catch (Exception e) {
                e.printStackTrace();
            }
        });

    }


    @Test

    public void testFindReplace() throws Exception {


        Map<String, String> map = ImmutableMap.<String, String>builder().
                put("$FirstName", "Renée").
                put("$LastName", "Mathew").
                put("$ImageName", "Mail Merge Demo").
                put("$Date", "12 March 2016").
                build();
        InputStream resourceAsStream = PoiDocxMergeTest.class.getClassLoader().getResourceAsStream(TEMPLATE_DOC);

        System.out.println("**********ASPOSE-FIND/REPLACE***************");
        IntStream.range(1, 10).forEach((index) -> {
            try {
                long currentMilliseconds = System.currentTimeMillis();
                AsposeMailMerge.findReplace(resourceAsStream, map, "target/findReplace_" + TEMPLATE_DOC);
                System.out.println("time taken : " + (System.currentTimeMillis() - currentMilliseconds) + " ms");
            } catch (Exception e) {
                e.printStackTrace();
            }
        });


    }

}