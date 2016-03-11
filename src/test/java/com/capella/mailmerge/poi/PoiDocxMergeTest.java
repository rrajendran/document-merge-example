package com.capella.mailmerge.poi;

import com.capella.mailmerge.BaseMailMergeTest;
import com.google.common.collect.ImmutableMap;
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.Map;
import java.util.stream.IntStream;

/**
 * Created on : 3/1/16
 *
 * @author Ramesh Rajendran
 */
public class PoiDocxMergeTest extends BaseMailMergeTest{

    private PoiDocxMerge poiDocxMerge = new PoiDocxMerge();

    @Test
    public void testMailMerge() throws IOException {
        System.out.println("**********POI - MAIL MERGE***************");

        Map<String, String> map = ImmutableMap.<String, String>builder().
                put("FirstName", "Renée").
                put("LastName", "Mathew").
                put("ImageField", "Mail Merge Demo").
                put("Date", "12 March 2016").
                build();

        IntStream.range(1, 10).forEach((index) -> {
            try {

                InputStream resourceAsStream = PoiDocxMergeTest.class.getClassLoader().getResourceAsStream(TEMPLATE_DOC);
                long currentMilliseconds = System.currentTimeMillis();

                String merge = poiDocxMerge.merge(resourceAsStream, map, OUTPUT_DIR + TEMPLATE_DOC, MergeType.MAIL_MERGE);

                System.out.println("time taken : " + (System.currentTimeMillis() - currentMilliseconds) + " ms");
            } catch (Exception e) {
                e.printStackTrace();
            }
        });


    }


    @Test
    public void testFindReplace() throws IOException {
        System.out.println("**********POI - FIND/REPLACE***************");

        Map<String, String> map = ImmutableMap.<String, String>builder().
                put("$FirstName", "Renée").
                put("$LastName", "Mathew").
                put("$ImageName", "Mail Merge Demo").
                put("$Date", "12 March 2016").
                build();

        IntStream.range(1, 10).forEach((index) -> {
            try {

                InputStream resourceAsStream = PoiDocxMergeTest.class.getClassLoader().getResourceAsStream("poi_2010.docx");
                long currentMilliseconds = System.currentTimeMillis();

                String merge = poiDocxMerge.merge(resourceAsStream, map, OUTPUT_DIR + "findReplace_poi.docx", MergeType.FIND_REPLACE);

                System.out.println("time taken : " + (System.currentTimeMillis() - currentMilliseconds) + " ms");
            } catch (Exception e) {
                e.printStackTrace();
            }
        });


    }

}