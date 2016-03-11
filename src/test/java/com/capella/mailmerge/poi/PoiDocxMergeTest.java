package com.capella.mailmerge.poi;

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
public class PoiDocxMergeTest {
    public static final String TEMPLATE_DOC = "poi_2010.docx";
    private PoiDocxMerge poiDocxMerge = new PoiDocxMerge();

    @Test
    public void testdocxMerge() throws IOException {
        System.out.println("**********POI***************");

        Map<String, String> map = ImmutableMap.<String, String>builder().
                put("$FirstName", "Renée").
                put("$LastName", "Mathew").
                put("$ImageName", "Mail Merge Demo").
                put("$Date", "12 March 2016").
                build();

        IntStream.range(1, 10).forEach((index) -> {
            try {

                InputStream resourceAsStream = PoiDocxMergeTest.class.getClassLoader().getResourceAsStream(TEMPLATE_DOC);
                long currentMilliseconds = System.currentTimeMillis();

                String merge = poiDocxMerge.merge(resourceAsStream, map, "target/" + TEMPLATE_DOC);

                System.out.println("time taken : " + (System.currentTimeMillis() - currentMilliseconds) + " ms");
            } catch (Exception e) {
                e.printStackTrace();
            }
        });


    }

}