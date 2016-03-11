package com.capella.mailmerge.poi;

import com.google.common.collect.ImmutableMap;
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

/**
 * Created on : 3/1/16
 *
 * @author Ramesh Rajendran
 */
public class PoiDocxMergeTest {
    public static final String TEMPLATE_DOC = "tables.docx";
    private PoiDocxMerge poiDocxMerge = new PoiDocxMerge();

    @Test
    public void testdocxMerge() throws IOException {
        long currentMilliseconds = System.currentTimeMillis();
        Map<String, String> map = ImmutableMap.<String, String>builder().
                put("Title", "Mr").
                put("FirstName", "Ramesh").
                put("LastName", new String("���".getBytes("UTF-8"))).
                build();
        InputStream resourceAsStream = PoiDocxMergeTest.class.getClassLoader().getResourceAsStream(TEMPLATE_DOC);

        String merge = poiDocxMerge.merge(resourceAsStream, (Map<String, String>) map, "target/" +TEMPLATE_DOC);

        System.out.println("time taken : " + (System.currentTimeMillis() - currentMilliseconds) + "ms");
        System.out.println("saved at  " + merge);
    }

}