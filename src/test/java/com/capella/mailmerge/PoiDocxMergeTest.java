package com.capella.mailmerge;

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
    PoiDocxMerge poiDocxMerge = new PoiDocxMerge();

    @Test
    public void testdocxMerge() throws IOException {
        long currentMilliseconds = System.currentTimeMillis();
        Map<String, String> map = ImmutableMap.<String, String>builder().
                put("{{firstName}}", "Ramesh").
                put("{{lastName}}", "Rajendran").
                build();
        InputStream resourceAsStream = PoiDocxMergeTest.class.getClassLoader().getResourceAsStream("tables.docx");
        String merge = poiDocxMerge.merge(resourceAsStream, (Map<String, String>) map, "target/poi_tables.docx");
        System.out.println("time taken : " + (System.currentTimeMillis() - currentMilliseconds) + "ms");
        System.out.println("saved at  " + merge);
    }

}