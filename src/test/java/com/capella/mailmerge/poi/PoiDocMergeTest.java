package com.capella.mailmerge.poi;

import com.google.common.collect.ImmutableMap;
import org.junit.Ignore;
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

/**
 * Created on : 3/1/16
 *
 * @author Ramesh Rajendran
 */
@Ignore
public class PoiDocMergeTest {
    PoiDocMerge poiDocMerge = new PoiDocMerge();
    @Test
    public void testMerge() throws IOException {
        Map<String, String> map = ImmutableMap.<String, String>builder().
                put("$FirstName", "Renée").
                put("$LastName", "Mathew").
                put("$ImageName", "Mail Merge Demo").
                build();
        InputStream resourceAsStream = PoiDocMergeTest.class.getClassLoader().getResourceAsStream("poi.doc");
        poiDocMerge.merge(resourceAsStream, map, "target/poi_tables.doc");
    }



}