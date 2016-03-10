package com.capella.mailmerge;

import org.docx4j.model.fields.merge.DataFieldName;
import org.junit.Test;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created on : 3/9/16
 *
 * @author Ramesh Rajendran
 */
public class Docx4JavaMailMergeTest {
    @Test
    public void testMerge() throws Exception {
        InputStream inputStream = PoiDocxMergeTest.class.getClassLoader().getResourceAsStream("tables.docx");

        List<Map<DataFieldName, String>> data = new ArrayList<Map<DataFieldName, String>>();

        // Instance 1
        Map<DataFieldName, String> map = new HashMap<DataFieldName, String>();
        map.put( new DataFieldName("FirstName"), "Daffy duck");
        map.put( new DataFieldName("LastName"), "Plutext");

        data.add(map);

        Docx4JavaMailMerge.merge(inputStream,data);
    }

}