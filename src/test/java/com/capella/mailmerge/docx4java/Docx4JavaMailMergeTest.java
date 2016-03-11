package com.capella.mailmerge.docx4java;

import com.capella.mailmerge.poi.PoiDocxMergeTest;
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.junit.Test;

import javax.xml.bind.JAXBException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.IntStream;

/**
 * Created on : 3/9/16
 *
 * @author Ramesh Rajendran
 */
public class Docx4JavaMailMergeTest {
    @Test
    public void testMerge() throws Exception {
        System.out.println("**********DOCX4JAVA***************");


        List<Map<DataFieldName, String>> data = new ArrayList<Map<DataFieldName, String>>();

        // Instance 1
        Map<DataFieldName, String> map = new HashMap<DataFieldName, String>();
        map.put( new DataFieldName("FirstName"), "Daffy duck");
        map.put(new DataFieldName("LastName"), "Renée");
        map.put(new DataFieldName("ImageField"), "Mail Merge Demo");

        data.add(map);

        IntStream.range(1, 10).forEach((index) -> {
            try {

                long currentMilliseconds = System.currentTimeMillis();
                InputStream inputStream = PoiDocxMergeTest.class.getClassLoader().getResourceAsStream("aspose.docx");
                Docx4JavaMailMerge.merge(inputStream, data);
                System.out.println("time taken : " + (System.currentTimeMillis() - currentMilliseconds) + " ms");
            } catch (Exception e) {
                e.printStackTrace();
            }
        });


    }

    @Test
    public void testFindReplace() throws JAXBException, Docx4JException {


        HashMap<String, String> map = new HashMap<>();
        map.put("$FirstName", "Renée");
        map.put("$LastName", "Mathew");
        map.put("$ImageName", "Mail Merge Demo");
        map.put("$Date", "12 March 2016");

        System.out.println("**********DOCX4JAVA - FIND/REPLACE***************");
        IntStream.range(1, 10).forEach((index) -> {
            try {

                long currentMilliseconds = System.currentTimeMillis();
                InputStream inputStream = PoiDocxMergeTest.class.getClassLoader().getResourceAsStream("poi_2010.docx");
                Docx4JavaMailMerge.merge(inputStream, map);
                System.out.println("time taken : " + (System.currentTimeMillis() - currentMilliseconds) + " ms");
            } catch (Exception e) {
                e.printStackTrace();
            }
        });


    }

}