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
public class DocumentMergeTest {
    DocumentMerge documentMerge = new DocumentMerge();
    DocxMerge docxMerge = new DocxMerge();
    @Test
    public void testMerge() throws IOException {
        Map<String, String> map = ImmutableMap.<String, String>builder().
                put("$firstName", "Ramesh Rajendran").
                put("$dob", "01 Jan, 1970").
                build();
        InputStream resourceAsStream = DocumentMergeTest.class.getClassLoader().getResourceAsStream("tables.doc");
        documentMerge.merge(resourceAsStream, (Map<String, String>) map, "target/tables.doc");
    }


    @Test
    public void testdocxMerge() throws IOException {
        Map<String, String> map = ImmutableMap.<String, String>builder().
                put("$firstName", "Ramesh Rajendran").
                put("$dob", "01 Jan, 1970").
                build();
        InputStream resourceAsStream = DocumentMergeTest.class.getClassLoader().getResourceAsStream("tables.docx");
        String merge = docxMerge.merge(resourceAsStream, (Map<String, String>) map, "target/tables.docx");

        System.out.println("saved at  " + merge);
    }

    /*@Test
    public void testMergeImageDoc() throws IOException {
        Map<String, String> map = ImmutableMap.<String, String>builder().
                put("$name", "Ramesh Rajendran").
                build();
        InputStream resourceAsStream = DocumentMergeTest.class.getClassLoader().getResourceAsStream("Images.doc");
        documentMerge.merge(resourceAsStream, (Map<String, String>) map, "target/Images.doc");
    }*/
}