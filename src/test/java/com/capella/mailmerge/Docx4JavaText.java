package com.capella.mailmerge;

import org.junit.Test;

import java.io.InputStream;

/**
 * Created on : 3/9/16
 *
 * @author Ramesh Rajendran
 */
public class Docx4JavaText {
    @Test
    public void testMerge() throws Exception {
        InputStream inputStream = PoiDocxMergeTest.class.getClassLoader().getResourceAsStream("docx4java.docx");
        FieldsMailMerge.merge(inputStream);
    }

}