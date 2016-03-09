package com.capella.mailmerge;

import org.junit.Test;

import java.io.InputStream;

/**
 * Created on : 3/9/16
 *
 * @author Ramesh Rajendran
 */
public class AsposeMergeTest {
    @Test
    public void testMerge() throws Exception {
        InputStream inputStream = PoiDocxMergeTest.class.getClassLoader().getResourceAsStream("Template.doc");
        MailMergeFormFields.merge(inputStream);
    }

}