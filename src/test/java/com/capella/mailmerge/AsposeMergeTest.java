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
        InputStream inputStream = PoiDocxMergeTest.class.getClassLoader().getResourceAsStream("tables.docx");

        String[] fieldNames = new String[] {"LastName", "FirstName", "FaxNumber", "PhoneNumber",
                "Subject", "Body", "Urgent", "ForReview", "PleaseComment"};
        Object[] fieldValues = new Object[] {"Josh", "Jenny", "123456789", "", "Hello",
                "Test message 1", true, false, true};

        AsposeMailMerge.merge(inputStream,fieldNames,fieldValues);
    }

}