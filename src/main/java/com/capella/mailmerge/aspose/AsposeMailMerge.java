package com.capella.mailmerge.aspose;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FieldMergingArgs;
import com.aspose.words.IFieldMergingCallback;
import com.aspose.words.ImageFieldMergingArgs;
import com.aspose.words.TextFormFieldType;

import java.io.InputStream;
import java.util.Map;


/**
 * This sample shows how to insert check boxes and text input form fields during mail merge into a document.
 */
public class AsposeMailMerge {
    public static void findReplace(InputStream inputStream, Map<String, String> map, String outFile) throws Exception {
        Document doc = new Document(inputStream);
        map.forEach((k, v) -> {
            try {
                doc.getRange().replace(k, v, false, false);
            } catch (Exception e) {
                e.printStackTrace();
            }
        });
        doc.save(outFile);
    }

    /**
     * The main entry point for the application.
     */
    public static void merge(InputStream inputStream, String[] fieldNames, Object[] fieldValues) throws Exception {
        // The path to the documents directory.

        // Load the template document.
        Document doc = new Document(inputStream);

        // Setup mail merge event handler to do the custom work.
        doc.getMailMerge().setFieldMergingCallback(new HandleMergeField());

        // This is the data for mail merge.


        // Execute the mail merge.
        doc.getMailMerge().execute(fieldNames, fieldValues);

        // Save the finished document.
        doc.save("target/Aspose_MailMerge_.doc");

    }

    private static class HandleMergeField implements IFieldMergingCallback {
        private DocumentBuilder mBuilder;

        /**
         * This handler is called for every mail merge field found in the document,
         * for every record found in the data source.
         */
        public void fieldMerging(FieldMergingArgs e) throws Exception {
            if (mBuilder == null)
                mBuilder = new DocumentBuilder(e.getDocument());

            // We decided that we want all boolean values to be output as check box form fields.
            if (e.getFieldValue() instanceof Boolean) {
                // Move the "cursor" to the current merge field.
                mBuilder.moveToMergeField(e.getFieldName());

                // It is nice to give names to check boxes. Lets generate a name such as MyField21 or so.
                String checkBoxName = java.text.MessageFormat.format("{0}{1}", e.getFieldName(), e.getRecordIndex());

                // Insert a check box.
                mBuilder.insertCheckBox(checkBoxName, (Boolean) e.getFieldValue(), 0);


                // Nothing else to do for this field.
                return;
            }

            // Another example, we want the Subject field to come out as text input form field.
            if ("Subject".equals(e.getFieldName())) {
                mBuilder.moveToMergeField(e.getFieldName());
                String textInputName = java.text.MessageFormat.format("{0}{1}", e.getFieldName(), e.getRecordIndex());
                mBuilder.insertTextInput(textInputName, TextFormFieldType.REGULAR, "", (String) e.getFieldValue(), 0);
            }
        }

        public void imageFieldMerging(ImageFieldMergingArgs args) throws Exception {
            // Do nothing.
        }
    }
}