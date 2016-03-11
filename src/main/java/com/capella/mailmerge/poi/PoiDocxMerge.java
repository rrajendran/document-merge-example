package com.capella.mailmerge.poi;

import com.google.common.base.Preconditions;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * Created on : 3/8/16
 *
 * @author Ramesh Rajendran
 */
public class PoiDocxMerge {

    private static final String PREFIX = "$";

    /**
     * Document merge docx
     *
     * @param inputStream
     * @param fieldValueMap
     * @param resultFileName
     * @return
     * @throws IOException
     */
    public String merge(InputStream inputStream, Map<String, String> fieldValueMap, String resultFileName, MergeType mergeType) throws IOException {

        String mergDocPath = null;
        try {
            Preconditions.checkNotNull(inputStream);
            XWPFDocument document = new XWPFDocument(inputStream);


            fieldValueMap.forEach((k, v) -> replacePOI(document, k, v, mergeType));
            mergDocPath = saveWord(resultFileName, document);

        } catch (Exception ex) {
            throw ex;
        } finally {
            if (inputStream != null) {
                inputStream.close();
            }
        }

        return mergDocPath;
    }

    private XWPFDocument replacePOI(XWPFDocument doc, String placeHolder, String replaceText, MergeType mergeType) {
        // REPLACE ALL HEADERS
        for (XWPFHeader header : doc.getHeaderList())
            replaceAllBodyElements(header.getBodyElements(), placeHolder, replaceText, mergeType);

        // REPLACE ALL FOOTER
        for (XWPFFooter footer : doc.getFooterList()) {
            replaceAllBodyElements(footer.getBodyElements(), placeHolder, replaceText, mergeType);
        }

        // REPLACE BODY
        replaceAllBodyElements(doc.getBodyElements(), placeHolder, replaceText, mergeType);
        return doc;
    }

    private void replaceAllBodyElements(List<IBodyElement> bodyElements, String placeHolder, String replaceText, MergeType mergeType) {
        for (IBodyElement bodyElement : bodyElements) {
            if (bodyElement.getElementType().compareTo(BodyElementType.PARAGRAPH) == 0)
                replaceParagraph((XWPFParagraph) bodyElement, placeHolder, replaceText, mergeType);
            if (bodyElement.getElementType().compareTo(BodyElementType.TABLE) == 0)
                replaceTable((XWPFTable) bodyElement, placeHolder, replaceText, mergeType);
        }
    }

    private void replaceTable(XWPFTable table, String placeHolder, String replaceText, MergeType mergeType) {
        for (XWPFTableRow row : table.getRows())
            for (XWPFTableCell cell : row.getTableCells())
                for (IBodyElement bodyElement : cell.getBodyElements()) {
                    if (bodyElement.getElementType().compareTo(BodyElementType.PARAGRAPH) == 0)
                        replaceParagraph((XWPFParagraph) bodyElement, placeHolder, replaceText, mergeType);
                    if (bodyElement.getElementType().compareTo(BodyElementType.TABLE) == 0)
                        replaceTable((XWPFTable) bodyElement, placeHolder, replaceText, mergeType);
                }
    }

    private void replaceParagraph(XWPFParagraph paragraph, String placeHolder, String replaceText, MergeType mergeType) {

        for (XWPFRun r : paragraph.getRuns()) {

            if (mergeType.equals(MergeType.MAIL_MERGE)) {
                for (CTText cttext : r.getCTR().getTArray()) {
                    if (cttext.getStringValue().contains(placeHolder)) {
                        CTText ctText = CTText.Factory.newInstance();
                        ctText.setStringValue(replaceText);
                        r.getCTR().setTArray(new CTText[]{ctText});
                    }
                }
            } else {
                String text = r.getText(r.getTextPosition());
                if (text != null && text.contains(placeHolder)) {
                    text = text.replace(placeHolder, replaceText);
                    r.setText(text, 0);
                }
            }


        }

    }

    private String saveWord(String filePath, XWPFDocument doc) throws IOException {
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(filePath);
            doc.write(out);
        } finally {
            if (out != null) {
                out.close();
            }
        }
        return filePath;
    }
}
