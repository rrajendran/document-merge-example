package com.capella.mailmerge;

import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.FileNotFoundException;
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
public class DocxMerge {




    public String merge(InputStream inputStream, Map<String,String> fieldValueMap, String resultFileName) throws IOException {

        String mergDocPath = null;
        try {
            XWPFDocument document = new XWPFDocument(inputStream);
            fieldValueMap.forEach((k,v) ->  replacePOI(document, k, v));
            mergDocPath = saveWord(resultFileName, document);

        } catch(Exception ex) {
           throw ex;
        }finally {
            if(inputStream != null){
                inputStream.close();
            }
        }

        return mergDocPath;
    }

    private XWPFDocument replacePOI(XWPFDocument doc, String placeHolder, String replaceText){
        System.out.println(placeHolder + "," + replaceText);
        // REPLACE ALL HEADERS
        for (XWPFHeader header : doc.getHeaderList())
            replaceAllBodyElements(header.getBodyElements(), placeHolder, replaceText);

        // REPLACE ALL FOOTER
        for(XWPFFooter footer : doc.getFooterList()){
            replaceAllBodyElements(footer.getBodyElements(), placeHolder, replaceText);
        }

        // REPLACE BODY
        replaceAllBodyElements(doc.getBodyElements(), placeHolder, replaceText);
        return doc;
    }

    private void replaceAllBodyElements(List<IBodyElement> bodyElements, String placeHolder, String replaceText){
        for (IBodyElement bodyElement : bodyElements) {
            if (bodyElement.getElementType().compareTo(BodyElementType.PARAGRAPH) == 0)
                replaceParagraph((XWPFParagraph) bodyElement, placeHolder, replaceText);
            if (bodyElement.getElementType().compareTo(BodyElementType.TABLE) == 0)
                replaceTable((XWPFTable) bodyElement, placeHolder, replaceText);
        }
    }

    private void replaceTable(XWPFTable table, String placeHolder, String replaceText) {
        for (XWPFTableRow row : table.getRows())
            for (XWPFTableCell cell : row.getTableCells())
                for (IBodyElement bodyElement : cell.getBodyElements()) {
                    if (bodyElement.getElementType().compareTo(BodyElementType.PARAGRAPH) == 0)
                        replaceParagraph((XWPFParagraph) bodyElement, placeHolder, replaceText);
                    if (bodyElement.getElementType().compareTo(BodyElementType.TABLE) == 0)
                        replaceTable((XWPFTable) bodyElement, placeHolder, replaceText);
                }
    }

    private void replaceParagraph(XWPFParagraph paragraph, String placeHolder, String replaceText) {
        for (XWPFRun r : paragraph.getRuns()) {
            String text = r.getText(r.getTextPosition());
            if (text != null && text.contains(placeHolder)) {
                text = text.replace(placeHolder, replaceText);
                r.setText(text, 0);
            }
        }
    }

    private String saveWord(String filePath, XWPFDocument doc) throws FileNotFoundException, IOException{
        FileOutputStream out = null;
        try{
            out = new FileOutputStream(filePath);
            doc.write(out);
        }
        finally{
            if (out != null) {
                out.close();
            }
        }
        return filePath;
    }
}