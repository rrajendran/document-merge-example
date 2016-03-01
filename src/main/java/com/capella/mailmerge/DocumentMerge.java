package com.capella.mailmerge;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

/**
 * Document Merge - word
 */
public class DocumentMerge {

    /**
     *
     * @param sourceDocument
     * @param fieldValueMap
     * @return
     * @throws IOException
     */
    public String merge(String sourceDocument, Map<String,String> fieldValueMap) throws IOException {
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(sourceDocument);
            return merge(inputStream, fieldValueMap);

        } catch(Exception ex) {
           ex.printStackTrace();
        }finally {
            if(inputStream != null){
                inputStream.close();
            }
        }

        return null;
    }

    /**
     *
     * @param inputStream
     * @param fieldValueMap
     * @return
     * @throws IOException
     */
    public String merge(InputStream inputStream, Map<String, String> fieldValueMap) throws IOException {
        HWPFDocument doc = new HWPFDocument(inputStream);
        doc = replaceText(doc, fieldValueMap);
        return saveWord("target/sample.doc", doc);
    }

    private HWPFDocument replaceText(HWPFDocument doc, Map<String,String> fieldValueMap){
        fieldValueMap.forEach((k,v) -> replaceText(doc,k,v));
        return doc;
    }

    private HWPFDocument replaceText(HWPFDocument doc, String findText, String replaceText){
        Range r1 = doc.getRange(); 

        for (int i = 0; i < r1.numSections(); ++i ) { 
            Section s = r1.getSection(i); 
            for (int x = 0; x < s.numParagraphs(); x++) { 
                Paragraph p = s.getParagraph(x); 
                for (int z = 0; z < p.numCharacterRuns(); z++) { 
                    CharacterRun run = p.getCharacterRun(z); 
                    String text = run.text();
                    if(text.contains(findText)) {
                        run.replaceText(findText, replaceText);
                    }
                }
            }
        } 
        return doc;
    }

    private static String saveWord(String filePath, HWPFDocument doc) throws FileNotFoundException, IOException{
        FileOutputStream out = null;
        try{
            out = new FileOutputStream(filePath);
            doc.write(out);
        }
        finally{
            out.close();
        }
        return filePath;
    }
}