package com.capella.mailmerge.docx4java;

import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.model.fields.merge.MailMerger.OutputField;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import javax.xml.bind.JAXBException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Example of how to process MERGEFIELD.
 * 
 * See http://webapp.docx4java.org/OnlineDemo/ecma376/WordML/MERGEFIELD.html
 *
 */
public class Docx4JavaMailMerge {

	public static final String OUT_FILE = "target/docx4java_mail_merge.docx";

	public static void merge(InputStream inputStream, HashMap<String, String> keyValues) throws Docx4JException, JAXBException {
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(inputStream);
		wordMLPackage.getMainDocumentPart().variableReplace(keyValues);

		wordMLPackage.save(new java.io.File("target/docx4java_find_replace.docx"));
	}
	public static void merge(InputStream inputStream, List<Map<DataFieldName, String>> data ) throws Exception {

		// Whether to create a single output docx, or a docx per Map of input data.
		// Note: If you only have 1 instance of input data, then you can just invoke performMerge
		boolean mergedOutput = true;
		
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(inputStream);
		

		
		
		if (mergedOutput) {
			/*
			 * This is a "poor man's" merge, which generates the mail merge  
			 * results as a single docx, and just hopes for the best.
			 * Images and hyperlinks should be ok. But numbering 
			 * will continue, as will footnotes/endnotes.
			 *  
			 * If your resulting documents aren't opening in Word, then
			 * you probably need MergeDocx to perform the merge.
			 */

			// How to treat the MERGEFIELD, in the output?
			org.docx4j.model.fields.merge.MailMerger.setMERGEFIELDInOutput(OutputField.KEEP_MERGEFIELD);
			
//			System.out.println(XmlUtils.marshaltoString(wordMLPackage.getMainDocumentPart().getJaxbElement(), true, true));
			
			WordprocessingMLPackage output = org.docx4j.model.fields.merge.MailMerger.getConsolidatedResultCrude(wordMLPackage, data, true);
			
//			System.out.println(XmlUtils.marshaltoString(output.getMainDocumentPart().getJaxbElement(), true, true));

			output.save(new java.io.File(OUT_FILE));
			
		} else {
			// Need to keep the MERGEFIELDs. If you don't, you'd have to clone the docx, and perform the
			// merge on the clone.  For how to clone, see the MailMerger code, method getConsolidatedResultCrude
			org.docx4j.model.fields.merge.MailMerger.setMERGEFIELDInOutput(OutputField.KEEP_MERGEFIELD);
			
			int i = 1;
			for (Map<DataFieldName, String> thismap : data) {
				org.docx4j.model.fields.merge.MailMerger.performMerge(wordMLPackage, thismap, true);
				org.docx4j.model.fields.merge.MailMerger.performMerge(wordMLPackage, thismap, true);
				wordMLPackage.save(new java.io.File(OUT_FILE));
				i++;
			}			
		}
		
	}

}