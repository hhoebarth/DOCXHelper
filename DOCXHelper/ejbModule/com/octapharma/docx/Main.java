package com.octapharma.docx;

import java.io.File;
import java.io.FileInputStream;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class Main {
    public static void main( String[] args )
    {
    	try {
	    	String input_DOCX = "C:/041VBE004-01-Tes1t.docx";
			
			// the instance data
			String input_XML = "C:/Test.xml";
					
			// resulting docx
			String OUTPUT_DOCX = "C:/VBE-MERGED.docx";
	
			// Load input_template.docx
			WordprocessingMLPackage wordMLPackage = Docx4J.load(new File(input_DOCX));
	
			// Open the xml stream
			FileInputStream xmlStream = new FileInputStream(new File(input_XML));
			
			DOCXMergerBean bean = new DOCXMergerBean();
			byte[] mergedResult = bean.mergeDOCX(IOUtils.toByteArray(new FileInputStream(new File(input_DOCX))), IOUtils.toByteArray(new FileInputStream(new File(input_XML))), null);
			
			FileUtils.writeByteArrayToFile(new File(OUTPUT_DOCX), mergedResult);
    	} catch (Throwable e) {
            e.printStackTrace();
        }
    }
}
