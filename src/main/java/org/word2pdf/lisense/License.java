package org.word2pdf.lisense;

import java.io.InputStream;


public class License {
	
	public static boolean getWordLicense() {
        boolean result = false;
        try {
            InputStream is = License.class.getClassLoader().getResourceAsStream("license.xml"); 
            com.aspose.words.License aposeLic = new com.aspose.words.License();
            aposeLic.setLicense(is);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

}
