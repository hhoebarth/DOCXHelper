package com.octapharma.docx;

import java.util.List;
import javax.xml.bind.JAXBException;
import org.docx4j.XmlUtils;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.wml.P;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class WatermarkUtils {

   private static Logger logger = LoggerFactory.getLogger(WatermarkUtils.class);

   /***********************************************************************
    * Adds the water mark on document.
    *
    * @param wmlPackage the wml package
    * @param backgroundTxt the background txt
    * @throws JAXBException the JAXB exception
    * @throws Docx4JException
    **********************************************************************/
   public static void addWaterMarkOnDocument(final WordprocessingMLPackage wmlPackage, final String backgroundTxt) 
                     throws JAXBException, Docx4JException {

      List<SectionWrapper> sections = wmlPackage.getDocumentModel().getSections();

      int i = 1;
      for (SectionWrapper sec : sections) {
         logger.info(String.format("check sections for headerparts: - idx %s", i));
         if (!sec.getSectPr().getEGHdrFtrReferences().isEmpty()) {
            HeaderPart part = sec.getHeaderFooterPolicy().getFirstHeader();
            if (part == null) {
               part = sec.getHeaderFooterPolicy().getDefaultHeader();
               if (part != null) {
                  logger.info(String.format("default headerPart found - idx %s", i));
                  part.getContent().add(getPForHdr(backgroundTxt));
               }
            } else {
               logger.info(String.format("first headerPart found - idx %s", i));
               part.getContent().add(getPForHdr(backgroundTxt));
            }
         }
         i++;
      }
   }

   /**
    * Gets the p for hdr.
    *
    * @param backgroundTxt the background txt
    * @return the p for hdr
    * @throws JAXBException the JAXB exception
    */
   private static P getPForHdr(final String backgroundTxt) throws JAXBException {

      String joinToken = "&#xA;";
      int quantifier = 5;
      String txtToken = "";
      for (int i = 0; i < quantifier; i++) {
         txtToken += backgroundTxt + joinToken;
      }
      txtToken = txtToken.substring(0, txtToken.length() - joinToken.length());

      String bounds = "margin-left:-3pt;margin-top:-20pt;width:518pt;height:750pt";

      String txtColor = "#ddd8c2";

      StringBuffer xmlBuf = new StringBuffer();
      xmlBuf.append("<w:p xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ");
      xmlBuf.append("xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\">");
      xmlBuf.append("<w:pPr><w:pStyle w:val=\"Kopfzeile\"/></w:pPr><w:r><w:pict>");
      xmlBuf.append("<v:shapetype adj=\"10800\" coordsize=\"21600,21600\" id=\"_x0000_t136\" o:spt=\"136\" path=\"m@7,l@8,m@5,21600l@6,21600e\">");
      xmlBuf.append("<v:formulas><v:f eqn=\"sum #0 0 10800\"/><v:f eqn=\"prod #0 2 1\"/><v:f eqn=\"sum 21600 0 @1\"/>");
      xmlBuf.append("<v:f eqn=\"sum 0 0 @2\"/><v:f eqn=\"sum 21600 0 @3\"/><v:f eqn=\"if @0 @3 0\"/>");
      xmlBuf.append( "<v:f eqn=\"if @0 21600 @1\"/><v:f eqn=\"if @0 0 @2\"/><v:f eqn=\"if @0 @4 21600\"/><v:f eqn=\"mid @5 @6\"/><v:f eqn=\"mid @8 @5\"/>");
      xmlBuf.append("<v:f eqn=\"mid @7 @8\"/><v:f eqn=\"mid @6 @7\"/><v:f eqn=\"sum @6 0 @5\"/></v:formulas>");
      xmlBuf.append("<v:path o:connectangles=\"270,180,90,0\" o:connectlocs=\"@9,0;@10,10800;@11,21600;@12,10800\" o:connecttype=\"custom\" textpathok=\"t\"/>");
      xmlBuf.append("<v:textpath fitshape=\"t\" on=\"t\"/><v:handles>");
      xmlBuf.append( "<v:h position=\"#0,bottomRight\" xrange=\"6629,14971\"/></v:handles><o:lock shapetype=\"t\" text=\"t\" v:ext=\"edit\"/></v:shapetype>");
      xmlBuf.append("<v:shape alt=\"VERTRAULICH\" fillcolor=\"" + txtColor + "\" id=\"PowerPlusWaterMarkObject357533252\" o:allowincell=\"f\" o:spid=\"_x0000_s2115\" stroked=\"f\" ");
      xmlBuf.append("style=\"position:absolute;" + bounds + ";z-index:-251638272;mso-position-horizontal-relative:margin;mso-position-vertical-relative:margin\" type=\"#_x0000_t136\">");
      xmlBuf.append("<v:fill opacity=\"26214f\"/><v:textpath string=\"" + txtToken + "\" style=\"font-family:&quot;Calibri&quot;color:" + txtColor + ";font-size:14pt\"/>");
      xmlBuf.append("<o:lock aspectratio=\"t\" v:ext=\"edit\"/><w10:wrap anchorx=\"margin\" anchory=\"margin\"/></v:shape></w:pict></w:r></w:p>");

      return (P) XmlUtils.unmarshalString(xmlBuf.toString());
   }
}