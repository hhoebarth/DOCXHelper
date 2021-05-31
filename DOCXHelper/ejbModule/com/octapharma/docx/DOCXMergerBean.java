package com.octapharma.docx;

import javax.annotation.PostConstruct;
import javax.ejb.Stateless;

import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.Collections;
import java.util.List;

import org.apache.commons.io.IOUtils;
import org.apache.log4j.Level;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.log4j.varia.NullAppender;

import org.docx4j.Docx4J;
import org.docx4j.Docx4jProperties;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.toc.TocGenerator;

import com.aspose.words.DocSaveOptions;
import com.aspose.words.Document;
import com.aspose.words.FileFontSource;
import com.aspose.words.FolderFontSource;
import com.aspose.words.FontInfo;
import com.aspose.words.FontSettings;
import com.aspose.words.FontSourceBase;
import com.aspose.words.HeaderFooter;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.HorizontalAlignment;
import com.aspose.words.LoadOptions;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.Paragraph;
import com.aspose.words.PdfCompliance;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.SaveFormat;
import com.aspose.words.SaveOptions;
import com.aspose.words.Section;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.words.SystemFontSource;
import com.aspose.words.VerticalAlignment;
import com.aspose.words.WrapType;
import com.sap.engine.services.webservices.espbase.configuration.ann.dt.AuthenticationDT;
import com.sap.engine.services.webservices.espbase.configuration.ann.dt.AuthenticationEnumsAuthenticationLevel;
import com.sap.tc.logging.Category;
import com.sap.tc.logging.Location;
import com.sap.tc.logging.MsgObject;
import com.sap.tc.logging.Severity;

import javax.jws.WebService;

/**
 * Session Bean implementation class DOCXMergerBean
 */
@WebService(endpointInterface = "com.octapharma.docx.DOCXMergerBeanLocal", portName = "DOCXMergerBeanPort", serviceName = "DOCXMergerService", targetNamespace = "http://octapharma.com/docx/")
@Stateless
@AuthenticationDT(authenticationLevel=AuthenticationEnumsAuthenticationLevel.BASIC)
public class DOCXMergerBean implements DOCXMergerBeanRemote, DOCXMergerBeanLocal {

	private static final Category category = Category.getCategory(Category.APPLICATIONS, "DOCXMergerBean");
	
	private com.aspose.words.License license;
	
	final Location logger = Location.getLocation(DOCXMergerBean.class);
	
    /**
     * Default constructor. 
     */
    public DOCXMergerBean() {
        // TODO Auto-generated constructor stub
    }
    
    @PostConstruct
    public void installLicense() {
    	this.license = __installLicense();
    }
    
    @Override    
    public byte[] mergeDOCX(byte[] docx, byte[] xml, String watermarkText)
    {	
    	Logger.getRootLogger().removeAllAppenders();
    	Logger.getRootLogger().addAppender(new NullAppender());
    	
    	List<Logger> loggers = Collections.<Logger>list(LogManager.getCurrentLoggers());
    	loggers.add(LogManager.getRootLogger());
    	for ( Logger logger : loggers ) {
    	    logger.setLevel(Level.OFF);
    	}
    	
    	try {
    		category.infoT(logger, "mergeDOCX", "method entered", null);
    		
	    	if(docx.length == 0) {
	    		category.infoT(logger, "mergeDOCX", "one or more input parameters null", null);
		    	return null;
	    	}
	    	
	    	WordprocessingMLPackage mergedDocx = null;
	    	if(xml.length == 0) {
	    		category.infoT(logger, "mergeDOCX", "parameter xml empty", null);
	    		mergedDocx = Docx4J.load(new ByteArrayInputStream(docx));
	    	}
	    	else
	    	{	
	    		mergedDocx = _mergeDocxWithXML(docx, xml);
	   
	    		// Watermark via DOCX4J
	    		if(!__isLicensed())
	    			WatermarkUtils.addWaterMarkOnDocument(mergedDocx, watermarkText);
	    		
	    		ByteArrayOutputStream result = new ByteArrayOutputStream();
	    		Docx4J.save(mergedDocx, result, Docx4J.FLAG_SAVE_ZIP_FILE);
	    	
	    		category.infoT(logger, "mergeDOCX", "document merged", null);
	    	    if(__isLicensed())
	    	    	// Watermark via Aspose
	    	    	return watermarkText != null?watermarkText.isEmpty()?result.toByteArray():_addWatermarkAspose(result.toByteArray(), watermarkText):result.toByteArray();
	    	    else
	    	    	return result.toByteArray();
	    	}
	    	
	    	if(mergedDocx == null) {
	    		category.infoT(logger, "mergeDOCX", "merged document null", null);
		    	return null;
	    	}
    	}
		catch(Throwable th){
			category.logThrowable(Severity.ERROR, logger, new MsgObject(th.getClass().getName(), th.getMessage()), th);
			th.printStackTrace();
			return null;
		}
    	
		return null;
    }

	protected WordprocessingMLPackage _mergeDocxWithXML(byte[] docx, byte[] xml) throws Throwable {
		InputStream docxStream = new ByteArrayInputStream(docx);
		InputStream xmlStream = new ByteArrayInputStream(xml);
			
		// Load input docx
		WordprocessingMLPackage wordMLPackage = Docx4J.load(docxStream);
			
		_mergeDocxPackageWithXML(wordMLPackage, xmlStream);

		return wordMLPackage;
	}
	
	protected byte[] _mergeDocxWithXMLToByteStream(byte[] docx, byte[] xml) throws Throwable {
		InputStream docxStream = new ByteArrayInputStream(docx);
		InputStream xmlStream = new ByteArrayInputStream(xml);
			
		ByteArrayOutputStream result = new ByteArrayOutputStream();
			
		// Load input docx
		WordprocessingMLPackage wordMLPackage = Docx4J.load(docxStream);
			
		_mergeDocxPackageWithXML(wordMLPackage, xmlStream);
				
		//Save the document 
		Docx4J.save(wordMLPackage, result, Docx4J.FLAG_SAVE_ZIP_FILE);
				
		return result.toByteArray();
	}
	
	protected WordprocessingMLPackage _mergeDocxPackageWithXML(WordprocessingMLPackage docx, InputStream xml) throws Throwable {
		// Do the binding:
		// FLAG_NONE means that all the steps of the binding will be done,
		// otherwise you could pass a combination of the following flags:
		// FLAG_BIND_INSERT_XML: inject the passed XML into the document
		// FLAG_BIND_BIND_XML: bind the document and the xml (including any OpenDope handling)
		// FLAG_BIND_REMOVE_SDT: remove the content controls from the document (only the content remains)
		// FLAG_BIND_REMOVE_XML: remove the custom xml parts from the document 
										
		//Docx4J.bind(wordMLPackage, xmlStream, Docx4J.FLAG_NONE);
		//If a document doesn't include the Opendope definitions, eg. the XPathPart,
		//then the only thing you can do is insert the xml
		//the example document binding-simple.docx doesn't have an XPathPart....
		Docx4J.bind(docx, xml, Docx4J.FLAG_BIND_INSERT_XML | Docx4J.FLAG_BIND_BIND_XML);
		
		// Update TOC
		try {
			TocGenerator tocGenerator = new TocGenerator(docx);
	
			Docx4jProperties.setProperty("docx4j.toc.BookmarksIntegrity.remediate", true);
			tocGenerator.updateToc(false);
		}
		catch (Exception e){}
		
		Docx4J.bind(docx, xml, Docx4J.FLAG_BIND_REMOVE_SDT | Docx4J.FLAG_BIND_REMOVE_XML);
		
		return docx;
	}
	
	protected byte[] _convertDocxToPDF(byte[] input) {
		InputStream docxStream = new ByteArrayInputStream(input);
		ByteArrayOutputStream result = new ByteArrayOutputStream();
			
		// Load input docx
		WordprocessingMLPackage wordMLPackage;
		try {
			wordMLPackage = Docx4J.load(docxStream);
			
			//Save the document 
			Docx4J.toPDF(wordMLPackage, result);
		} catch (Docx4JException e) {
			category.logThrowable(Severity.ERROR, logger, new MsgObject(e.getClass().getName(), e.getMessage()), e);
		}
		//Docx4J.save(wordMLPackage, result, Docx4J.FLAG_NONE);
				
		return result.toByteArray();
	}

	protected byte[] _convertDocxPackageToPDF(WordprocessingMLPackage input) throws Throwable {
		ByteArrayOutputStream result = new ByteArrayOutputStream();
			
		//Save the document 
		Docx4J.toPDF(input, result);
		//Docx4J.save(wordMLPackage, result, Docx4J.FLAG_NONE);
				
		return result.toByteArray();
	}
	
	protected Document _convertToAsposeDoc(byte[] docx) throws Exception {
		category.infoT(logger, "convertToAsposeDoc", "method entered", null);
		
		return new Document(new ByteArrayInputStream(docx));
	}
	
	protected byte[] _convertToByteArray(Document doc, SaveOptions options) throws Exception {
		ByteArrayOutputStream result = new ByteArrayOutputStream();
		doc.save(result, options);
		
		return result.toByteArray();
	}
	
	protected byte[] _addWatermarkAspose(byte[] docx, String watermarkText) throws Exception {
		Document doc = _convertToAsposeDoc(docx);
		
	    // Create a watermark shape. This will be a WordArt shape.
	    // You are free to try other shape types as watermarks.
	    Shape watermark = new Shape(doc, ShapeType.TEXT_PLAIN_TEXT);

	    // Set up the text of the watermark.
	    watermark.getTextPath().setText(watermarkText);
	    watermark.getTextPath().setFontFamily("Arial");
	    watermark.setWidth(500);
	    watermark.setHeight(100);

	    // Text will be directed from the bottom-left to the top-right corner.
	    watermark.setRotation(-40);

	    // Remove the following two lines if you need a solid black text.
	    watermark.getFill().setColor(Color.GRAY);
	    // Try LightGray to get more Word-style watermark
	    watermark.setStrokeColor(Color.GRAY);
	    // Try LightGray to get more Word-style watermark
	    // Place the watermark in the page center.
	    watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
	    watermark.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
	    watermark.setWrapType(WrapType.NONE);
	    watermark.setVerticalAlignment(VerticalAlignment.CENTER);
	    watermark.setHorizontalAlignment(HorizontalAlignment.CENTER);
	    //watermark.setBehindText(true);

	    // Create a new paragraph and append the watermark to this paragraph.
	    Paragraph watermarkPara = new Paragraph(doc);
	    watermarkPara.getParagraphBreakFont().setSize(0);
	    watermarkPara.appendChild(watermark);

	    // Insert the watermark into all headers of each document section.
	    for (Section sect : doc.getSections())
	    {
	        // There could be up to three different headers in each section, since we want
	        // the watermark to appear on all pages, insert into all headers.
	        __insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_PRIMARY);

	        //__insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_FIRST);

	        //__insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_EVEN);
	    }
		
		return _convertToByteArray(doc, new OoxmlSaveOptions(SaveFormat.DOCX));
	}

	private void __insertWatermarkIntoHeader(Paragraph watermarkPara, Section sect, int headerType) throws Exception
	{
	    HeaderFooter header = sect.getHeadersFooters().getByHeaderFooterType(headerType);
	    if (header == null)
	    {
	        // There is no header of the specified type in the current section, create it.
	        header = new HeaderFooter(sect.getDocument(), headerType);
	        sect.getHeadersFooters().add(header);
	    }

	    // Insert a clone of the watermark into the header.
	    header.appendChild(watermarkPara.deepClone(true));
	}

	@Override
	public byte[] mergeDOCXasPDF(byte[] docx, byte[] xml, String watermarkText) {
		category.infoT(logger, "mergeDOCXasPDF", "method entered", null);

		return convertDOCXtoPDF(mergeDOCX(docx, xml, watermarkText));
	}

	@Override
	public byte[] convertDOCXtoPDF(byte[] docx) {
		category.infoT(logger, "convertDOCXtoPDF", "method entered", null);
		
		if(__isLicensed())
			// PDF via Aspose
			return _convertDOCXtoPDFAspose(docx);
		else 
			// PDF via DOCX4J
			return _convertDocxToPDF(docx);
	}
	
	protected byte[] _convertDOCXtoPDFAspose(byte[] docx) {
		category.infoT(logger, "convertDOCXtoPDFAspose", "method entered", null);

		try {
			FontSettings.setFontsSources(new FontSourceBase[] {new SystemFontSource(), new FolderFontSource("/usr/sap/cust/Fonts", true)});
			//FontSettings.setFontsSources(new FontSourceBase[] {/*new SystemFontSource(),*/ new FileFontSource("/tmp/times.ttf"), new FileFontSource("/tmp/arial.ttf"), new FileFontSource("/tmp/wingding.ttf")});
			category.infoT(logger, "convertDOCXtoPDFAspose", "Default font: " + FontSettings.getDefaultFontName(), null);
			
			//__installLicense();
			Document doc = _convertToAsposeDoc(docx);
			for(FontInfo fontInfo:doc.getFontInfos())
				category.infoT(logger, "convertDOCXtoPDFAspose", "Used font: " + fontInfo.getName(), null);
			
			//FontSettings.setFontsSources(new FontSourceBase[] {/*new SystemFontSource(),*/ new FolderFontSource("/usr/sap/DPO/J46/j2ee/os_libs/adssap/FontManagerService/fonts", true)});
			
			// set PDF format
			PdfSaveOptions options = new PdfSaveOptions();
			options.setCompliance(PdfCompliance.PDF_A_1_B);
			options.getOutlineOptions().setHeadingsOutlineLevels(5);
			options.getOutlineOptions().setExpandedOutlineLevels(1);

			return _convertToByteArray(doc, options);
		} catch (Exception e) {
			category.logThrowable(Severity.ERROR, logger, new MsgObject(e.getClass().getName(), e.getMessage()), e);
			return null;
		}
	}
	
	/**
	 * install Aspose License
	 */
	@SuppressWarnings("deprecation")
	private com.aspose.words.License __installLicense() {
		category.infoT(logger, "__installLicense", "method entered", null);
		
		try {
			// Word license
			com.aspose.words.License wordLicense = new com.aspose.words.License();
			
			//FileInputStream stream = new FileInputStream(new File("com/octapharma/docx/Aspose.Words.lic"));
			wordLicense.setLicense(IOUtils.toInputStream("<License>\r\n" + 
					"  <Data>\r\n" + 
					"    <LicensedTo>DHC Dr. Herterich ^ Consultants GmbH</LicensedTo>\r\n" + 
					"    <EmailTo>it@dhc-gmbh.com</EmailTo>\r\n" + 
					"    <LicenseType>Developer OEM</LicenseType>\r\n" + 
					"    <LicenseNote>Limited to 1 developer, unlimited physical locations</LicenseNote>\r\n" + 
					"    <OrderID>181206052004</OrderID>\r\n" + 
					"    <UserID>97830</UserID>\r\n" + 
					"    <OEM>This is a redistributable license</OEM>\r\n" + 
					"    <Products>\r\n" + 
					"      <Product>Aspose.Words for Java</Product>\r\n" + 
					"    </Products>\r\n" + 
					"    <EditionType>Enterprise</EditionType>\r\n" + 
					"    <SerialNumber>aac95b74-9d95-43d5-9986-55c98e48e5ea</SerialNumber>\r\n" + 
					"    <SubscriptionExpiry>20191206</SubscriptionExpiry>\r\n" + 
					"    <LicenseVersion>3.0</LicenseVersion>\r\n" + 
					"    <LicenseInstructions>https://purchase.aspose.com/policies/use-license</LicenseInstructions>\r\n" + 
					"  </Data>\r\n" + 
					"  <Signature>vti70WUZu5R4KBBmqAFEFCG4T33qVIejKgoFt+MmH6CCAvQa4ZWlHqgU8+BFZQ9gr75wafZ1QLzQ2niuJJFqNnExh27lEee69h4wBW/+tUDJf5z7vG+iBe76mQVbAV5EJ/9c9vXBQ5IWz0U5rM4u/dr+1TgZLB+XIq8lvzonwVo=</Signature>\r\n" + 
					"</License>"));
			return wordLicense;
		} catch (Exception e) {
			category.logThrowable(Severity.ERROR, logger, new MsgObject(e.getClass().getName(), e.getMessage()), e);
			return null;
		}
	}
	
	private boolean __isLicensed() {
		return this.license != null;
	}
}
