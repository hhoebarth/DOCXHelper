package com.octapharma.docx;

import javax.ejb.Local;
import javax.jws.WebService;
import javax.jws.WebMethod;
import javax.jws.WebParam;

@WebService(name = "DOCXMergerBeanLocal", targetNamespace = "http://octapharma.com/docx/")
@Local
public interface DOCXMergerBeanLocal {

	@WebMethod(operationName = "mergeDOCX")
	public byte[] mergeDOCX (@WebParam(name = "docx") byte[] docx, @WebParam(name = "xml") byte[] xml, @WebParam(name = "watermarkText") String watermarkText);

	@WebMethod(operationName = "mergeDOCXAsPDF")
	public byte[] mergeDOCXasPDF (@WebParam(name = "docx") byte[] docx, @WebParam(name = "xml") byte[] xml, @WebParam(name = "watermarkText") String watermarkText);

	@WebMethod(operationName = "convertDOCXtoPDF")
	public byte[] convertDOCXtoPDF (@WebParam(name = "docx") byte[] docx);
}
