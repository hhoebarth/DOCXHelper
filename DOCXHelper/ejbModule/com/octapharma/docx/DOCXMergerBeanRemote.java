package com.octapharma.docx;

import javax.ejb.Remote;

@Remote
public interface DOCXMergerBeanRemote {

	public byte[] mergeDOCX (byte[] docx, byte[] xml, String watermarkText);
	
	public byte[] mergeDOCXasPDF (byte[] docx, byte[] xml, String watermarkText);

	public byte[] convertDOCXtoPDF (byte[] docx);
}
