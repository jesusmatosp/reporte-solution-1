package com.example.demo;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.util.HashMap;

import org.docx4j.Docx4J;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.springframework.stereotype.Service;

@Service
public class DocxGenerator {

	private static final String TEMPLATE_NAME = "template.docx";
	
	public byte[] generateDocxFileFromTemplate(UserInformation userInformation) throws Exception {
		
		InputStream templateInputStream = 
				this.getClass().getClassLoader().getResourceAsStream(TEMPLATE_NAME);
		WordprocessingMLPackage wordprocessingMLPackage = WordprocessingMLPackage.load(templateInputStream);
		
		MainDocumentPart documentPart = wordprocessingMLPackage.getMainDocumentPart();
		
		VariablePrepare.prepare(wordprocessingMLPackage);
		
		HashMap<String, String> variable = new HashMap<String, String>();
		variable.put("firstName", userInformation.getFirstName());
		variable.put("lastName", userInformation.getLastName());
		variable.put("salutation", userInformation.getSalutation());
		variable.put("message", userInformation.getMessage());
		
		documentPart.variableReplace(variable);
		
		// Add Sign:
		File image = new File("C:\\temp\\signature.jpg");
		byte[] fileContent = Files.readAllBytes(image.toPath());
		BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordprocessingMLPackage,
				fileContent);
		Inline inline = imagePart.createImageInline("Mit Solution image", "Alt Text", 1, 2, false);
		P Imageparagraph = addImageToParagraph(inline);
		documentPart.getContent().add(Imageparagraph);
		// end..
		
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		wordprocessingMLPackage.save(outputStream);
		
		// Convert PDF
		savePDF(wordprocessingMLPackage);
		
		return outputStream.toByteArray();
		
	}
	
	public P addImageToParagraph(Inline inline) {
		ObjectFactory factory = new ObjectFactory();
		P p = factory.createP();
		R r = factory.createR();
		p.getContent().add(r);
		Drawing drawing = factory.createDrawing();
		r.getContent().add(drawing);
		drawing.getAnchorOrInline().add(inline);
		return p;
	}
	
	public File savePDF(WordprocessingMLPackage wordprocessingMLPackage) throws Exception {
		String nameFile = "reporte-mit";
		File file = new File("C:\\temp\\" +nameFile + ".pdf");
		
		OutputStream os = new FileOutputStream(file);
		Docx4J.toPDF(wordprocessingMLPackage, os);
		
		os.flush();
		os.close();
		
		return file;
	}
	
}
