package com.example.demo;

import javax.validation.Valid;
import javax.validation.constraints.NotNull;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

@Controller
public class MessageController {

	@Autowired
	private DocxGenerator docxGenerator;
	
	@RequestMapping(name =  "message", method = RequestMethod.POST)
	public ResponseEntity createNewDocxMessage(@RequestBody UserInformation userInformation) {
		byte[] result;
		HttpHeaders headers = new HttpHeaders();
		try {
            result = docxGenerator.generateDocxFileFromTemplate(userInformation);
            
            headers.add("Content-Disposition", "attachment; filename=\"message.docx\"");
            
		} catch (Exception e) {
			e.printStackTrace();
			return ResponseEntity.badRequest().build();
		}
		
		return ResponseEntity.ok()
	            .headers(headers)
	            .contentType(MediaType.parseMediaType("application/octet-stream"))
	            .body(result);
//		
//		return Response.ok(result, MediaType.APPLICATION_OCTET_STREAM)
//                .header("Content-Disposition", "attachment; filename=\"message.docx\"")
//                .build();
	}
	
}
