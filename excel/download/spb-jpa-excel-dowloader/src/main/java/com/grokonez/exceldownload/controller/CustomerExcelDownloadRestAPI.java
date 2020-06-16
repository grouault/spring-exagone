package com.grokonez.exceldownload.controller;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.http.ResponseEntity.BodyBuilder;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.grokonez.exceldownload.model.Customer;
import com.grokonez.exceldownload.repository.CustomerRepository;
import com.grokonez.exceldownload.util.ExcelGenerator;

@RestController
@RequestMapping("/api/customers")
public class CustomerExcelDownloadRestAPI {
    @Autowired
    CustomerRepository customerRepository;
 
    @GetMapping(value = "/download/customers.xlsx")
    public ResponseEntity<InputStreamResource> excelCustomersReport() throws IOException {
        List<Customer> customers = (List<Customer>) customerRepository.findAll();
		
		ByteArrayInputStream in = ExcelGenerator.customersToExcel(customers);
		// return IOUtils.toByteArray(in);

		// headers
		// - content-type
		// - content-disposition
		HttpHeaders headers = new HttpHeaders();
        headers.add("Content-Disposition", "attachment; filename=customers.xlsx");
        String mimeType = "application/vmd.openxmlformats-officedocument.wordprocessingml.document";
        MediaType mediaType = MediaType.parseMediaType(mimeType);
        
        
        BodyBuilder response = ResponseEntity.ok();
        response.contentType(mediaType);
        
		 return response
	                .headers(headers)
	                .body(new InputStreamResource(in));
    }
}
