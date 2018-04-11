package com.example.spring.ews.demo;

import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/ews")
public class MSExchangeEmailController {
	
	@Autowired
	MSExchangeEmailService emailService;
	
	@RequestMapping(value="/v1/readEmails", method=RequestMethod.GET)
    public List<Map<String, String>> readEmails() {
		return emailService.readEmails();
    }
}
