package com.example.spring.ews.demo;

import java.net.URI;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.jboss.logging.Logger;
import org.springframework.stereotype.Service;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.WebProxy;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;

@Service
public class MSExchangeEmailService {

	private static Logger logger = Logger.getLogger(MSExchangeEmailService.class);
	
	private static ExchangeService service;
	
	// only latest 10 emails/appointments are fetched.
    private static Integer NUMBER_EMAILS_FETCH = 10; 
    
    /**
     * Firstly check, whether "https://webmail.xxxx.com/ews/Services.wsdl" and "https://webmail.xxxx.com/ews/Exchange.asmx"
     * is accessible, if yes that means the Exchange Webservice is enabled on your MS Exchange.
     */
    static {
        try {
        	logger.info("Connecting...");
        	
            service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
        	service.setUrl(new URI("https://outlook.office365.com/ews/exchange.asmx"));
        } catch (Exception e) {
        	logger.error(e);
        }
    }
    
    /**
     * Initialize the Exchange Credentials.
     * Don't forget to replace the "USRNAME","PWD","DOMAIN_NAME" variables.
     */
    public MSExchangeEmailService() {
    	// Set proxy, if any.
    	WebProxy proxy = new WebProxy("PROXY_NAME", 80);
    	service.setWebProxy(proxy);
    	
        ExchangeCredentials credentials = new WebCredentials("USERNAME", "PWD", "DOMAIN_NAME");
        service.setCredentials(credentials);
    }
    
	/**
	 * Number of email we want to read is defined as NUMBER_EMAILS_FETCH, 
	 */
    public List<Map<String, String>> readEmails() {
        List<Map<String, String>> msgDataList = new ArrayList<>();
        
        try {
            Folder folder = Folder.bind(service, WellKnownFolderName.Inbox);
            FindItemsResults<Item> results = service.findItems(folder.getId(), new ItemView(NUMBER_EMAILS_FETCH));
            int i = 1;
            for (Item itemObj : results) {
                Map<String, String> messageData = readEmailItem(itemObj.getId());
                logger.info("\nEmails #" + (i++) + ":");
                logger.info("subject : " + messageData.get("subject").toString());
                logger.info("Sender : " + messageData.get("senderName").toString());
                msgDataList.add(messageData);
            }
        } catch (Exception e) {
        	logger.error(e);
        }
        
        return msgDataList;
    }

    /**
     * Reading one email at a time. Using Item ID of the email.
     * Creating a message data map as a return value.
     */
    public Map<String, String> readEmailItem(ItemId itemId) {
        Map<String, String> messageData = new HashMap<>();
        try {
            Item itemObj = Item.bind(service, itemId, PropertySet.FirstClassProperties);
            
            EmailMessage emailMessage = EmailMessage.bind(service, itemObj.getId());
            messageData.put("emailItemId", emailMessage.getId().toString());
            messageData.put("subject", emailMessage.getSubject().toString());
            messageData.put("fromAddress", emailMessage.getFrom().getAddress().toString());
            messageData.put("senderName", emailMessage.getSender().getName().toString());
            messageData.put("SendDate", emailMessage.getDateTimeCreated().toString());
            messageData.put("RecievedDate", emailMessage.getDateTimeReceived().toString());
            messageData.put("Size", emailMessage.getSize() + "");
            messageData.put("emailBody", emailMessage.getBody().toString());
        } catch (Exception e) {
            logger.error(e);
        }
        
        return messageData;
    }
    
}
