package co.com.novaip.tecuido;

import java.time.LocalDate;
import java.util.Calendar;
import java.util.Date;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl;
import microsoft.exchange.webservices.data.autodiscover.exception.AutodiscoverLocalException;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.LogicalOperator;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

@SpringBootApplication
public class TestCorreo2ReadEmailsApplication {

	public static void main(String[] args) throws Exception {

		ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);

		//credentials for login to your exchange email example: email       password
		ExchangeCredentials credentials = new WebCredentials("your-username", "your-password-here");
		service.setCredentials(credentials);
		try {
			//works fine using autodiscovers instead writting the hostname
			//example email@domain.com
			service.autodiscoverUrl("your-usaername-email-complete", new IAutodiscoverRedirectionUrl() {
				@Override
				public boolean autodiscoverRedirectionUrlValidationCallback(String url)
						throws AutodiscoverLocalException {
					return url.toLowerCase().startsWith("https://");
				}
			});
		} catch (Exception e) {
			//error loging to the email server
			System.out.println("error de inicio de sesion: " + e);
		}

		ItemView view = new ItemView(2000);
		
		//getting inbox folder
		Folder folder = Folder.bind(service, WellKnownFolderName.Inbox);
		
		//filter for unread emails
		SearchFilter unreadFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
	            new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, Boolean.FALSE));

		
		Calendar cal = Calendar.getInstance();
	    cal.add(Calendar.HOUR, -24);
		
	    //creating parent filter. You can add multiple filters for example by date, bysender, etc.
		SearchFilter.SearchFilterCollection search = new SearchFilter.SearchFilterCollection();
	    
		//adding unread filter
		search.add(new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, cal.getTime()));
	    
		//adding date filter
		search.add(unreadFilter);
		
		
		//setting the filters in the ItemsResults
		FindItemsResults<Item> findResults = service.findItems(folder.getId(), search,view);
		

		// MOOOOOOST IMPORTANT: load messages' properties before
		service.loadPropertiesForItems(findResults, PropertySet.FirstClassProperties);
		
		
		//iterating the emails
		for (Item item: findResults.getItems() ) {
			
				System.out.println("-----------------------------------"+"CORREO NUMERO: ------------------------------------------------");
			// Do something with the item as shown
			System.out.println("id==========" + item.getId().getUniqueId());
			System.out.println("sub==========" + item.getSubject());
			System.out.println("get is new==========" + item.getIsNew());
			System.out.println("new==========" + item.isNew());
			System.out.println("time created==========" + item.getDateTimeCreated() );
			System.out.println("importancia==========" + item.getImportance());
			System.out.println("recibido==========" + item.getDateTimeReceived());
			System.out.println("body==========" + item.getBody());
			System.out.println("headers" + item.getInternetMessageHeaders().getItems().toString());
			EmailMessage a = (EmailMessage)item;
			
			
			System.out.println("quien lo envio objeto" + a.getSender().getAddress());
			
			System.out.println("-----------------------------------------------------------------------------------");
			
			
			
			
		}
	

//		SpringApplication.run(TestCorreo2ReadEmailsApplication.class, args);
	}

}
