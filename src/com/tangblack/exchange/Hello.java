package com.tangblack.exchange;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;

import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
import microsoft.exchange.webservices.data.core.enumeration.property.LegacyFreeBusyStatus;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.LogicalOperator;
import microsoft.exchange.webservices.data.core.enumeration.search.SortDirection;
import microsoft.exchange.webservices.data.core.enumeration.service.ConflictResolutionMode;
import microsoft.exchange.webservices.data.core.enumeration.service.DeleteMode;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.CalendarFolder;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.AppointmentSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.search.CalendarView;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

public class Hello
{
	/**
	 * Main.
	 *
	 * @param args N/A.
	 * @throws Exception 
	 */
	public static void main(String[] args) throws Exception
	{
		System.out.println("Initilate exchange service...");
		ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
//		ExchangeCredentials credentials = new WebCredentials("chief.tang@itri.org.tw", "Etd2GewC");
		ExchangeCredentials credentials = new WebCredentials(User.ID, User.PASSWORD);
//		ExchangeCredentials credentials = new WebCredentials("itri\\A30183", "Etd2GewC");
		service.setCredentials(credentials);
		service.autodiscoverUrl(User.MAIL, new RedirectionUrlCallback());
		
		
		System.out.println("Create");
		createAppointment(service, "週會", "2016-05-15 09:00:00", "2016-05-15 10:00:00", "討論專案進度");
		createAppointment(service, "xx專案kick-off會議", "2016-05-16 09:00:00", "2016-05-16 10:00:00", "如標題所示，啟動xx專案。");
		createAppointment(service, "下午茶", "2016-05-17 09:00:00", "2016-05-17 10:00:00", "如標題所示");
		createAppointment(service, "foo", "2016-05-17 09:00:00", "2016-05-17 10:00:00", "I am bar's father");
		List<Appointment> appointments = findAppointments(service, "2016-05-01 00:00:00", "2016-05-30 00:00:00");
		for (Appointment appointment : appointments)
		{
			print(appointment);
		}
		
		
		System.out.println("Search");
		searchAppointments1(service, "專案");
		searchAppointments1(service, "標題");
		searchAppointments1(service, "bar");
		
		
		System.out.println("Update");
		for (Appointment appointment : appointments)
		{
			updateAppointment(service, appointment.getId().toString());
		}
		appointments = findAppointments(service, "2016-05-01 00:00:00", "2016-05-30 00:00:00");
		for (Appointment appointment : appointments)
		{
			print(appointment);
		}
		
		
		System.out.println("Delete");
		for (Appointment appointment : appointments)
		{
			deleteAppointment(service, appointment.getId().toString());
		}
	}
	
	private static List<Appointment> findAppointments(ExchangeService service, String inputStartDate, String inputEndDate) throws Exception
	{
		SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		Date startDate = formatter.parse(inputStartDate);
		Date endDate = formatter.parse(inputEndDate);

		// Bind to the Calendar.
		CalendarFolder calendarFolder = CalendarFolder.bind(service, WellKnownFolderName.Calendar);
		FindItemsResults<Appointment> findResults = calendarFolder.findAppointments(new CalendarView(startDate, endDate));
//		for (Appointment appt : findResults.getItems())
//		{
//			System.out.println("SUBJECT=====" + appt.getSubject());
//			System.out.println("BODY========" + appt.getBody());
//		}
		
		
		// Bugfix: You must load or assign this property before you can read its value.
		// http://stackoverflow.com/questions/3304157/error-when-i-try-to-read-update-the-body-of-a-task-via-ews-managed-api-you-m
		List<Item> items = new ArrayList<Item>();
		List<Appointment> appointments = new ArrayList<Appointment>();
		for (Appointment appointment : findResults.getItems())
		{
			items.add(appointment);
		}
		service.loadPropertiesForItems(items, PropertySet.FirstClassProperties); //MOOOOOOST IMPORTANT: load messages' properties before
		for (Appointment appointment : findResults.getItems())
		{
			appointments.add(appointment);
		}
		
		return appointments;
	}
	
	private static void createAppointment(ExchangeService service, 
			String subject,
			String inputStartDate, 
			String inputEndDate,
			String body) throws Exception
	{
		Appointment appointment = new Appointment(service);
		appointment.setSubject(subject);
		appointment.setLocation("台灣G體電路製造股份有限公司中科廠");
		
		
		SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		Date startDate = formatter.parse(inputStartDate);
		Date endDate = formatter.parse(inputEndDate);
		appointment.setStart(startDate);//new Date(2010-1900,5-1,20,20,00));
		appointment.setEnd(endDate); //new Date(2010-1900,5-1,20,21,00));
		
		appointment = setAllDayEvent(appointment);

		// 重複性工作
//		formatter = new SimpleDateFormat("yyyy-MM-dd");
//		Date recurrenceEndDate = formatter.parse("2010-07-20");
//		appointment.setRecurrence(new Recurrence.DailyPattern(appointment.getStart(), 3));
//		appointment.getRecurrence().setStartDate(appointment.getStart());
//		appointment.getRecurrence().setEndDate(recurrenceEndDate);
		
		// 邀請對象
//		appointment.getRequiredAttendees().add("user1@contoso.com");
//		appointment.getRequiredAttendees().add("user2@contoso.com");
//		appointment.getOptionalAttendees().add("user3@contoso.com");
		
		// 提示
		appointment.setReminderMinutesBeforeStart(5);
//		appointment.setReminderDueBy(startDate);
		
		// 顯示為
		appointment.setLegacyFreeBusyStatus(LegacyFreeBusyStatus.Busy);
//		appointment.setLegacyFreeBusyStatus(LegacyFreeBusyStatus.Free);
//		appointment.setLegacyFreeBusyStatus(LegacyFreeBusyStatus.NoData);
//		appointment.setLegacyFreeBusyStatus(LegacyFreeBusyStatus.OOF);
//		appointment.setLegacyFreeBusyStatus(LegacyFreeBusyStatus.Tentative);
		
		// 附註
//		appointment.setBody(MessageBody.getMessageBodyFromText(body));
		appointment.setBody(new MessageBody(BodyType.Text, body)); // Bugfix: 預設會加入 html 標籤，而 body 會在 <body> 後換行，造成 EWS 搜尋功能失效。
		
		appointment.save();
	}
	
	private static void deleteAppointment(ExchangeService service, String uniqueId) throws Exception
	{
		Appointment appointment = Appointment.bind(service, new ItemId(uniqueId));
		appointment.delete(DeleteMode.HardDelete);
//		appointment.delete(DeleteMode.MoveToDeletedItems);
//		appointment.delete(DeleteMode.SoftDelete);
	}
	
	private static void updateAppointment(ExchangeService service, String uniqueId) throws Exception
	{
		Appointment appointment = Appointment.bind(service, new ItemId(uniqueId));
		SimpleDateFormat formatter = new  SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		Date startDate = formatter.parse("2016-05-19 13:00:00");
		Date endDate = formatter.parse("2016-05-19 14:00:00");

		appointment.setBody(MessageBody.getMessageBodyFromText("Appointement UPDATE done"));

		appointment.setStart(startDate);
		appointment.setEnd(endDate);
		appointment.setSubject("Appointement UPDATE");
//		appointment.getRequiredAttendees().add("someone@contoso.com");
		
		appointment.update(ConflictResolutionMode.AutoResolve);
	}

	/**
	 * 
	 *
	 * @param service
	 * @param searchString
	 * @throws Exception
	 * @see <a href="https://msdn.microsoft.com/zh-tw/library/office/dn579422(v=exchg.150).aspx">操作方法： 使用搜尋篩選與 Exchange 中的 EWS</a>
	 */
	private static void searchAppointments1(ExchangeService service, String searchString) throws Exception
	{
		System.out.println("searchString=" + searchString);
		ItemView view = new ItemView(10);
//		view.getOrderBy().add(ItemSchema.DateTimeReceived, SortDirection.Ascending);
		view.getOrderBy().add(AppointmentSchema.Start, SortDirection.Ascending);
//		view.setPropertySet(new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.DateTimeReceived));

		FindItemsResults<Item> findResults =
		        service.findItems(WellKnownFolderName.Calendar,
		            new SearchFilter.SearchFilterCollection(
		                LogicalOperator.Or, 
		                new SearchFilter.ContainsSubstring(ItemSchema.Subject, searchString),
		                new SearchFilter.ContainsSubstring(ItemSchema.Body, searchString)
		                ), view);

		// MOOOOOOST IMPORTANT: load items properties, before
		System.out.println("Total number of items found: " + findResults.getTotalCount());
		if (findResults.getTotalCount() == 0)
		{
			return; // Jump!
		}
		
		List<Item> items = new ArrayList<Item>();
		for (Item appointment : findResults.getItems())
		{
			items.add(appointment);
		}
		service.loadPropertiesForItems(items, PropertySet.FirstClassProperties); //MOOOOOOST IMPORTANT: load messages' properties before
		for (Item appointment : findResults.getItems())
		{
			print((Appointment) appointment);
		}
		
//		service.loadPropertiesForItems(findResults, PropertySet.FirstClassProperties);
//		for (Item item : findResults)
//		{
//			System.out.println(item.getSubject());
//			System.out.println(item.getBody());
//			System.out.println();
//			// Do something with the item.
//		}
	}
	
	private static Appointment setAllDayEvent(Appointment appointment) throws Exception
	{
		// JAVA 版本沒有 setAllDayEvent
		// https://github.com/OfficeDev/ews-java-api/issues/328
		SimpleDateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd");
		Date startDate = dateFormatter.parse(dateFormatter.format(appointment.getStart()));

		Date endDate = dateFormatter.parse(dateFormatter.format(appointment.getEnd()));
	    endDate = new Date(endDate.getTime() + TimeUnit.DAYS.toMillis(1));
	    
	    appointment.setStart(startDate);
	    appointment.setEnd(endDate);
		return appointment;
	}
	
	private static void print(Appointment appointment) throws ServiceLocalException
	{
		System.out.println(appointment.getId());
		System.out.println(appointment.getDateTimeCreated());
		System.out.println(appointment.getSubject());
		System.out.println(appointment.getBody());
		System.out.println(appointment.getStart());
		System.out.println(appointment.getEnd());
		System.out.println(appointment.getLocation());
		System.out.println();
	}
}

class RedirectionUrlCallback implements IAutodiscoverRedirectionUrl
{
	public boolean autodiscoverRedirectionUrlValidationCallback(String redirectionUrl)
	{
		return redirectionUrl.toLowerCase().startsWith("https://");
	}
}
