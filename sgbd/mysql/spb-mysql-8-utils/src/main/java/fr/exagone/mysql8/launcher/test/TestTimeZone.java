package fr.exagone.mysql8.launcher.test;

import java.util.Calendar;
import java.util.TimeZone;

public class TestTimeZone {

	public static void test1() {
		//
		// TEST CALENDAR
		//
		TimeZone timeZoneLosAngeles = TimeZone.getTimeZone("America/Los_Angeles");
		TimeZone timeZoneParis = TimeZone.getTimeZone("Europe/Paris");

		// CALENDAR PARIS
		Calendar cal = Calendar.getInstance();
		// 1- test avec meme date
		// cal.set(2013, 06, 12);
		// System.out.println("Paris : " + cal.getTimeInMillis());
		// 2- test avec meme timestamp
		cal.setTimeInMillis(1264881305341L);
		System.out.println("Paris - hour : " + cal.get(Calendar.HOUR));
		TimeZone timeZone = cal.getTimeZone();
		System.out.println("timeZone = " + timeZone);
		
		// CALENDAR LOS ANGELES
		Calendar calLA = cal.getInstance();
		calLA.setTimeZone(timeZoneLosAngeles);
		// 1- test avec meme date
		// calLA.set(2013, 06, 12);
		// System.out.println("LA    : " + calLA.getTimeInMillis());
		// 2- test avec meme time-stamp
		calLA.setTimeInMillis(1264881305341L);
		System.out.println("Los Angeles - hour : " + calLA.get(Calendar.HOUR));
		
	}
	
}
