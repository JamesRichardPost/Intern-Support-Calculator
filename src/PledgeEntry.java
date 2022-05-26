import java.time.*;
import java.util.Date;

public class PledgeEntry {

	private double amount = 0;
	private Date startDate = null;
	private int recurrances = 0;
	private int frequency = 0;
	private static int FISCAL_END_MONTH = 17;
	private String name = null;
	
	public PledgeEntry()
	{
		
	}
	
	public PledgeEntry(double amountIn, Date startDateIn, String frequencyIn, String nameIn)
	{
		amount = amountIn;
		startDate = startDateIn;
		name = nameIn;
		frequency = calFrequency(frequencyIn);
		recurrances = getRecurrances(frequency, startDate);
	}
	
	public double getAmount()
	{
		return amount;
	}
	
	public int getRecurrances()
	{
		return recurrances;
	}
	
	
	// figures the numeric representation of frequency, or
	// removes all non-digit characters from a string and returns them as an int
	public int calFrequency(String stringIn)
	{
		if (stringIn.contains("one time"))
			return 0;
		else if (stringIn.equalsIgnoreCase("yearly"))
			return 1;
		else if (stringIn.equalsIgnoreCase("monthly"))
			return 12;
		else if (stringIn.equalsIgnoreCase("quarterly"))
			return 4;
		else if (stringIn.contains("twice a year"))
			return 2;
		
		else
		{	
			String number = stringIn.replaceAll("[^\\d.]", "");
			return Integer.parseInt(number);
		}
	}
	
	public int getRecurrances(int frequencyIn, Date dateIn)
	{
		Date now = new Date();
		//Date now = new Date("7/30/2020");
		
		// one times and yearly need to be handled differently 
		if (frequencyIn == 0)
		{
			if (dateIn.before(now))
			{
				return 0;
			}
			
			else
			{
				return 1;
			}
		}
		
		// yearly/one time case
		if (frequencyIn == 1)
		{
			// need to make sure annual gifts from before the fiscal year's start
			// are still counted later
			
			// "year" count begins in 1900
			Date yearStart = new Date("6/1/" + Integer.toString((dateIn.getYear() + 1900)));
			
			if (dateIn.before(yearStart))
			{
				return 1;
			}
			
			else if (dateIn.before(now))
			{
				return 0;
			}
			
			else
			{
				return 1;
			}
		}
		
		// monthly
		if (frequencyIn == 12)
		{
			int month = now.getMonth()+1;
			int testMonth = now.getMonth();
			int testMonth2 = dateIn.getMonth();
			int testDay = now.getDate();
			int testDay2 = dateIn.getDate();
			boolean isIt = now.before(dateIn);
			
			//if (now.getMonth() == dateIn.getMonth() && now.before(dateIn))
			if (now.getDate() < dateIn.getDate())	
				return FISCAL_END_MONTH - (month - 1);
			else
				return FISCAL_END_MONTH - month;
		}		
		
		// quarterly
		if (frequencyIn == 4)
		{
			Date time = dateIn;
			if (now.after(dateIn))
				time = now;
			
			//since we're running these things in the summer, the end date
			// is always next June
			int endYear = LocalDate.now().getYear() + 1;
			int startYear = LocalDate.now().getYear();
			int endMonth = 6;
			int startMonth = time.getMonth()+1;
			
			return Math.floorDiv(((endYear*12+endMonth)-(startYear*12+startMonth))/3, 1);
		}
		
		//n-thly
		else
		{
			// make sure we don't double count the past
			Date time = dateIn;
			if (now.after(dateIn))
				time = now;
			
			//since we're running these things in the summer, the end date
			// is always next June
			int endYear = LocalDate.now().getYear() + 1;
			int startYear = LocalDate.now().getYear();
			int endMonth = 6;
			int startMonth = time.getMonth()+1;
			
			return Math.floorDiv(((endYear*12+endMonth)-(startYear*12+startMonth))/frequencyIn-1, 1);
		}
	}
	
	public String toString()
	{
		return  "Name: "
				+ name
				+ ", Amount: "
				+ amount
				+ ",\t Recurrances: "
				+ recurrances
				+ ",\t Frequency: "
				+ frequency
				+ ",\t Start Date: "
				+ startDate;
	}
	
}
