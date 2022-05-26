
public class GiftEntry {

	private String name = null;
	private String date = null;
	private double amount = 0;
	private double uniqueId = 0;
	
	public GiftEntry()
	{
		
	}
	
	public GiftEntry(String nameIn, String dateIn, double amountIn, double idIn)
	{
		name = nameIn;
		date = dateIn;
		amount = amountIn;
		uniqueId = idIn;
	}
	
	public String getName()
	{
		return name;
	}
	
	public void setName(String nameIn)
	{
		name = nameIn;
	}
	
	public String getDate()
	{
		return date;
	}
	
	public void setDate(String dateIn)
	{
		date = dateIn;
	}
	
	public double getAmount()
	{
		return amount;
	}
	
	public void setAmount(double amountIn)
	{
		amount = amountIn;
	}
	
	public double getUniqueId()
	{
		return uniqueId;
	}
	
	public void setUniqueId(int idIn)
	{
		uniqueId = idIn;
	}
	
	public String toString()
	{
		return "Name: "
				+ name
				+ ", Date: "
				+ date
				+ ", Amount: "
				+ amount;
	}
}

