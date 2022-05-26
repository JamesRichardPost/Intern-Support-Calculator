import java.awt.BorderLayout;
import java.awt.FileDialog;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.File;  
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintStream;
import java.sql.Date;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.time.LocalDate;
import java.util.Iterator;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
import org.apache.poi.ss.usermodel.Sheet;  
import org.apache.poi.ss.usermodel.*;  
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextPane;
import javax.swing.SwingUtilities;
import javax.swing.text.BadLocationException;
import javax.swing.text.Document;
import javax.swing.text.NumberFormatter;

import java.util.ArrayList;

public class XLSXReader 
{  
	private static int idColumn = 7;
	private static int dateColumn = 4;
	private static int amountColumn = 3;
	private static int nameColumn = 2;
	
	
	
	@SuppressWarnings({ "deprecation", "static-access" })
	public static void main(String args[]) throws IOException  
	{  

		// GUI area
		JFrame frame = new JFrame("Intern Support Calculator");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800,800);
        JButton chooseFile = new JButton("Select the file you want to calculate from.");
        frame.getContentPane().add(BorderLayout.NORTH, chooseFile);
        frame.setVisible(true);
        JTextPane ta = new JTextPane();

        frame.getContentPane().add(BorderLayout.SOUTH, ta);
        ta.setEditable(false);
        Font font = new Font("Default", Font.BOLD, 12);
        ta.setFont(font);
        
        
		
		chooseFile.addActionListener(new ActionListener()
		{
			public void actionPerformed (ActionEvent e)
			{
				JFileChooser jfc = new JFileChooser();
			    
				jfc.showDialog(null,"Select the File");
			    jfc.setVisible(true);
			    
			    File filename = jfc.getSelectedFile();
			    ta.setText("File name: "+filename.getName());
				
			    String path = filename.getPath();
				//obtaining input bytes from a file  
				File file = new File(path);  

				if (getFileExtension(file).equalsIgnoreCase("xlsx"))
					excelCalculator(file, ta);
				
				else if (getFileExtension(file).equalsIgnoreCase("csv"))
					csvCalculator(path, ta);
			}
		});
		
		
        JScrollPane jsp = new JScrollPane(ta);
		frame.getContentPane().add(jsp);
		frame.setVisible(true);
		
	}
	
		
	@SuppressWarnings({ "deprecation", "static-access" })
	public static String readCellData(int rowIn, int columnIn, XSSFWorkbook workbookIn, boolean isDate)  
	{  	
		String value = null;
		
		Sheet sheet = workbookIn.getSheetAt(0);   //getting the XSSFSheet object at given index  
		Row row = sheet.getRow(rowIn); //returns the logical row  
		Cell cell = row.getCell(columnIn); //getting the cell representing the given column  
		
		if (cell.getCellType() == cell.CELL_TYPE_STRING)
			value = cell.getStringCellValue();    //getting cell value
		else if (cell.getCellType() == cell.CELL_TYPE_NUMERIC && !isDate)
			value = String.valueOf(cell.getNumericCellValue());
		else if (cell.getCellType() == cell.CELL_TYPE_NUMERIC && isDate)
			value = cell.getDateCellValue().toString();

		return value;               //returns the cell value  
	}  
	
	private static String getFileExtension (File file) 
	{
        String fileName = file.getName();
        if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
        	return fileName.substring(fileName.lastIndexOf(".")+1);
        
        else return "";
    }
	
	public static void excelCalculator(File file, JTextPane ta)
	{
		double cashOnHand = 0;
		double pledgeTotal = 0;
		double budget = 0;
		
		try
		{
		FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
		//creating Workbook instance that refers to .xlsx file  
		XSSFWorkbook wb = new XSSFWorkbook(fis);   

		XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
		Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
		
		int rowNumber = 0;
		int giftsHeaderRow = 0;
		int giftsHeaderColumn = 0;
		int pledgesHeaderRow = 0;
		int pledgesHeaderColumn = 0;
		int tasksAndInteractionsRow = 0;
		int groupsOffset = 0;
		boolean isCheckMessage = false;
		String[] checkMessage = new String[20];
		int checkMessageIndex = 0;
		
		// The first thing we need to do is find where each section begins so we can extract
		//and use the data from each section
		
		// one iteration gives us the header for gifts and pledges and also the toal budget
		while (itr.hasNext())                 
		{  
			Row row = itr.next();
			
			
			//Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
			
			Cell cell = row.getCell(0); 
			
			if (cell != null)
			{
				int type = cell.getCellType();	
			
				
			if (type == cell.CELL_TYPE_STRING)
			{
				if	(cell.getStringCellValue().contains("GIFTS"))
				{
					giftsHeaderRow = rowNumber + 4;
				}
				
				else if (cell.getStringCellValue().contains("PLEDGES"))
				{
					pledgesHeaderRow = rowNumber + 6;
				}
				
				else if (cell.getStringCellValue().contains("TASKS & INTERACTIONS"))
				{
					tasksAndInteractionsRow = rowNumber + 8;
				}
				
				else if (cell.getStringCellValue().contains("GROUPS"))
				{
					groupsOffset = 3;
				}
				
				else if(cell.getStringCellValue().equalsIgnoreCase("yearly_pledge_goal"))
				{
					budget = row.getCell(1).getNumericCellValue();
				}
			}	
			}
			
			rowNumber++; 
			
		}
		
		// Now that we have the limits of that section, we're going to grab the data we need
		// to make an array of GiftEntries
		List<GiftEntry> gifts = new ArrayList<>();
		ta.setText(ta.getText() + "\n\n GIFTS: \n\n");
		
		for (int i = giftsHeaderRow+2; i < pledgesHeaderRow - 2; i++)
		{
			String name = readCellData(i, nameColumn, wb, false);
			String date = readCellData(i, dateColumn, wb, false);
			double amount = Double.parseDouble(readCellData(i, amountColumn, wb, false));
			double id = Double.parseDouble((readCellData(i, idColumn, wb, false)));
			
			GiftEntry gift = new GiftEntry(name, date, amount, id);		
			
			String ptext = ta.getText();
			ta.setText(ptext + "\n" + gift.toString());
			
			gifts.add(gift);
		}
		
		// Now we have our gift objects, but we haven't gotten rid of those read-only entries
		// so let's see what we can do about that.
		
		for (int i = 0; i < gifts.size(); i++)
		{
			GiftEntry gift = gifts.get(i);
			if (gift.getName().contains("Foundation") || gift.getName().contains("Fund") || gift.getName().contains("Charitable"))
				{
				GiftEntry nextGift = gifts.get(i + 1);
				if (gift.getAmount() == nextGift.getAmount() && gift.getUniqueId() == nextGift.getUniqueId()+1)
					gifts.get(i).setAmount(0);
				else if (gift.getAmount() == nextGift.getAmount() && gift.getDate().equalsIgnoreCase(nextGift.getDate()))
				{
					gifts.get(i).setAmount(0);
					
					String ptext = ta.getText();
					ta.setText(ptext + "\n" +"We believe this person recieved a gift from " + gift.getName()
						+ " that should be marked as view only. \nHowever, because the unique id's were"
						+ " not consecutive (indicating that they may have come in at different times)"
						+ " we encourage you to double check. \nThe gift in question has not been included"
						+ " in the final total. \n");
				}
			}
		}
		
		// now let's sum up our cash on hand
		
		for (int i = 0; i < gifts.size(); i++)
		{
			cashOnHand += gifts.get(i).getAmount();
		}
		
		
		
		// now we build our list of pledge entries
		ta.setText(ta.getText() + "\n\n PLEDGES: \n\n");
		
		ArrayList<PledgeEntry> pledges = new ArrayList<PledgeEntry>();
		
		for (int i = pledgesHeaderRow+2; i < tasksAndInteractionsRow-2; i++)
		{
			double amount = Double.parseDouble(readCellData(i, 1, wb, false));
			Date startDate = new Date(Date.parse(((readCellData(i, 5, wb, true)))));
			String frequency = readCellData(i, 2, wb, false);
			String name = readCellData(i, 0, wb, false);
			
			PledgeEntry pledge = new PledgeEntry(amount, startDate, frequency, name);
			
			String ptext = ta.getText();
			ta.setText(ptext + "\n" + pledge.toString());
			
			pledges.add(pledge);
		}
		
		// now we total the amount pledged over the year
		for (int i = 0; i < pledges.size(); i++)
		{
			PledgeEntry pledge = pledges.get(i);
			pledgeTotal += (pledge.getAmount() * pledge.getRecurrances());
		}
		
		DecimalFormat df = new DecimalFormat("###, ###, ###.00");
		NumberFormat pf = NumberFormat.getPercentInstance();
		ta.setText(ta.getText() + "\n\n RESULTS: \n\n");
		
		String text = ta.getText();
		ta.setText(text
		+ "\nCash Total: $" + df.format(cashOnHand)
		+ "\nPledge Total: $" + df.format(pledgeTotal)
		+ "\nGifts + Pledges: $" + df.format((cashOnHand + pledgeTotal))
		+ "\nBudget: $" + df.format(budget)
		+ "\n85 Percent: $" + df.format((budget * .85))
		+ "\nCurrent Percentage: " + (pf.format((cashOnHand + pledgeTotal)/budget)));
		
	}
		catch (Exception e)
		{
			ta.setText("Something went wrong opening the file. More than likely"
					+ " the file you are trying to open is not saved as an Excel Workbook."
					+ " \nTo fix this, open the file you are trying to load in Excel, select"
					+ " 'Save As' and choose 'Excel Workbook'.");
		}
	}
	
	public static void csvCalculator(String path, JTextPane ta)
	{
		try
		{
			String row = "";
			ArrayList<ArrayList> sheet = new ArrayList<ArrayList>();
			BufferedReader csvReader = new BufferedReader(new FileReader(path));
			
			int giftsHeaderRow = 0;
			int pledgesHeaderRow = 0;
			int tasksAndInteractionsRow = 0;
			double cashOnHand = 0;
			double pledgeTotal = 0;
			double budget = 0;
			String[] checkMessage = new String[20];
			int checkIndex = 0;
			
			// recreate the sheet in an arraylist
			while ((row = csvReader.readLine()) != null)
			{
				String[] data = row.split(",");
				ArrayList<String> thisRow = new ArrayList<>();
				for (int i = 0; i < data.length; i++)
				{
					thisRow.add(data[i]);
				}
				sheet.add(thisRow);
			}
			
			// find the indexes of our three headers
			for (int i = 0; i < sheet.size(); i++)
			{
				if (sheet.get(i).get(0).toString().contains("GIFTS"))
				{
					giftsHeaderRow = i;
				}
				
				if (sheet.get(i).get(0).toString().contains("PLEDGES"))
				{	
					pledgesHeaderRow = i;
				}
				if (sheet.get(i).get(0).toString().contains("TASKS & INTERACTIONS"))
				{
					tasksAndInteractionsRow = i;
				}
				
				if (sheet.get(i).get(0).toString().equalsIgnoreCase("yearly_pledge_goal"))
				{
					budget = Double.parseDouble(sheet.get(i).get(1).toString());
				}
				
			}
			
			// Now that we have the limits of that section, we're going to grab the data we need
			// to make an array of GiftEntries
			List<GiftEntry> gifts = new ArrayList<>();
			ta.setText(ta.getText() + "\n\n GIFTS: \n\n");
			
			for (int i = giftsHeaderRow+2; i < pledgesHeaderRow - 2; i++)
			{
				String name = sheet.get(i).get(nameColumn).toString();
				
				// we need an offset for organizations that include a comma in their name
				// like bad people who are bad
				int offset = 0;
				int index = 3;
				while (sheet.get(i).get(index).toString().contains("\""))
				{
					offset++;
					index++;
				}
				
				double amount = Double.parseDouble(sheet.get(i).get(amountColumn + offset).toString());
				String date = sheet.get(i).get(dateColumn + offset).toString();
				double id = Double.parseDouble(sheet.get(i).get(idColumn + offset).toString());
				
				GiftEntry gift = new GiftEntry(name, date, amount, id);		
				
				String ptext = ta.getText();
				ta.setText(ptext + "\n" + gift.toString());
				
				gifts.add(gift);
			}
			
			// Now we have our gift objects, but we haven't gotten rid of those read-only entries
			// so let's see what we can do about that.
			
			for (int i = 0; i < gifts.size(); i++)
			{
				GiftEntry gift = gifts.get(i);
				if (gift.getName().contains("Foundation") || gift.getName().contains("Fund") || gift.getName().contains("Charitable"))
				{
					// make sure we don't run over the end of the array
					if (i != gifts.size()-1)
					{
					GiftEntry nextGift = gifts.get(i + 1);
					if (gift.getAmount() == nextGift.getAmount() && gift.getUniqueId() == nextGift.getUniqueId()+1)
						gifts.get(i).setAmount(0);
					else if (gift.getAmount() == nextGift.getAmount() && gift.getDate().equalsIgnoreCase(nextGift.getDate()))
					{
						gifts.get(i).setAmount(0);
						
						checkMessage[checkIndex] = "\n\n" + "We believe this person recieved a gift from " + gift.getName()
							+ " that should be marked as view only. \nHowever, because the unique id's were"
							+ " not consecutive (indicating that they may have come in at different times)"
							+ " we encourage you to double check. \nThe gift in question has not been included"
							+ " in the final total. \n";
						checkIndex++;
					}
					}
					
					// in a .csv sometimes the foundation comes second, for whatever reason (although the uid
					// is still +1
					
					// make sure we don't run off the beginning of the array
					if (i != 0)
					{
					GiftEntry previousGift = gifts.get(i - 1);
					if (gift.getAmount() == previousGift.getAmount() && gift.getUniqueId() == previousGift.getUniqueId()+1)
						gifts.get(i).setAmount(0);
					else if (gift.getAmount() == previousGift.getAmount() && gift.getDate().equalsIgnoreCase(previousGift.getDate()))
					{
						gifts.get(i).setAmount(0);
						
						checkMessage[checkIndex] = "\n\n" + "We believe this person recieved a gift from " + gift.getName()
							+ " that should be marked as view only. \nHowever, because the unique id's were"
							+ " not consecutive (indicating that they may have come in at different times)"
							+ " we encourage you to double check. \nThe gift in question has not been included"
							+ " in the final total. \n";
						checkIndex++;
					}
					}
				}
			}
			
			// now let's sum up our cash on hand
			
			for (int i = 0; i < gifts.size(); i++)
			{
				cashOnHand += gifts.get(i).getAmount();
			}
			
			// now we build our list of pledge entries
			ArrayList<PledgeEntry> pledges = new ArrayList<PledgeEntry>();
			ta.setText(ta.getText() + "\n\n PLEDGES: \n\n");
			
			for (int i = pledgesHeaderRow+2; i < tasksAndInteractionsRow-2; i++)
			{
				// same thing as in gifts, we have to fix when bad people 
				// put commas in places where they don't belong
				int offset = 0;
				int index = 1;
				while (sheet.get(i).get(index).toString().contains("\""))
				{
					offset++;
					index++;
				}
				
				double amount = Double.parseDouble(sheet.get(i).get(1 + offset).toString());
				String date = sheet.get(i).get(5 + offset).toString();
				String[] dates = date.split("-");
				
				int year = Integer.parseInt(dates[0]) - 1900;
				int month = Integer.parseInt(dates[1]) - 1;
				int day = Integer.parseInt(dates[2]);
				
				Date startDate = new Date(year, month, day);
				String frequency = sheet.get(i).get(2 + offset).toString();
				String name = sheet.get(i).get(0).toString();
				
				PledgeEntry pledge = new PledgeEntry(amount, startDate, frequency, name);
				
				String ptext = ta.getText();
				ta.setText(ptext + "\n" + pledge.toString());
				
				pledges.add(pledge);
			}
			
			// now we total the amount pledged over the year
			for (int i = 0; i < pledges.size(); i++)
			{
				PledgeEntry pledge = pledges.get(i);
				pledgeTotal += (pledge.getAmount() * pledge.getRecurrances());
			}
			
			DecimalFormat df = new DecimalFormat("###, ###, ###.00");
			NumberFormat pf = NumberFormat.getPercentInstance();
			ta.setText(ta.getText() + "\n\n RESULTS: \n\n");
			
			String text = ta.getText();	
			ta.setText(text +
			"\nCash Total: $" + df.format(cashOnHand)
			+ "\nPledge Total: $" + df.format(pledgeTotal)
			+ "\nGifts + Pledges: $" + df.format((cashOnHand + pledgeTotal))
			+ "\nBudget: $" + df.format(budget)
			+ "\n85 Percent: $" + df.format((budget * .85))
			+ "\nCurrent Percentage: " + (pf.format((cashOnHand + pledgeTotal)/budget)));
			
			
			for (int i = 0; i < checkIndex; i++)
			{
				text = ta.getText();
				ta.setText(text + checkMessage[i]);
			}
			

		}
		
		catch (Exception e)
		{
			System.out.println(e.getMessage());
		}
	}
	
}

	
	
 
