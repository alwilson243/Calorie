// Include Apache POI elements
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;

// Include GUI elements
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;

/**
 * This class tracks daily caloric intake by receiving input from the user in the form
 * of food elements consumed.
 * @author Andrew Wilson
 */
public class CalorieExcel extends JFrame
{
	// Keeps track of date alongside calorie total
    private static Calendar calendar = Calendar.getInstance();
    private double date = dateStamp();
    
    // Used when overwriting today's calorie total
    private static boolean overwrite = false;
    
    // Lists and arrays of food and associated calories
    private static ArrayList<String> foodList = new ArrayList<String>();
    private static ArrayList<String> servingList = new ArrayList<String>();
    private static ArrayList<Double> calorieList = new ArrayList<Double>();
    private static String[] merged;
    
    // Represents running calorie count for current day
    private double count = 0;
    private String countStr = Double.toString(count);
    private double calorie = calorieList.get(0);
    
    /**
     * The constructor initializes GUI elements
     */
    public CalorieExcel()
    {
    	// Initialize JFrame
        JFrame frame = new JFrame();
        frame.setTitle("Calorie Calculator");
        frame.setSize(400, 300);
        frame.setLocationRelativeTo(null);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        
        // Initialize JPanel
        JPanel panel = new JPanel();
        panel.setLayout(new FlowLayout());
        frame.add(panel);
        
        // Initialize JLabel
        JLabel label = new JLabel("Calories:");
        panel.add(label);
        
        // Initialize JTextField containing string calorie value
        final JTextField text = new JTextField(countStr);
        text.setColumns(4);
        panel.add(text);
        
        // Initialize JComboBox containing food list
        final JComboBox combo = new JComboBox(merged);
        combo.addActionListener(new ActionListener()
        {
            public void actionPerformed(ActionEvent arg0)
            {
                calorie = calorieList.get(combo.getSelectedIndex());    
            }
        });            
        panel.add(combo);
        
        // Initialize JButton that adds calories to running total
        JButton add = new JButton("Add");
        add.addActionListener(new ActionListener()
        {
          public void actionPerformed(ActionEvent e)
          {
            count += calorie;
            countStr = Double.toString(count);
            text.setText(countStr);
          }
        });
        panel.add(add);
        
        // Initialize JButton that saves running total of calories to excel file
        JButton save = new JButton("Save");
        save.addActionListener(new ActionListener()
        {
            public void actionPerformed(ActionEvent e)
            {
                try
                {
                    // If overwriting current day's calories
                    if(overwrite)
                        overwriteSave();
                    // Not the same day as previous writing
                    else
                        save();
                } 
                catch (IOException e1)
                {
                    e1.printStackTrace();
                }
            }
        });
        panel.add(save);
        
        // Initialize JButton that clears running total
        JButton clear = new JButton("Clear");
        clear.addActionListener(new ActionListener()
        {
            public void actionPerformed(ActionEvent e)
            {
                count = 0;
                countStr = "0";
                text.setText("0.0");
            }
        });
        panel.add(clear);

        frame.setVisible(true); 
    }
    
    /**
     * This method converts cell values to strings
     * @param cell
     * @return String representation of the cell
     */
    public static String cellToString (XSSFCell cell)
    {
        int type;
        Object result = null;
        type = cell.getCellType();
        
        // Different cell types must be accounted for
        switch(type)
        {
            case 0: result = cell.getNumericCellValue();
                break;
            case 1: result = cell.getStringCellValue();
                break;
        }
        return result.toString();
    }
    
    /**
     * This method reads in the input book from the excel file,
     * populating the respective lists with food, serving size,
     * and calorie information
     * @param stringCell String representation of the cell
     * @param column from the excel table dictates destination list
     */
    public static void populate(String stringCell, int column)
    {
        switch(column)
        {
            case 0: foodList.add(stringCell);
                break;
            case 1: servingList.add(stringCell);
                break;
            case 2: calorieList.add(Double.parseDouble(stringCell));
                break;
        }
    }
    
    /**
     * This method merges the food and serving size strings for display
     * in the drop down box
     * @return merged array of Strings
     */
    public static String[] mergeStrings()
    {
        int length = foodList.size();
        String[] merged = new String[length];
        
        for(int i = 0; i < length; i++)
        {
            merged[i] = foodList.get(i) + " (" + servingList.get(i) + ")";
        }
        
        return merged;
    }
    
    /**
     * This method creates a unique datestamp for the current day
     * @return date in unique format
     */
    public static double dateStamp()
    {
        int day = calendar.get(Calendar.DAY_OF_MONTH);
        int month = calendar.get(Calendar.MONTH) + 1;
        int year = calendar.get(Calendar.YEAR);
        
        double date = year * 10000 + month * 100 + day;
        
        return date;
    }
    
    /**
     * This method is called when the user wishes to save their running
     * calorie total to the excel file and they do not wish to overwrite
     * any existing information because it is a new day
     * @throws IOException
     */
    public void save() throws IOException
    {
        FileInputStream saveFile = new FileInputStream(new File("calorie.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(saveFile);
        XSSFSheet outputSheet = workbook.getSheetAt(1);
        
        int outputColumnLength = outputSheet.getLastRowNum();

        XSSFRow outputRow = outputSheet.createRow(outputColumnLength + 1);
        
        XSSFCell cell0 = outputRow.createCell(0);
        cell0.setCellValue(date);
        XSSFCell cell1 = outputRow.createCell(1);
        cell1.setCellValue(count);
        
        saveFile.close();
        FileOutputStream out = new FileOutputStream(new File("calorie.xlsx"));
        workbook.write(out);
        out.close();
    }
    
    /**
     * This method is called when the user wishes to save their running
     * calorie total to the excel file and they do  wish to overwrite
     * existing information because it is the same day
     * @throws IOException
     */
    public void overwriteSave() throws IOException
    {
        FileInputStream saveFile = new FileInputStream(new File("calorie.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(saveFile);
        XSSFSheet outputSheet = workbook.getSheetAt(1);
        
        int outputColumnLength = outputSheet.getLastRowNum();
        
        XSSFRow outputrow = outputSheet.getRow(outputColumnLength);
        XSSFCell datecell = outputrow.getCell(0);
        XSSFCell cell1 = outputrow.createCell(1);
        
        cell1.setCellValue(count);

        saveFile.close();
        FileOutputStream out = new FileOutputStream(new File("calorie.xlsx"));
        workbook.write(out);
        out.close();
    }

}
