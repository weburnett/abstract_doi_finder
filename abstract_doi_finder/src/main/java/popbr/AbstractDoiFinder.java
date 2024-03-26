package popbr;

import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import java.util.Date;
import java.util.ArrayList;
import java.util.Scanner;
import java.util.Objects;

import java.net.UnknownHostException;
import java.net.MalformedURLException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;

import org.apache.commons.io.FileUtils;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class AbstractDoiFinder {
   public static void main(String[] args) throws Exception {
      /*
       * Welcome message and looking for spreadsheet to process.
       */
      System.out.println("Welcome to Abstract/DOI Finder.");
      try{
         String[] providedFile = new String[1];
         String numberRange = "*";
         if (args.length > 2)
            throw new IOException("There are too many arguments. \nPlease try removing spaces in your file name or in your list of numbers.");
         if (args.length == 2)
         {
            /*
               Checks which argument is the excel file
               If there are more than 1 argument, but neither is the excel file, we stop the program.
               Right now, the way the arguments are input are very important and could lead to user error.
             */
            if (args[0].endsWith(".xlsx"))
            {
               providedFile[0] = args[0]; // if the first argument ends with .xlsx, it is the excel file.
               numberRange = args[1];
            }
            else
            {
               if (!args[1].endsWith(".xlsx"))
                  throw new IOException("You provided 2 arguments, but one of them was not an excel file. If you're just trying to include a sheet range, please check to make sure there are no numbers between your comma-separated values.");
               providedFile[0] = args[1];
               numberRange = args[0];
            }
         }
         if (args.length == 1)
         {
            // check if the file ends in .xlsx, if not, it is the number range
            if (args[0].endsWith(".xlsx"))
               providedFile[0] = args[0];
            else
               numberRange = args[0];
         }
         // need to implement logic to find which one would be the file. First argument could be excel file.
         /*
             Four possibilities:
             1. File with sheet range
             2. File without sheet range
             3. Range without file 
             4. No arguments provided (should be relatively easy)
             Right now, we're going to assume if one of the arguments does not in .xlsx, then it is the sheet range.
          */
         File inputPath = FindInputFile(providedFile);
         if(inputPath.length() == 0 ){ // If the file returned is empty.
            System.out.println("This file seems to be empty.\nPlease check " + inputPath + ".");
         }
         else{ // If the file returned is not empty.

            int startingSheet;

            File outputPath = CreateOutput(inputPath);
            XSSFWorkbook wb = new XSSFWorkbook(outputPath);
            int number_of_sheets = wb.getNumberOfSheets();

            String[] sheetNumbers = numberRange.split("[,]");
            for (int i = 0; i < sheetNumbers.length; i++)
            {
               sheetNumbers[i] = sheetNumbers[i].trim();  // gets rid of the white space if the user puts spaces in between their comma separated values.
            }
            ArrayList<Integer> sheetNumbers2 = new ArrayList<Integer>(); // will come up with a better name for this soon

            /*
              The way I am trying to do it, it will need a starting index either way because if the user enters *, I'm going to have to make it 2
              I will look for more ways when I am able to, but for now, I want to figure this out.
             */

            for (int i = 0; i < sheetNumbers.length; i++)
            {
               if (sheetNumbers[i].equals("*"))
               {
                  for (int j = 0; j < number_of_sheets; j++){
                     sheetNumbers2.add(j);
                  }
                  break;
               }
               if (sheetNumbers[i].contains("-"))
               {
                  String[] rangeNumbers = sheetNumbers[i].split("-"); //should only have two numbers, the first in the range and the last "number"
                  if (rangeNumbers[0].equals("*"))
                     throw new Exception("You cannot include * as the first argument in a range. Please try again with different sheet selections.");
                  
                  if (rangeNumbers[1].equals("*"))
                  {
                     if (Integer.parseInt(rangeNumbers[0]) > number_of_sheets)
                        throw new Exception("The numbers provided in the range exceed the number of sheets in the specified excel file.");
                     for (int j = Integer.parseInt(rangeNumbers[0]); j <= number_of_sheets; j++)
                        sheetNumbers2.add(j);
                     continue;
                  }
                  else
                  {
                     for (int j = Integer.parseInt(rangeNumbers[0]); j <= Integer.parseInt(rangeNumbers[1]); j++)
                     {
                        if (j > number_of_sheets)
                           break;
                        sheetNumbers2.add(j);
                     }
                     continue;
                  }
               }
               if (Integer.parseInt(sheetNumbers[i]) > number_of_sheets)
               {
                  System.out.print(sheetNumbers[i] + " exceeds the range of the number of sheets so the program will only run on the valid sheets specified before this number.");
                  break;
               }
               sheetNumbers2.add(Integer.parseInt(sheetNumbers[i]));
            }

            wb.close();

            startingSheet = sheetNumbers2.get(0);

            //int startingSheet = 2; // This is bad practise: I am assuming that the sheets starting with the 3 (included) needs to gets filled with doi / abstracts.
            // Cf. https://github.com/popbr/abstract_doi_finder/issues/9 on how to address that.
            // And on how to get a more clever way of bounding the number of sheets explored.
            
            // We now open the spreadsheet to count its number of sheets.
             
            //File outputPath = CreateOutput(inputPath);
           // XSSFWorkbook wb = new XSSFWorkbook(outputPath);
            //int number_of_sheets = wb.getNumberOfSheets();
            //wb.close();
            
            /* 
             * This is the main part of the program.
             * For each sheet, between startingSheet and number_of_sheets, it 
             *     - reads from the sheet the various values needed to perform the queryy (Read_From_Excel), 
             *     - performs the queries (RetrieveData), 
             *     - writes the data retrieved (Write_To_Excel) in the appropriate sheet.
             * This is the time-consuming part.
             */
            /*
               Need to fix the program to include the new logic for when specific sheets are treated. 
               It does not crash, but it would not run on the sheets provided right now. 
             */
            for (int sheetIndex=sheetNumbers2.get(0); sheetIndex < number_of_sheets; sheetIndex++){
               if (!sheetNumbers2.contains(sheetIndex))
                  continue;
               ArrayList<String> searchList = Read_From_Excel(sheetIndex, inputPath); // Returns a searchList that has the author's name and all of the titles for our search query
               ArrayList<ArrayList<String>> contentList = RetrieveData(searchList); // Takes a few minutes to accomplish due to having to search on the Internet
               Write_To_Excel(contentList, sheetIndex, outputPath); // Currently only does one sheet at a time and needs to be manually update
            }
            System.out.println("Thanks for coming! Your abstracts and DOIs should be in your Excel file now");
         }
      } catch(IOException e) { // Those are the exceptions returned by the CreateOutput method.
         System.out.println(e.getMessage());
         System.exit(0);
      }
   }
   
   /* 
    * This method tries to establish if a spreadsheet 
    * can be found. It will throw an exception if 
    * - The input/ folder does not exist,
    * - The file specified by the user does not exist, 
    * - The file specified by the user is not a spreadheet,
    * - Or, if the user did not specify an input file, 
    *   if the input/ folder does not contain only one single spreasheet,
    * Otherwise, the spreadsheet is returned as a file.
    */
   
   public static File FindInputFile(String[] args) throws Exception {
      // We first make sure that the input/ folder exists.
      String BasePath = System.getProperty("user.dir"); // Current folder.
      File inputFolder = new File(BasePath + File.separator + "input"); // Input folder.
      File inputPath; // This variable will hold the file (path) to process.
      
      // Create a FileFilter, cf. https://www.geeksforgeeks.org/file-listfiles-method-in-java-with-examples/
      // This will let us filter files that ends with xlsx
      FileFilter xlsxFilter = new FileFilter() { 
         public boolean accept(File f) 
         { 
            return f.getName().endsWith("xlsx"); 
         } 
      }; 
      
      if (!inputFolder.exists() || !inputFolder.isDirectory()) { // If input/ does not exists, or if it is not a folder.
         throw new IOException("Sorry, there is no input/ folder, there is nothing I can do.\nPlease, create an input/ folder and place your spreadsheet in it.");
      }
      else{ // The input/ folder exists
         if (args[0] == null || args.length == 0) { // Did the user gave an argument?
            System.out.println("You did not provide a file to process, I will look in the input/ folder for a spreadsheet.");
            File[] filesInInputFolder = inputFolder.listFiles(xlsxFilter); // We filter the list of files in input/, looking for files that ends with xlsx.
            if (filesInInputFolder.length > 1){
               String exceptionToReturn = "There seems to be multiple xlsx files in the input/ folder:\n";
               // List the names of the files 
               for (int i = 0; i < filesInInputFolder.length; i++) { 
                  exceptionToReturn += "\t- " + filesInInputFolder[i] + "\n"; 
               }
               exceptionToReturn += "Please, place only one xlsx file in the input/ folder.";
               throw new IOException(exceptionToReturn);
            }
            else if (filesInInputFolder.length == 0){
               throw new IOException("There seems to no xlsx files in the input/ folder. Please, provide one xlsx file.");
            }
            else{ // there is exactly one spreadsheet in the input/ folder
               System.out.println("I found exactly one spreadsheet in the input folder:\n\t" + filesInInputFolder[0] + "\nAnd will process that file.");
               inputPath = filesInInputFolder[0];
            }
         }
         else{ // The user provided an argument.
            inputPath = new File(BasePath + File.separator + "input" + File.separator + args[0]);
            if (!inputPath.exists()){
               throw new IOException("You requested that I process the file\n\t" + inputPath + "\nBut this file does not seem to exist.");
            }
            else if (!inputPath.getName().endsWith("xlsx")) { // If the argument does not match a file, or if 
               throw new IOException("You requested that I process the file\n\t" + inputPath + "\nBut this file does not seem to be a spreadsheet.");
            }
            else { // The file exists, and it ends with xlsx
               System.out.println("Ok, I will now process:\n\t" + inputPath + "\n");
            }
         }
      }
      return inputPath;
   }
   
   public static File CreateOutput(File inputPath) throws Exception{
      /*
       * First, we make sure that the name we want to use is available.
       */
      String BasePath = System.getProperty("user.dir"); // Current folder.
      // We first create the output/ folder
      new File(BasePath + File.separator + "output").mkdirs();
      String timeStamp = new java.text.SimpleDateFormat("yyyy.MM.dd.HH.mm.ss").format(new java.util.Date());
      File outputPath = new File(BasePath + File.separator + "output" + File.separator + timeStamp + ".xlsx");
      if (outputPath.exists()){
         throw new IOException("It seems that the file\n\t" + outputPath + "\nalready exists. Please rename it so that it doesn't get overridden.");
      }
      else{
         System.out.println("Output file will stored in\n\t" + outputPath + ".");
      }
   
      FileOutputStream outputFile = new FileOutputStream(outputPath); // FileOutputStream needed for write operation later on.
      FileUtils.copyFile(inputPath, outputPath); // We make a simple "file" copy of the input sheet in the target path.
            
      /*
       * Now, we open the outputPath spreadsheet and shift the columns to insert a DOI column in 10th row.
       */
      
      XSSFWorkbook wb = new XSSFWorkbook(outputPath);

      wb = ShiftColumns(wb);

      /*
       * Finally, we write the spreadsheet we obtained in the outputFile and return its path.
       */
      
      wb.write(outputFile);
      outputFile.close();
      wb.close();
      return outputPath;
   }

   public static XSSFWorkbook ShiftColumns(XSSFWorkbook wb) throws Exception {
      int startingSheet = 0; // This is bad practise: We're assuming that the sheets starting with the 0 (included) needs to be shifted
      int abstractColumn = 0, doiColumn = 0; // Initializing the values for the columns, it throws an error otherwise. (title should be part of the general format, although it will most likely be there anyway)
      boolean hasAbstractColumn = false, hasDOIColumn = false, hasTitle = false; // If the sheet has a column, we do not want to add another one of the same type

      /* 
        Loops through every sheet of the workbook to see if an abstract or doi column already exists.
        If it does, it does not make a new column for our specified attributes.
        If not, it makes a new column for the missing column.

        Additionally, we take the liberty of where the abstract and doi columns go if they do not exist.
        This can be reworked later.
      */
      for (int sn = startingSheet; sn < wb.getNumberOfSheets(); sn++)
      {
         hasTitle = false;
         hasAbstractColumn = false;
         hasDOIColumn = false; // resets the boolean values back to false before switching to the next sheet
         Sheet sheet = wb.getSheetAt(sn);
         int noOfColumns;
         if (sheet == null || 
            Objects.isNull(sheet.getRow(0))) // if the sheet is null, further lines will crash the program.
            continue;
         if (Objects.isNull(sheet.getRow(0).getPhysicalNumberOfCells()))
            continue;
         if (sheet.getRow(0).getPhysicalNumberOfCells() == 0)
            continue;   
         else
            noOfColumns = sheet.getRow(0).getPhysicalNumberOfCells(); // This is bad practise: We're am assuming that each row has the same number of columns as the first one.

         for (int i = 0; i < noOfColumns; i++)
         {
            Cell cell = sheet.getRow(0).getCell(i);
            if (cell == null)
               continue;
            if (cell.getCellType() == CellType.STRING)
            {
               String cellValue = cell.getStringCellValue().trim();
               if (cellValue.toLowerCase().equals("title") || cellValue.toLowerCase().equals("titles"))
               {
                  hasTitle = true;
                  abstractColumn = i + 1;
                  doiColumn = i + 2; // updates the values of where the columns should be since we are putting it after the title.
               }
               if (cellValue.toLowerCase().equals("abstract") || cellValue.toLowerCase().equals("abstracts"))
                  hasAbstractColumn = true;
               if (cellValue.toLowerCase().equals("doi") || cellValue.toLowerCase().equals("dois"))
                  hasDOIColumn = true;
            }
         }
         if (Boolean.TRUE.equals(hasTitle))
         {
            if (Boolean.FALSE.equals(hasAbstractColumn))
            {
               sheet.shiftColumns(abstractColumn, noOfColumns, 1);
               sheet.getRow(0).createCell(abstractColumn, CellType.STRING).setCellStyle(sheet.getRow(0).getCell(0).getCellStyle()); // creates the cell with the specified cell style
               sheet.getRow(0).getCell(abstractColumn).setCellValue("Abstract"); // Then we add the desired attribute name to the cell
               System.out.println("An abstract column has been inserted for each sheet with a title column with no abstract column already existing.");
            }
            if (Boolean.FALSE.equals(hasDOIColumn))
            {
               sheet.shiftColumns(doiColumn, noOfColumns, 1);
               sheet.getRow(0).createCell(doiColumn, CellType.STRING).setCellStyle(sheet.getRow(0).getCell(0).getCellStyle()); // creates the cell with the specified cell style
               sheet.getRow(0).getCell(doiColumn).setCellValue("DOI"); // Then we add the desired attribute name to the cell
               System.out.println("A DOI column has been inserted for each sheet with a title column with no DOI column already existing.");
            }
         }  
      }
      return wb;
   }
   
   // make the new return type ArrayList<ArrayList<String>> to return two lists, so we can make this program a bit more efficient
   // need to figure out how to differentiate between what threw an error (we now are doing two things that can throw a NullPointerException 
   // and we don't want the error to crash the program because of the amount of entries we have --> right now, I'm just making it set the text for both to not having
   // one, even if it does have one, but not the other
   // one idea --> have a boolean variable for each to see how far it made it into the program
   public static ArrayList<ArrayList<String>> RetrieveData(ArrayList<String> searchFor)
   {

      String abstracttext = " "; // Will be overwritten by the abstract if we succeed.
      
      String doiText = " "; // Will be overwritten by the DOI if we succeed
      
      Document doc; // creates a new Document object that we will use to extract the html page and then extract the abstract text
      
      ArrayList<String> abstractList = new ArrayList<String>(); // creates a list that we will store our abstracts in
      
      ArrayList<String> doiList = new ArrayList<String>(); // creates a list to store the doi in
      
      ArrayList<ArrayList<String>> returnedList = new ArrayList<ArrayList<String>>();

      if (searchFor == null)
      {
         returnedList = null;
         return returnedList;
      }
      
      boolean hasAbstract = false, hasDOI = false; //declares and initializes them
      
      try 
      {
         /*
          *           Our current searchstring uses the author's name + the name of an article from our searchFor list.
          *           This currently works for most cases since all test cases provide 1 search result
          *           If the search query needs to be improved, the documentation below will help: 
          *           https://pubmed.ncbi.nlm.nih.gov/help/#citation-matcher-auto-search
          */
         
         String searchString = "";     
         
         for(int i = 1; i < searchFor.size(); i++)
         {
            try {
               
               hasAbstract = false; //resets them for each loop in case they got set to true
               hasDOI = false;
               
               searchString = searchFor.get(0) + " " + searchFor.get(i);
               
               /*
                *                Uses the searchString (author's name + title of article) to search PubMed using a heuristic search on PubMed
                *                This most likely fails, but since PubMed defaults to an auto search if the heuristic search fails
                *                It still allows us to find the article we are searching for
                */
               doc = Jsoup.connect("https://pubmed.ncbi.nlm.nih.gov/?term=" + java.net.URLEncoder.encode(searchString, "UTF-8")).get(); 
               
               // Selects the id "abstract" and look for the paragraph element of the first occurrence of the id abstract
               // In theory, this should not cause an issue, since only one HTML element is allowed to have the id abstract
               // Meaning it should be okay to only search for the first occurrence
               // More documentation: https://jsoup.org/apidocs/org/jsoup/nodes/Element.html#selectFirst(java.lang.String)
               // More documentation: https://jsoup.org/cookbook/extracting-data/selector-syntax
               
               Element abstractelement = doc.selectFirst("#abstract p");
               
               Element doiElement = doc.selectFirst("span.identifier.doi a");
               
               abstracttext = abstractelement.text(); // gets only the text of the abstract from the paragraph (<p>) HTML element
               // For more info: https://jsoup.org/apidocs/org/jsoup/nodes/Element.html#text(java.lang.String)
               
               abstractList.add(abstracttext);
               
               hasAbstract = true; //if we make it to this part, we will have an abstract (no exception thrown)
               
               doiText = doiElement.text();
               
               doiList.add(doiText);
               
               hasDOI = true; //if we make it to this part, we will have a DOI (no exception thrown)
               
            }
            catch (NullPointerException npe) { //need to implement the boolean checking 
               /*
                *               abstracttext = "no abstract on PubMed";
                *               abstractList.add(abstracttext);
                *               doiText = "no doi on PubMed";
                *               doiList.add(doiText);
                */
               if (Boolean.FALSE.equals(hasAbstract))
               {
                  abstracttext = "no abstract on PubMed";
                  abstractList.add(abstracttext);
                  try 
                  {
                     doiText = RetrieveDOI(searchString);
                  }
                  catch (Exception e)
                  {
                     e.printStackTrace();
                  }
                  doiList.add(doiText);
               }
               if (Boolean.FALSE.equals(hasDOI) && hasAbstract) // the abstractList will already have the abstract so you do not want to double add it to the list
               {
                  doiText = "no doi on PubMed";
                  doiList.add(doiText);
               }
            }
            catch (MalformedURLException mue) {
               mue.printStackTrace();
               abstracttext = "error";
               abstractList.add(abstracttext);
               doiText = "error";
               doiList.add(doiText);
            }
         }
      } catch (IOException e) {
         e.printStackTrace();
      }
      int count = 0, doiCount = 0;
      for (int k = 0; k < abstractList.size(); k++)
      {
         if (abstractList.get(k).equals("no abstract on PubMed"))
            count++;
         if (doiList.get(k).equals("no doi on PubMed"))
            doiCount++;
      }
      /* 
       * We now display information about the doi / abstract found.
       * In comment, the code to display information about the doi / abstract *not* found.
       */
      // System.out.println("Number of publications that did not have an abstract for the sheet with the author \'" + searchFor.get(0) + "\' on PubMed: " + count + "/" + abstractList.size());
      System.out.println("Number of publications that had an abstract for the sheet with the author \'" + searchFor.get(0) + "\'' on PubMed: " + (abstractList.size() - count) + "/" + abstractList.size());
      // System.out.println("Number of publications that did not have a DOI for the sheet with the author \'" + searchFor.get(0) + "\' on PubMed: " + doiCount + "/" + doiList.size());
      System.out.println("Number of publications that had a DOI for the sheet with the author \'" + searchFor.get(0) + "\' on PubMed: " + (doiList.size() - doiCount) + "/" + doiList.size());
      
      returnedList.add(abstractList);
      returnedList.add(doiList);
      
      return returnedList;
   }
   
   public static ArrayList<String> Read_From_Excel(int sheetIndex, File inputPath) throws IOException, Exception{
      
      ArrayList<String> searchList = new ArrayList<String>();
      
      FileInputStream fins = new FileInputStream(inputPath);
      
      XSSFWorkbook wb = new XSSFWorkbook(fins); // creates a workbook that we can search, which allows us to get the author's name and the titles of each publication
      if (wb.getNumberOfSheets() < sheetIndex){
         System.out.println("Inside Read_From_Excel, something is amiss. The sheet has" + wb.getNumberOfSheets() + "sheets, but I am looking for sheet #" + sheetIndex);
      }
      
      /*
       *  Parameter sheetIndex gives sheet to be extracted.
       */
      XSSFSheet sheet = wb.getSheetAt(sheetIndex);
      System.out.println("The program is currently working on the sheet: " + sheet.getSheetName());
      
      int rows = sheet.getLastRowNum(); // gets number of rows
      int cols = sheet.getRow(0).getLastCellNum(); // gets the number of columns
      
      XSSFRow row = sheet.getRow(0); // starting the row at 0 for current sheet.
      
      boolean hasTitle = false;
      for (int i = 0; i < cols; i++)
      {
         XSSFCell cell = row.getCell(i);
         if (cell == null)
            continue;
         if (cell.getCellType() == CellType.STRING)
         {
            String cellValue = cell.getStringCellValue().trim();
            if (cellValue.toLowerCase().equals("title") || cellValue.toLowerCase().equals("titles"))
               hasTitle = true;
         }
      }

      if (hasTitle) // if the sheet has a title column
      {
         for (int i = 0; i < cols; i++)
         {
            XSSFCell cell = row.getCell(i);
         
            // tests if the cell is null, since testing the cell type would throw an error if null (<= this is not the case, we are not throwing an error. Why?)
            // This is only intended for the titles of each column in our target excel file, since we will not need data from any other column

            if (cell == null)
               continue; // Well, we don't actually throw an error. Why not?
               if (cell.getCellType() == CellType.STRING)
               {
                  String cellValue = cell.getStringCellValue().trim(); // gets the value of the cell if it is a string value
                  if (cellValue.toLowerCase().equals("researcher") || cellValue.toLowerCase().equals("researchers") || cellValue.toLowerCase().equals("author") || cellValue.toLowerCase().equals("authors")) //if the value of the cell is equal to "researcher", then we get the name of that researcher
                  {
                     XSSFRow tempRow = sheet.getRow(1);
                     XSSFCell tempCell = tempRow.getCell(i); //creating temp objects so we do not accidentally shift the row and cells, since we still need the titles
                     cellValue = tempCell.getStringCellValue().trim();
                     searchList.add(cellValue);
                  }
                  if (cellValue.toLowerCase().equals("title") || cellValue.toLowerCase().equals("titles"))
                  {
                     for (int j = 1; j <= rows; j++)
                     {
                        row = sheet.getRow(j);
                        cell = row.getCell(i);
                        if (cell == null)
                           break;
                        cellValue = cell.getStringCellValue().trim(); // loops through each cell in the specified "title" column until we have all the titles in our list
                        searchList.add(cellValue);
                     }
                  }
               }
         }
      }
      else
      { 
         System.out.println(sheet.getSheetName() + " was skipped due to not having a title column.");
         searchList = null;
      }

      fins.close(); //closes the inputstream
      wb.close();
      // the author's name will always be the first index followed by the titles
      return searchList;
   }
   
   public static void Write_To_Excel(ArrayList<ArrayList<String>> writeList, int sheetIndex, File outputPath) throws Exception {
      try {

         if (writeList == null)
            return;

         ArrayList<String> abstractList = writeList.get(0);
         ArrayList<String> doiList = writeList.get(1); 
         FileInputStream fins = new FileInputStream(outputPath);
         XSSFWorkbook wb = new XSSFWorkbook(fins);
         if (wb.getNumberOfSheets() < sheetIndex){
            System.out.println("Indiside Write_To_Excel, something is amiss. The sheet has" + wb.getNumberOfSheets() + "sheets, but I am looking for sheet #" + sheetIndex);
         }
         
         /*
          *  Parameter sheetIndex gives sheet to be extracted.
          */
         XSSFSheet sheet = wb.getSheetAt(sheetIndex);
         
         // SocketTimeoutException causing it to not work in some cases
         // Using the size of the list allows us to still run the code
         // May need to rerun since this is often caused by connection issues
         int rows = abstractList.size(); // SocketTimeoutException causing it to not work in some cases
         int doiRows = doiList.size();
         int cols = sheet.getRow(0).getLastCellNum(); //gets the number of columns in the sheet
         
         XSSFRow row = sheet.getRow(0);
         
         for(int i = 0; i < cols; i++)
         {
            XSSFCell cell = row.getCell(i);
            if (cell == null) // if the cell is null for whatever reason, it will throw an error when trying to get the cell type
               continue; // Well, we don't actually throw an error. Why not?
               if (cell.getCellType() == CellType.STRING)
               {
                  String valueOfCell = cell.getStringCellValue().trim();
                  
                  if (valueOfCell.toLowerCase().equals("abstract") || valueOfCell.toLowerCase().equals("abstracts"))
                  {
                     for (int j = 1; j <= rows; j++)
                     {
                        int abIndex = j - 1; // allows us to access the correct abstract in our list
                        row = sheet.getRow(j); // sets us on the right row 
                        row.createCell(i, CellType.STRING).setCellValue(abstractList.get(abIndex));
                        // we then "create" a cell which has a cell type of String, which allows us to write our abstract to the cell.
                     }
                     row = sheet.getRow(0); //sets the row back to 0 after running
                  }
                  if (valueOfCell.toUpperCase().equals("DOI") || valueOfCell.toUpperCase().equals("DOIS"))
                  {
                     for (int l = 1; l <= doiRows; l++)
                     {
                        int doiIndex = l - 1; // allows us to access the correct abstract for each row, since row would be 1 more than the actual index
                        row = sheet.getRow(l); // sets us on the correct row
                        row.createCell(i, CellType.STRING).setCellValue(doiList.get(doiIndex));
                     }
                     row = sheet.getRow(0); //sets the row back to 0 after running
                  }
               }
         }
         FileOutputStream fos = new FileOutputStream(outputPath);
         wb.write(fos);
         fos.close();
         wb.close();
      }
      catch (Exception e) {
         e.printStackTrace();
      }   
   }
   
   public static String RetrieveDOI(String search) throws Exception {
      String searchString = search;
      String doiText = " ";
      Document doc;
      try
      {
         doc = Jsoup.connect("https://pubmed.ncbi.nlm.nih.gov/?term=" + java.net.URLEncoder.encode(searchString, "UTF-8")).get();
         Element doiElement = doc.selectFirst("span.identifier.doi a");
         doiText = doiElement.text();
      }
      catch (NullPointerException npe) 
      {
         doiText = "no doi on PubMed";
         return doiText;
      }
      return doiText;
   }
}
