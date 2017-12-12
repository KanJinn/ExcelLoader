using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using ExcelApp = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;

namespace ExcelLoader
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            OpenFileDialog dialog = new OpenFileDialog();
            //dialog.Filter = "Excel Files|*.xls;*.xlsx";
            dialog.Filter = "CSV Files|*.csv";
            dialog.Multiselect = false;
            dialog.InitialDirectory = Directory.GetCurrentDirectory();
            //dialog.Title = "Select Excel File";
            dialog.Title = "Select CSV File";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string fname = dialog.FileName;

                readCSV(fname, fname); //Read CSV Files

             }

            else
            {
                MessageBox.Show("Please select excel file.");
            }

        }


        
        public void readCSV(String fileName, String excelFileName)
        {
            using (var reader = new StreamReader(fileName))
            {

                List<string> entryForSplit = new List<string>();
                while (!reader.EndOfStream)
                {
                    entryForSplit.Add(reader.ReadLine());
                }

                createExcelFile(fileName, entryForSplit);

            }
        }

        public void createExcelFile(String excelFileName, List<string> entryForSplit)
        {

            ExcelApp.Application xlApp;
            ExcelApp.Workbook xlBook;
            ExcelApp.Worksheet xlSheet;
            ExcelApp.Range range;

            xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }

            try
            {

                xlApp.Visible = true;;
                xlApp.DisplayAlerts = false;
                xlBook = xlApp.Workbooks.Add(1);
                //xlSheet = xlApp.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                xlSheet = xlApp.Worksheets[1];

                //var xlSheets = xlBook.Sheets as ExcelApp.Sheets;
                //var xlNewSheet = (ExcelApp.Worksheet)ExcelApp.Sheets.Add(xlSheets[xlBook.Sheets.Count], Type.Missing, Type.Missing, Type.Missing);
                //xlNewSheet.Name = "new shit"

                uint processId = 0;
                GetWindowThreadProcessId(new IntPtr(xlApp.Hwnd), out processId);

                xlSheet.Name = "ConsolidatedPaymentReceipt";


                xlSheet.get_Range("A:A").EntireColumn.Hidden = true;
                xlSheet.get_Range("T:T").EntireColumn.Hidden = true;

                range = xlSheet.get_Range("B2", "G2");
                range.Merge();
                xlSheet.Cells["2", "B"].Value2 = "Consolidated Payment Receipt";

                range = xlSheet.get_Range("B4", "C4");
                range.Merge();
                xlSheet.Cells["4", "B"].Value2 = "Finance Organization Code";

                xlSheet.Cells["4", "K"].Value2 = "Location Code";

                range = xlSheet.get_Range("B6", "C6");
                range.Merge();
                xlSheet.Cells["6", "B"].Value2 = "Currency Code";

                xlSheet.Cells["6", "K"].Value2 = "Department";

                range = xlSheet.get_Range("B8", "C8");
                range.Merge();
                xlSheet.Cells["8", "B"].Value2 = "FOP";

                xlSheet.Cells["8", "K"].Value2 = "Payment Date";

                xlSheet.Cells["4", "E"].Value2 = ":";
                xlSheet.Cells["6", "E"].Value2 = ":";
                xlSheet.Cells["8", "E"].Value2 = ":";

                xlSheet.Cells["4", "M"].Value2 = ":";
                xlSheet.Cells["6", "M"].Value2 = ":";
                xlSheet.Cells["8", "M"].Value2 = ":";

                xlSheet.Cells["10", "B"].Value2 = "Organization Name";
                range = xlSheet.get_Range("B10","B10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xlSheet.Cells["10", "C"].Value2 = "Agent Code";
                range = xlSheet.get_Range("C10", "F10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                range.Merge();


                xlSheet.Cells["10", "G"].Value2 = "First Name";
                range = xlSheet.get_Range("G10", "L10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                range.Merge();

                xlSheet.Cells["10", "M"].Value2 = "Last Name";
                range = xlSheet.get_Range("M10", "O10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                range.Merge();

                xlSheet.Cells["10", "P"].Value2 = "Record Locator";
                range = xlSheet.get_Range("P10", "Q10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                range.Merge();

                xlSheet.Cells["10", "R"].Value2 = "Created Organisation Code";
                range = xlSheet.get_Range("R10", "R10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                
                xlSheet.Cells["10", "T"].Value2 = "Source Organisation Code";
                range = xlSheet.get_Range("T10", "T10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xlSheet.Cells["10", "U"].Value2 = "Payment Code";
                range = xlSheet.get_Range("U10", "U10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xlSheet.Cells["10", "V"].Value2 = "Payment Id";
                range = xlSheet.get_Range("V10", "V10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xlSheet.Cells["10", "W"].Value2 = "Authorization Status";
                range = xlSheet.get_Range("W10", "W10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xlSheet.Cells["10", "X"].Value2 = "Currency Code";
                range = xlSheet.get_Range("X10", "X10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xlSheet.Cells["10", "Y"].Value2 = "Booking Amount";
                range = xlSheet.get_Range("Y10", "Y10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xlSheet.Cells["10", "Z"].Value2 = "Collected Currency Code";
                range = xlSheet.get_Range("Z10", "Z10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xlSheet.Cells["10", "AA"].Value2 = "Collected Amount";
                range = xlSheet.get_Range("AA10", "AA10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xlSheet.Cells["10", "AB"].Value2 = "Converted Currency Code";
                range = xlSheet.get_Range("AB10", "AB10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xlSheet.Cells["10", "AC"].Value2 = "Converted Amount";
                range = xlSheet.get_Range("AC10", "AC10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xlSheet.Cells["10", "AD"].Value2 = "Payment Text";
                range = xlSheet.get_Range("AD10", "AD10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xlSheet.Cells["10", "AE"].Value2 = "Passenger Name";
                range = xlSheet.get_Range("AE10", "AF10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                range.Merge();

                xlSheet.Cells["10", "AG"].Value2 = "Route";
                range = xlSheet.get_Range("AG10", "AH10");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                range.Merge();

                int rowCount = 0;
                int baseCount = 7;
                int numberOfOrganisations = 0;
                string forComparison = "";
                string billingDate = "";

                List<string> organisationListCode = new List<string>();
                List<string> organisationListNames = new List<string>();

                List<string> sheetList = new List<string>();
                
                foreach (var entry in entryForSplit)
                {
                    List<string> OrganisationList = new List<string>();
                    var splitData = SplitCSV(entry);
                    String[] line = new String[32];
                    for (int count = 0; count < 32; count++)
                    {

                        try
                        {
                            if (splitData[count] != null)
                            {
                                line[count] = splitData[count];
                            }

                            else
                            {
                                line[count] = "";
                            }
                        }

                        catch (SystemException e)
                        {

                            line[count] = "";
                        }


                    }



                    String organisationName = line[0];
                    String agentName = line[1];
                    String firstName = line[2];
                    String lastName = line[3];
                    String recordLocator = line[4];
                    String createdOrganisation = line[5];
                    String sourceOrganisation = line[6];
                    String paymentMethodCode = line[7];
                    String paymentID = line[8];
                    String authorisationStatus = line[9];
                    String currencyCode = line[10];
                    String paymentAmount = line[11];
                    String collectedCurrencyCode= line[12];
                    String collectedAmount = line[13];
                    String toCurrencyCode = line[14];
                    String convertedAmount = line[15];
                    String paymentText = line[16];
                    String paxFname = line[17];
                    String paxLname = line[18];
                    String departureStation = line[19];
                    String arrivalStation = line[20];
                    String blankBox70 = line[21];
                    String blankBox72 = line[22];
                    String blankBox73 = line[23];
                    String blankBox74 = line[24];
                    String blankBox75 = line[25];
                    String blankBox76 = line[26];
                    String convertedAmountDup = line[27];
                    String blankBox85 = line[28];
                    String blankBox86 = line[29];
                    String blankBox87 = line[30];
                    String blankBox45 = line[31];


                    sheetList.Add(

                        organisationName + "|" +
                        agentName + "|" +
                        firstName + "|" +
                        lastName + "|" +
                        recordLocator + "|" +
                        createdOrganisation + "|" +
                        paymentMethodCode + "|" +
                        currencyCode + "|" +
                        paymentAmount + "|" +
                        paxFname + "|" +
                        paxLname + "|" +
                        departureStation + "|" +
                        arrivalStation

                        );

                    
                    if (rowCount == 1)
                    {
                        xlSheet.Cells["4", "F"].Value2 = organisationName;
                        xlSheet.Cells["8", "N"].Value2 = createdOrganisation;
                        xlSheet.Cells["6", "F"].Value2 = firstName;
                        //xlSheet.Cells[""]
                        billingDate = createdOrganisation;
                        rowCount++;
                    }

                    else if (rowCount > 3)
                    {

                        if (rowCount.Equals(4))
                        {
                            forComparison = createdOrganisation;
                            organisationListCode.Add(forComparison);
                            organisationListNames.Add(organisationName + "|" + createdOrganisation);
                            numberOfOrganisations++;
                        }

                        if(!forComparison.Equals(createdOrganisation) && (rowCount>4))
                        {
                            forComparison = createdOrganisation;
                            organisationListCode.Add(forComparison);
                            organisationListNames.Add(organisationName + "|" + createdOrganisation);
                            numberOfOrganisations++;
                        }

                        xlSheet.Cells[baseCount + rowCount, "B"].Value2 = organisationName;
                        range = xlSheet.get_Range("B" + baseCount + rowCount, "B" + baseCount + rowCount);
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        xlSheet.Cells[baseCount + rowCount, "C"].Value2 = agentName;
                        range = xlSheet.get_Range("C" + Convert.ToString(baseCount + rowCount), "F" + Convert.ToString(baseCount + rowCount));
                        range.Merge();

                        xlSheet.Cells[baseCount + rowCount, "G"].Value2 = firstName;
                        range = xlSheet.get_Range("G" + Convert.ToString(baseCount + rowCount), "L" + Convert.ToString(baseCount + rowCount));
                        range.Merge();
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        xlSheet.Cells[baseCount + rowCount, "M"].Value2 = lastName;
                        range = xlSheet.get_Range("M" + Convert.ToString(baseCount + rowCount), "O" + Convert.ToString(baseCount + rowCount));
                        range.Merge();
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        xlSheet.Cells[baseCount + rowCount, "P"].Value2 = recordLocator;
                        range = xlSheet.get_Range("P" + Convert.ToString(baseCount + rowCount), "Q" + Convert.ToString(baseCount + rowCount));
                        range.Merge();
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        xlSheet.Cells[baseCount + rowCount, "R"].Value2 = createdOrganisation;
                        range = xlSheet.get_Range("R" + Convert.ToString(baseCount + rowCount), "S" + Convert.ToString(baseCount + rowCount));
                        range.Merge();
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        xlSheet.Cells[baseCount + rowCount, "T"].Value2 = sourceOrganisation;
                        range = xlSheet.get_Range("T" + Convert.ToString(baseCount + rowCount), "T" + Convert.ToString(baseCount + rowCount));
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        xlSheet.Cells[baseCount + rowCount, "U"].Value2 = paymentMethodCode;
                        range = xlSheet.get_Range("U" + Convert.ToString(baseCount + rowCount), "U" + Convert.ToString(baseCount + rowCount));
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        xlSheet.Cells[baseCount + rowCount, "V"].Value2 = paymentID;
                        range = xlSheet.get_Range("V" + Convert.ToString(baseCount + rowCount), "V" + Convert.ToString(baseCount + rowCount));
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        xlSheet.Cells[baseCount + rowCount, "W"].Value2 = authorisationStatus;
                        range = xlSheet.get_Range("W" + Convert.ToString(baseCount + rowCount), "W" + Convert.ToString(baseCount + rowCount));
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        xlSheet.Cells[baseCount + rowCount, "X"].Value2 = currencyCode;
                        range = xlSheet.get_Range("X" + Convert.ToString(baseCount + rowCount), "X" + Convert.ToString(baseCount + rowCount));
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        xlSheet.Cells[baseCount + rowCount, "Y"].Value2 = paymentAmount.Replace("\"", "");
                        range = xlSheet.get_Range("Y" + Convert.ToString(baseCount + rowCount), "Y" + Convert.ToString(baseCount + rowCount));
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        xlSheet.Cells[baseCount + rowCount, "Z"].Value2 = collectedCurrencyCode;
                        //range = xlSheet.get_Range("Z" + Convert.ToString(baseCount + rowCount), "Z" + Convert.ToString(baseCount + rowCount));
                        //range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        xlSheet.Cells[baseCount + rowCount, "AA"].Value2 = collectedAmount.Replace("\"", "");
                        range = xlSheet.get_Range("AA" + Convert.ToString(baseCount + rowCount), "AA" + Convert.ToString(baseCount + rowCount));
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        xlSheet.Cells[baseCount + rowCount, "AB"].Value2 = toCurrencyCode;
                        range = xlSheet.get_Range("AB" + Convert.ToString(baseCount + rowCount), "AB" + Convert.ToString(baseCount + rowCount));
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        xlSheet.Cells[baseCount + rowCount, "AC"].Value2 = convertedAmount.Replace("\"","");

                        xlSheet.Cells[baseCount + rowCount, "AD"].Value2 = paymentText;

                        xlSheet.Cells[baseCount + rowCount, "AE"].Value2 = paxFname;
                        range = xlSheet.get_Range("AE" + Convert.ToString(baseCount + rowCount), "AE" + Convert.ToString(baseCount + rowCount));
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;


                        xlSheet.Cells[baseCount + rowCount, "AF"].Value2 = paxLname;
                        range = xlSheet.get_Range("AF" + Convert.ToString(baseCount + rowCount), "AF" + Convert.ToString(baseCount + rowCount));
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;


                        xlSheet.Cells[baseCount + rowCount, "AG"].Value2 = departureStation;
                        range = xlSheet.get_Range("AG" + Convert.ToString(baseCount + rowCount), "AG" + Convert.ToString(baseCount + rowCount));
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        xlSheet.Cells[baseCount + rowCount, "AH"].Value2 = arrivalStation;
                        range = xlSheet.get_Range("AH" + Convert.ToString(baseCount + rowCount), "AH" + Convert.ToString(baseCount + rowCount));
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        
                        rowCount++;
                    }
                    else
                    {
                        rowCount++;
                    }
                }

               

                int sheetCount = 1;

                List<string> summaryTotal = new List<string>();

                foreach (var organisation in organisationListNames)
                {
                    String organisationName = organisation.Split('|')[0];
                    String createdOrgSplit = organisation.Split('|')[1];

                    String sheetName = "";
                    
                    xlSheet = xlApp.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    sheetName = checkStringLength(organisationName);

                    try
                    {
                        bool found = false;
                        foreach(ExcelApp.Worksheet sheet in xlApp.Worksheets)
                        {
                            if (sheet.Name == sheetName)
                            {
                                found = true;
                                break;
                            }
                        }

                        if (found)
                        {
                            sheetName = duplicateSheetName(sheetCount + sheetName);
                            xlSheet.Name = sheetName.Remove('\"');
                            sheetCount++;
                        }

                        else
                        {
                            

                            xlSheet.Name = sheetName; //Addition of new sheets

                            xlSheet.Shapes.AddPicture(Directory.GetCurrentDirectory() + @"\cebupac.png", MsoTriState.msoFalse, MsoTriState.msoCTrue, 50, 50, 300, 45);
                            rowCount = 0;

                            foreach (string list in sheetList)
                            {
                                if (rowCount == 0)
                                    {
                                        xlSheet.Cells["8", "B"].Value2 = billingDate;
                                        


                                        xlSheet.Cells["10", "B"].Value2 = "Organization Name";
                                        range = xlSheet.get_Range("B10", "B10");
                                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;


                                        xlSheet.Cells["10", "C"].Value2 = "Agent Code";
                                        range = xlSheet.get_Range("C10", "F10");
                                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                        range.Merge();

                                        xlSheet.Cells["10", "G"].Value2 = "First Name";
                                        range = xlSheet.get_Range("G10", "L10");
                                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                        range.Merge();

                                        xlSheet.Cells["10", "M"].Value2 = "Last Name";
                                        range = xlSheet.get_Range("M10", "O10");
                                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                        range.Merge();

                                        xlSheet.Cells["10", "P"].Value2 = "Record Locator";
                                        range = xlSheet.get_Range("P10", "Q10");
                                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                        range.Merge();

                                        xlSheet.Cells["10", "R"].Value2 = "Created Organization";
                                        range = xlSheet.get_Range("R10", "S10");
                                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                        range.Merge();

                                        xlSheet.get_Range("T:T").EntireColumn.Hidden = true;

                                        xlSheet.Cells["10", "U"].Value2 = "Payment Code";

                                        xlSheet.get_Range("V:V").EntireColumn.Hidden = true;

                                        xlSheet.get_Range("W:W").EntireColumn.Hidden = true;

                                        xlSheet.Cells["10", "X"].Value2 = "Currency Code";

                                        xlSheet.Cells["10", "Y"].Value2 = "Booking Amount";

                                        xlSheet.get_Range("Z:Z").EntireColumn.Hidden = true;

                                        xlSheet.get_Range("AA:AA").EntireColumn.Hidden = true;

                                        xlSheet.get_Range("AB:AB").EntireColumn.Hidden = true;

                                        xlSheet.get_Range("AC:AC").EntireColumn.Hidden = true;

                                        xlSheet.get_Range("AD:AD").EntireColumn.Hidden = true;

                                        xlSheet.Cells["10", "AE"].Value2 = "Passenger Name";
                                        range = xlSheet.get_Range("AE10", "AF10");
                                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                        range.Merge();

                                        xlSheet.Cells["10", "AG"].Value2 = "Route";
                                        range = xlSheet.get_Range("AG10", "AH10");
                                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                        range.Merge();


                                    rowCount++;
                                    }

                                    else
                                    {
                                    if (rowCount == 4) 
                                    {
                                        int rowCounting = 11;
                                        foreach(String splitEntry in sheetList)
                                        {
                                            if(splitEntry.Split('|')[5].Equals(createdOrgSplit))
                                            {
                                                xlSheet.Cells[rowCounting, "B"].Value2 = splitEntry.Split('|')[0];

                                                xlSheet.Cells[rowCounting, "C"].Value2 = splitEntry.Split('|')[1];
                                                range = xlSheet.get_Range("C" + rowCounting, "F" + rowCounting);
                                                range.Merge();
                                                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                                                xlSheet.Cells[rowCounting, "G"].Value2 = splitEntry.Split('|')[2];
                                                range = xlSheet.get_Range("G" + rowCounting, "L" + rowCounting);
                                                range.Merge();
                                                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                                                xlSheet.Cells[rowCounting, "M"].Value2 = splitEntry.Split('|')[3];
                                                range = xlSheet.get_Range("M" + rowCounting, "O" + rowCounting);
                                                range.Merge();
                                                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                                                xlSheet.Cells[rowCounting, "P"].Value2 = splitEntry.Split('|')[4];
                                                range = xlSheet.get_Range("P" + rowCounting, "Q" + rowCounting);
                                                range.Merge();
                                                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                                                xlSheet.Cells[rowCounting, "R"].Value2 = splitEntry.Split('|')[5];
                                                xlSheet.Cells[rowCounting, "U"].Value2 = splitEntry.Split('|')[6];
                                                xlSheet.Cells[rowCounting, "X"].Value2 = splitEntry.Split('|')[7];
                                                xlSheet.Cells[rowCounting, "Y"].Value2 = (splitEntry.Split('|')[8]).Replace("\"", "");
                                                xlSheet.Cells[rowCounting, "AE"].Value2 = splitEntry.Split('|')[9];
                                                xlSheet.Cells[rowCounting, "AF"].Value2 = splitEntry.Split('|')[10];
                                                xlSheet.Cells[rowCounting, "AG"].Value2 = splitEntry.Split('|')[11];
                                                xlSheet.Cells[rowCounting, "AH"].Value2 = splitEntry.Split('|')[12];




                                                rowCounting++;
                                            }
                                            
                                        }
                                        rowCount++;
                                    }

                                    else
                                    {
                                        rowCount++;
                                    }
                                    

                                        
                                    }

                                }

                            range = xlSheet.get_Range("B" + Convert.ToString(baseCount), "AH" + Convert.ToString(baseCount + rowCount));
                            range.Columns.AutoFit();



                        }


                    }

                    catch(Exception e)
                    {
                        MessageBox.Show(Convert.ToString(e));
                    }

                    
                }
                

                xlBook.SaveAs(excelFileName + "_Converted.xlsx", ExcelApp.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, ExcelApp.XlSaveAsAccessMode.xlNoChange, ExcelApp.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);

                MessageBox.Show(excelFileName + "_Converted.xlsx has been generated.");

                try
                {
                    if (processId != 0)
                    {
                        Process excelProcess = Process.GetProcessById((int)processId);
                        excelProcess.CloseMainWindow();
                        excelProcess.Refresh();
                        excelProcess.Kill();
                    }
                }
                catch
                {
                    // Process was already killed
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();


            }
            catch (Exception e)
            {
                MessageBox.Show(Convert.ToString(e));
            }
        }



        public static string[] SplitCSV(string input)
        {
            Regex csvSplit = new Regex("(?:^|,)(\"(?:[^\"]+|\"\")*\"|[^,]*)", RegexOptions.Compiled);
            List<string> list = new List<string>();
            string curr = null;
            foreach (Match match in csvSplit.Matches(input))
            {
                curr = match.Value;
                if (0 == curr.Length)
                {
                    list.Add("");
                }

                list.Add(curr.TrimStart(','));
            }

            return list.ToArray<string>();
        }

        public static string checkStringLength(string input)
        {

            if ((String.IsNullOrEmpty(input)) || (String.IsNullOrWhiteSpace(input)))
            {
                input = "blank";
                return input;
            }

            else
            {
                //int index = input.IndexOf(" ");
                //input = index >= 0 ? input.Substring(input.IndexOf(" ")+ 1) : input;

                int length = input.Length;
                

                if (length > 30)
                {
                    int temp = length - 30;

                    length = length - temp;

                    input = input.Substring(0, length);
                    return input;
                }

                else if ((length >= 10) && (length <= 30))
                {
                    input = input.Substring(0, length);
                    return input;
                }

                else
                {
                    return input;


                }

            }

        }

        public static string duplicateSheetName(String input)
        {
            input = "Dup_" + input;
            int length = input.Length;

            if (length > 30)
            {
                int temp = length - 30;

                length = length - temp;

                input = input.Substring(0, length);
                return input;
            }

            else
            {
                return input;
            }
            
        }
       


    }
}
