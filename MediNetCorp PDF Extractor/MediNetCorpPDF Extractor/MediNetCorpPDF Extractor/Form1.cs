using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace MediNetCorpPDF_Extractor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            DirectoryClean();
        }

        public void DirectoryClean()
        {
            try
            {
                string ExcelFiles = Path.Combine(Directory.GetCurrentDirectory() + @"\ExcelFiles");
                string[] AllExcelFiles = Directory.GetFiles(ExcelFiles);
                foreach (string file in AllExcelFiles)
                {
                    File.Delete(file);
                }

                string PDFFiles = Path.Combine(Directory.GetCurrentDirectory() + @"\PDFFiles");
                string[] PDFAllfiles = Directory.GetFiles(PDFFiles);
                foreach (string file in PDFAllfiles)
                {
                    File.Delete(file);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void btn_upload_excel_Click(object sender, EventArgs e)
        {
            try
            {
                string excelFileDestination = Path.Combine(Directory.GetCurrentDirectory() + @"\ExcelFiles");               
                OpenFileDialog dlg = new OpenFileDialog();
                dlg.Multiselect = true;
                if (Directory.Exists(excelFileDestination))
                {
                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        foreach (string file in dlg.FileNames)
                        {
                            if (file.Contains(".xlsx") || file.Contains(".xls"))
                            {
                                File.Copy(file, Path.Combine(excelFileDestination, Path.GetFileName(file)), true);
                                MessageBox.Show(file + " uploaded  Successfully");
                                ProcessExcel(file);
                                DirectoryClean();
                            }   
                           
                            else
                            {
                                MessageBox.Show("Upload Excel files only");
                                return;
                            }
                        }
                    }
                }
                else
                {
                    System.IO.Directory.CreateDirectory(Directory.GetCurrentDirectory() + @"\ExcelFiles");
                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        foreach (string file in dlg.FileNames)
                        {
                            if (file.Contains(".xlsx") || file.Contains(".xls"))
                            {
                                File.Copy(file, Path.Combine(excelFileDestination, Path.GetFileName(file)), true);
                                MessageBox.Show(file +" uploaded  Successfully");
                                ProcessExcel(file);    
                                DirectoryClean();
                            }
                            else
                            {
                                MessageBox.Show("Upload Excel files only");
                                return;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
            
        }

        public void ProcessExcel(string file)
        {
            
            try
            {                //"Provider", "Misc",
                //string[] FilterList = new string[] { "Provider", "Misc", "Duplicate & Others" };

                //foreach (string pdf in FilterList)
                //{

                    //create the Application object we can use in the member functions.
                    Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
                    _excelApp.Visible = true;

                    //open the workbook
                    Workbook workbook = _excelApp.Workbooks.Open(file,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

                    //select the first sheet        
                    Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

                    //find the used range in worksheet
                    Range excelRange = worksheet.UsedRange;

                    //get an object array of all of the cells in the worksheet (their values)
                    object[,] valueArray = (object[,])excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);


                   /* dynamic rngHeader = worksheet.UsedRange;
                

                    rngHeader.AutoFilter(6, pdf,
                                    Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);


                    rngHeader.Sort(rngHeader.Columns[5, Type.Missing], Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending, // the first sort key Column 1 for Range
                  rngHeader.Columns[2, Type.Missing], Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,// second sort key Column 6 of the range
                  Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,  // third sort key nothing, but it wants one
                  Microsoft.Office.Interop.Excel.XlYesNoGuess.xlGuess, Type.Missing, Type.Missing,
                  Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns, Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                  Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                  Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                  Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);*/

                    var cellValue = (worksheet.Cells[2, 2] as Microsoft.Office.Interop.Excel.Range).Text;

                    int rowCount = worksheet.UsedRange.Rows.Count;
                    int columnCount = worksheet.UsedRange.Columns.Count;

                    List<string> columnValue = new List<string>();
                    List<string> BookmarkColumnValue = new List<string>();
                    Microsoft.Office.Interop.Excel.Range visibleCells = worksheet.UsedRange.SpecialCells(
                                     Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeVisible,
                               Type.Missing);
                    var dictionary = new List<KeyValuePair<string, string>>();
                    var providerDictionary = new List<KeyValuePair<string, string>>();
                    List<Tuple<string, string, string>> list = new List<Tuple<string,string, string>>();
                    foreach (Microsoft.Office.Interop.Excel.Range area in visibleCells.Areas)
                    {
                        foreach (Microsoft.Office.Interop.Excel.Range row in area.Rows)
                        {
                            if (((Microsoft.Office.Interop.Excel.Range)row.Cells[1, 6]).Text != "Pg#")
                            {

                                columnValue.Add(((Microsoft.Office.Interop.Excel.Range)row.Cells[1, 6]).Text);
                            }

                           /* if (((Microsoft.Office.Interop.Excel.Range)row.Cells[1, 5]).Text != "Provider" || ((Microsoft.Office.Interop.Excel.Range)row.Cells[1, 7]).Text != "Bookmark" || ((Microsoft.Office.Interop.Excel.Range)row.Cells[1, 1]).Text != "Page no.")
                            {  
                              
                                list.Add(new Tuple<string, string, string>(((Microsoft.Office.Interop.Excel.Range)row.Cells[1, 5]).Text, ((Microsoft.Office.Interop.Excel.Range)row.Cells[1, 7]).Text, ((Microsoft.Office.Interop.Excel.Range)row.Cells[1, 1]).Text));                       
                               
                            }*/    

                            if (((Microsoft.Office.Interop.Excel.Range)row.Cells[1, 5]).Text != "Patient Name" || ((Microsoft.Office.Interop.Excel.Range)row.Cells[1, 6]).Text != "Pg#")
                            {

                                dictionary.Add(new KeyValuePair<string, string>(((Microsoft.Office.Interop.Excel.Range)row.Cells[1, 5]).Text, ((Microsoft.Office.Interop.Excel.Range)row.Cells[1, 6]).Text));                       
                            }                        
                        }
                    }

                    string[] arr = new string[] {};
                    arr = columnValue.ToArray();

                    string[] BookmarkColumnArray = new string[] { };
                    BookmarkColumnArray = BookmarkColumnValue.ToArray();

                    List<string> termsList = new List<string>();

                    foreach (string ar in arr)
                    {
                        if (ar.Contains('-') || ar.Contains(','))
                        {
                            termsList.Add(ar);
                        }
                        else
                        {
                            termsList.Add(ar);
                        }
                    }
                    string[] terms = termsList.ToArray();

                    string[] words = new string[] { };
                     int[] termsq = new int [] {};
                     List<int> termsListq = new List<int>();

                    foreach (string ar in terms)
                    {
                        string words1 = ar.Split('-',',').First();
                        string words2 = ar.Split('-',',').Last();

                        int x = Int32.Parse(words1);
                        int y = Int32.Parse(words2);

                      
                        for (int i = x; i <= y; i++)
                        {
                            termsListq.Add(i);
                        }
                        
                        for (int i = 0; i < termsq.Length; i++)
                        {
                            Console.WriteLine(termsq[i]);
                        }
                    }
                    termsq = termsListq.ToArray();

                    string pdf = "pdfname";

                    ExtractPages(pdf, termsq, dictionary, providerDictionary, list);
                    //clean up stuffs                                      
                    workbook.Close(false, Type.Missing, Type.Missing);
                    _excelApp.Quit();
               // }            
               
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
        }

        public void ExtractPages(string pdfname, int[] extractThesePages, List<KeyValuePair<string, string>> dictionary, List<KeyValuePair<string, string>> providerDictionary, List<Tuple<string, string, string>> list)
        {
            //PdfReader reader = null;
           // Document sourceDocument = null;
            //PdfCopy pdfCopyProvider = null;
            //PdfImportedPage importedPage = null;

             string sourcePdfPath = "";

             string destination = Path.Combine(Directory.GetCurrentDirectory() + @"\PDFFiles");
             string[] fileArray = Directory.GetFiles(destination, "*.pdf");
             string currentfilename = "";
            foreach (string file in fileArray)
                        {
                            if (file.Contains(".pdf"))
                            {
                                
                                sourcePdfPath = file;
                                currentfilename = Path.GetFileNameWithoutExtension(file);
                                
                            }
                        }
                    
                
            try
            {
                
                
                int count = 1;
                
                foreach( var i in dictionary)
                {
                    // Intialize a new PdfReader instance with the 
                    // contents of the source Pdf file:
                    PdfReader reader = new PdfReader(sourcePdfPath);

                    // For simplicity, I am assuming all the pages share the same size
                    // and rotation as the first page:
                    Document sourceDocument = new Document(reader.GetPageSizeWithRotation(extractThesePages[0]));

                    // Initialize an instance of the PdfCopyClass with the source 
                    // document and an output file stream:
                    PdfCopy pdfCopyProvider = new PdfCopy(sourceDocument,
                        new System.IO.FileStream("Input File- "+ currentfilename + " Patient Name- " + i.Key +" Page Name- "+ i.Value + " Date- "+ DateTime.Now.ToString("dd-MM-yyyy-hh") + ".pdf", System.IO.FileMode.Create));

                    sourceDocument.Open();
                    ArrayList bookmarks = new ArrayList();
                    string words1 = i.Value.Split('-', ',').First();
                    string words2 = i.Value.Split('-', ',').Last();

                    int x = Int32.Parse(words1);
                    int y = Int32.Parse(words2);
                    if (y < x)
                    {
                       // x = y;
                    }
                    
                    for (int e = x; e <= y; e++)
                    {
                        PdfImportedPage importedPage = pdfCopyProvider.GetImportedPage(reader, e);
                        pdfCopyProvider.AddPage(importedPage);
                       
                        var h = importedPage.Height; // get height of 1st page
                        string page = Convert.ToString(i.Value);
                        // Add first item to bookmarks.                       
                        int pre = e;
                        int current = x+1;                          
                        Hashtable test = new Hashtable();
                        if (pre < current)
                        {
                            test.Add("Action", "GoTo");
                            test.Add("Title", page);
                            if (pdfCopyProvider.CurrentPageNumber == 2)
                            {                                
                                test.Add("Page", count + " XYZ 0 " + h + " 0");
                            }                            
                            else 
                            {
                                test.Add("Page", (pdfCopyProvider.CurrentPageNumber - 2) + " XYZ 0 " + count + " 0");
                            }                           
                            bookmarks.Add(test);
                        }
                        /*else if(pre < current) 
                        {
                            test.Add("Page", (pdfCopyProvider.CurrentPageNumber-1) + " XYZ 0 " + count + " 0");
                        }*/
                        
                        //if (pdfname == "Provider" || pdfname =="Misc")
                        //{
                            pdfCopyProvider.Outlines = bookmarks;
                       // }
               
                    }                 
                    count++;
                    sourceDocument.Close();
                    reader.Close();
                }
               
               
                MessageBox.Show(currentfilename +" " + pdfname + DateTime.Now.ToString("  dd-MM-yyyy-hh") + ".pdf" + " Created Successfully");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void btn_pdf_upload_Click(object sender, EventArgs e)
        {
            try
            {
                string pdfFileDestination = Path.Combine(Directory.GetCurrentDirectory() + @"\PDFFiles");
                OpenFileDialog dlg = new OpenFileDialog();
                dlg.Multiselect = true;
                if (Directory.Exists(pdfFileDestination))
                {
                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        foreach (string file in dlg.FileNames)
                        {
                            if (file.Contains(".pdf"))
                            {
                                File.Copy(file, Path.Combine(pdfFileDestination, Path.GetFileName(file)), true);
                                MessageBox.Show(file + " uploaded  Successfully");
                            }

                            else
                            {
                                MessageBox.Show("Upload Pdf file only");
                                return;
                            }
                        }
                    }
                }
                else
                {
                    System.IO.Directory.CreateDirectory(Directory.GetCurrentDirectory() + @"\PDFFiles");
                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        foreach (string file in dlg.FileNames)
                        {
                            if (file.Contains(".pdf"))
                            {
                                File.Copy(file, Path.Combine(pdfFileDestination, Path.GetFileName(file)), true);
                                MessageBox.Show(file + " uploaded Successfully");
                            }
                            else
                            {
                                MessageBox.Show("Upload Pdf file only");
                                return;
                            }
                        }
                    }
                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }        

    }

}
