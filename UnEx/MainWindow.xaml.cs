using System;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.Win32;
using Ionic.Zip;


namespace UnEx//Excel Unlocker//Unprotect Excel Worksheets//@SlavenHvar
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            string url = openFileDialog.FileName;
            string ZIP_url;
            try
            {
                #region 1. Get the index of the file extension dot and check the file extension
                int idx = url.LastIndexOf('.');
                string fileExtension = url.Substring(idx);
                string[] excelExt = { ".xls", ".xlsx", ".xlsm" };
                if(excelExt.Contains(fileExtension))
                {
                    string new_url;//New url of the copied file
                    #endregion

                    #region 2. Make a copy of the original protected excel file

                    Excel.Application xlApp;
                    Excel.Workbook xlWorkBook;
                    object misValue = System.Reflection.Missing.Value;
                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(url, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    bool oldMacro = false;

                    if (string.Equals(url.Split('.').Last(), "xlsm", StringComparison.CurrentCultureIgnoreCase))//macro enabled file
                    {
                        new_url = Methods.SaveAs_xlsm(url, idx, xlWorkBook, xlApp, misValue);
                    }
                    else if (string.Equals(url.Split('.').Last(), "xls", StringComparison.CurrentCultureIgnoreCase))
                    {
                        MessageBoxResult result = MessageBox.Show("Is your xls file a macro enabled workbook ", "Type of xls file", MessageBoxButton.YesNo, MessageBoxImage.Question);
                        if (result == MessageBoxResult.Yes)// check if a old excel file has macro enabled content
                        {
                            new_url = Methods.SaveAs_xlsm(url, idx, xlWorkBook, xlApp, misValue);
                            oldMacro = true;
                        }
                        else
                        {
                            new_url = Methods.SaveAs_xlsx(url, idx, xlWorkBook, xlApp, misValue);
                        }
                    }
                    else
                    {
                        new_url = Methods.SaveAs_xlsx(url, idx, xlWorkBook, xlApp, misValue);
                    }

                    Methods.releaseObject(xlWorkBook);
                    Methods.releaseObject(xlApp);
                    #endregion

                    #region 3. Change the file extension from xlsx (xls,xlsm) to zip
                    File.Move(new_url, System.IO.Path.ChangeExtension(new_url, ".zip"));
                    ZIP_url = (new_url.Substring(0, idx)) + "_copy.zip";//New url with the new file extension
                    #endregion

                    #region 4. Create a new folder called Temp
                    string exe_url = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                    if (!Directory.Exists(exe_url + @"\Temp"))//Check if directory exists
                    {
                        Directory.CreateDirectory(exe_url + @"\Temp");
                    }
                    else//If the directrory exists delete it and create a new one
                    {
                        new DirectoryInfo(exe_url + @"\Temp").Delete(true);
                        Directory.CreateDirectory(exe_url + @"\Temp");
                    }
                    #endregion

                    #region 5. Extract zip archive to the newly created Temp folder
                    using (var zip = ZipFile.Read(ZIP_url))
                    {
                        zip.ExtractAll(exe_url + @"\Temp");
                    }
                    #endregion

                    #region 6. Delete the sheetProtection and the workbookProtection nodes from the xml files inside the Temp folder
                    string folderPath_worksheets = exe_url + @"\Temp\xl\worksheets";
                    string folderPath_workbook = exe_url + @"\Temp\xl";
                    Methods.deleteXMLsheetProtectionNodes(folderPath_worksheets);
                    Methods.deleteXMLworkbookProtectionNodes(folderPath_workbook);
                    #endregion

                    #region 7. Delete the xl folder from the ZIP archive 
                    using (var zip = ZipFile.Read(ZIP_url))
                    {
                        zip.RemoveSelectedEntries("xl/*"); // Remove folder and all its contents
                        zip.Save(ZIP_url);
                    }
                    #endregion

                    #region 8. Add the new xl folder with the modified unprotected xml files to the zip archive
                    using (var zip = ZipFile.Read(ZIP_url))
                    {
                        zip.AddDirectory(exe_url + @"\Temp\xl", "xl");
                        zip.Save(ZIP_url);
                    }
                    #endregion

                    #region 9. Delete the Temp folder
                    new DirectoryInfo(exe_url + @"\Temp").Delete(true);
                    #endregion

                    #region 10. Change the file extansion back from zip to xlsx or xlsm and add _UNPROTECTED to the file name
                    if (string.Equals(url.Split('.').Last(), "xlsm", StringComparison.CurrentCultureIgnoreCase) || oldMacro)//macro file
                    {
                        File.Move(ZIP_url, System.IO.Path.ChangeExtension(ZIP_url, ".xlsm"));
                        File.Move(new_url, new_url.Substring(0, idx) + "_UNPROTECTED.xlsm");
                    }
                    else
                    {
                        File.Move(ZIP_url, System.IO.Path.ChangeExtension(ZIP_url, ".xlsx"));
                        File.Move(new_url, new_url.Substring(0, idx) + "_UNPROTECTED.xlsx");
                    }
                    #endregion

                    MessageBox.Show("Your Excel file is unprotected!");
                }
                else
                {
                    MessageBox.Show("The file you have selected is not an Excel file! Please select again!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
                MessageBox.Show("Your Excel file could not be unprotected");
            }
        }
    }
}