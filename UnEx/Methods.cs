using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Xml;
using System.Windows;

namespace UnEx
{
    public static class Methods
    {
        public static string SaveAs_xlsx(string url,int idx,Excel.Workbook xlWorkBook, Excel.Application xlApp, object misValue)
        {
            string new_url = (url.Substring(0, idx)) + "_copy.xlsx";
            xlWorkBook.SaveAs(new_url, Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
            misValue, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
            Excel.XlSaveConflictResolution.xlUserResolution, true,
            misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            return new_url;
        }
        public static string SaveAs_xlsm(string url, int idx, Excel.Workbook xlWorkBook, Excel.Application xlApp, object misValue)
        {
            string new_url = (url.Substring(0, idx)) + "_copy.xlsm";
            xlWorkBook.SaveAs(new_url, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled, misValue,
            misValue, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
            Excel.XlSaveConflictResolution.xlUserResolution, true,
            misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            return new_url;
        }
        public static void deleteXMLsheetProtectionNodes(string folderPath)
        {
            foreach (string file in Directory.EnumerateFiles(folderPath, "*.xml"))
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(file);
                var nsmgr = new XmlNamespaceManager(doc.NameTable);// xmlns = "-http://schemas.openxmlformats.org/spreadsheetml/2006/main"//
                nsmgr.AddNamespace("rootNode", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");// The definition of the XML-a contains rootNode //xmlns="-http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                XmlElement el = (XmlElement)doc.SelectSingleNode("//rootNode:sheetProtection", nsmgr);
                if (el != null) { el.ParentNode.RemoveChild(el); }
                doc.Save(file);
            }
        }
        public static void deleteXMLworkbookProtectionNodes(string folderPath)
        {
            foreach (string file in Directory.EnumerateFiles(folderPath, "*.xml"))
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(file);
                var nsmgr = new XmlNamespaceManager(doc.NameTable);// xmlns = "-http://schemas.openxmlformats.org/spreadsheetml/2006/main"//
                nsmgr.AddNamespace("rootNode", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");// The definition of the XML-a contains rootNode//xmlns="-http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                XmlElement el = (XmlElement)doc.SelectSingleNode("//rootNode:workbookProtection", nsmgr);
                if (el != null) { el.ParentNode.RemoveChild(el); }
                doc.Save(file);
            }
        }
        public static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
