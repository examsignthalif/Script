using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Xps.Packaging;
using System.Windows.Xps.Serialization;
using ApplicationClass = Microsoft.Office.Interop.Word.ApplicationClass;

namespace Script
{
    public class WordProcessHelper
    {
        public List<XpsDocument> Convert_WordList_XpsList(List<string> WordFile_List)
        {
            List<XpsDocument> xpsDocList = new List<XpsDocument>();
            foreach (string filePath in WordFile_List)
            {
                //xpsDocList.Add(Convert_WordToXPS(filePath));
            }
            return xpsDocList;
        }

        public string Convert_WordToXPS(string word_Path)
        {
            string xps_Destination = word_Path.Replace(".docx", ".xps");
            ApplicationClass wordApp = new ApplicationClass();
            Microsoft.Office.Interop.Word.Document wordDoc = null;
            try
            {
                wordDoc = wordApp.Documents.Open(word_Path);
                wordDoc.SaveAs2(xps_Destination, WdSaveFormat.wdFormatXPS);
                wordDoc.Close();
                wordApp.Quit();

            }
            catch (Exception ex) { }
            return xps_Destination;
        }


        public List<FixedDocumentSequence> Convert_XpsToFixedDocumentSequence(IEnumerable<XpsDocument> xps_List)
        {
            List<FixedDocumentSequence> toReturn = new List<FixedDocumentSequence>();
            foreach (XpsDocument item in xps_List)
            {
                toReturn.Add(item.GetFixedDocumentSequence());
            }
            return toReturn;
        }
        public List<string> Get_WordDocxFrom(string folderPath)
        {
            return Directory.GetFiles(folderPath).ToList();
        }
        public byte[] Convert_XpsTOByteArray(XpsDocument xpsDocument)
        {
            byte[] toSave;
            using (MemoryStream ms = new MemoryStream())
            {
                var writer = new XpsSerializerFactory().CreateSerializerWriter(ms);
                writer.Write(xpsDocument.GetFixedDocumentSequence());
                toSave = ms.ToArray();
            }
            return toSave;
        }
        public FixedDocumentSequence ByteToFixedDocumentSequence(byte[] sourceXPS)
        {
            MemoryStream ms = new MemoryStream(sourceXPS);
            string memoryName = "memorystream://ms.xps";

            Uri memoryUri = new Uri(memoryName);
            try
            {
                PackageStore.RemovePackage(memoryUri);
            }
            catch (Exception) { }

            Package package = Package.Open(ms);
            PackageStore.AddPackage(memoryUri, package);

            XpsDocument xps = new XpsDocument(package, CompressionOption.SuperFast, memoryName);
            FixedDocumentSequence fixedDocumentSequence = xps.GetFixedDocumentSequence();
            return fixedDocumentSequence;
        }

        public XpsDocument ConvertPDFtoXPS(string path, string fileName)
        {
            string destination = new DirectoryInfo(Environment.CurrentDirectory).ToString() + "\\XPSfolder";
            if (!Directory.Exists(destination)) Directory.CreateDirectory(destination);

            string OutFilePath = destination + "\\" + fileName;

            //Spire.Pdf.PdfDocument myPDF = new Spire.Pdf.PdfDocument();
            //myPDF.LoadFromFile(path);
            //myPDF.SaveToFile(OutFilePath, Spire.Pdf.FileFormat.XPS);
            XpsDocument xpsDoc = new XpsDocument(OutFilePath, FileAccess.Read);
            return xpsDoc;
        }
    }
}
