using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Xps.Packaging;

namespace Script
{
    public class FileItem
    {
        public int Qno { get; set; }
        public string FilePath { get; set; }
        public int Marks { get; set; }
        public byte[] XpsByteData { get; set; }
        public XpsDocument XPS { get; set; }
        public FileItem(int qno, string filePath)
        {
            this.Qno = qno;
            this.FilePath = filePath;
        }
        public FileItem()
        {

        }
    }
}
