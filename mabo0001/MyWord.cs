using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace mabo0001
{
    internal class MyWord
    {
        public string fileName;
        public Word.Application cls = null;
        public Word.Document doc = null;
        public Microsoft.Office.Interop.Word.Table table = null;
        public object missing = System.Reflection.Missing.Value;
 
        private bool openState;


        public MyWord(string fileName)
        {
            this.fileName = fileName;
        }

        public void Open()
        {
            try
            {
                killWinWordProcess();
                cls = new Word.Application();
                cls.Visible = true;
                object templateName = this.fileName;
                doc = cls.Documents.Open(ref templateName, ref missing, ref missing, ref missing,
                                         ref missing, ref missing, ref missing, ref missing,
                                         ref missing, ref missing, ref missing, ref missing,
                                         ref missing, ref missing, ref missing, ref missing);
                
                openState = true;
            }
            catch
            {
                openState = false;
            }
        }


        public void WriteTable(int rowIndex, int colIndex,int tableIndex)
        {
            table = doc.Tables[tableIndex];
            string text = table.Cell(rowIndex, colIndex).Range.Text.ToString();
            text = text.Substring(0, text.Length - 2);
            Console.WriteLine(text);
            Console.WriteLine(doc.Tables.Count.ToString());
            Console.WriteLine(1);
        }

        // 保存新文件
        public void SaveDocument(string filePath)
        {
            object fileName = filePath;
            object format = Word.WdSaveFormat.wdFormatDocument;//保存格式
            object miss = System.Reflection.Missing.Value;
            doc.SaveAs(ref fileName, ref format, ref miss,
                ref miss, ref miss, ref miss, ref miss,
                ref miss, ref miss, ref miss, ref miss,
                ref miss, ref miss, ref miss, ref miss,
                ref miss);
            //关闭wordDoc，wordApp对象
            object SaveChanges = Word.WdSaveOptions.wdSaveChanges;
            object OriginalFormat = Word.WdOriginalFormat.wdOriginalDocumentFormat;
            object RouteDocument = false;
            doc.Close(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
            cls.Quit(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
        }


        // 在书签处插入值
        public bool InsertValue(string bookmark, string value)
        {
            object bkObj = bookmark;
            if (cls.ActiveDocument.Bookmarks.Exists(bookmark))
            {
                cls.ActiveDocument.Bookmarks.get_Item(ref bkObj).Select();
                cls.Selection.TypeText(value);
                return true;
            }
            return false;
        }



        // 杀掉winword.exe进程
        public void killWinWordProcess()
        {
            System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("WINWORD");
            foreach (System.Diagnostics.Process process in processes)
            {
                bool b = process.MainWindowTitle == "";
                if (process.MainWindowTitle == "")
                {
                    process.Kill();
                }
            }
        }

    }
}