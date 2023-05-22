using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace LABA_26_OOP
{
    class WordHelper
    {
        private FileInfo _fileInfo;

        public WordHelper(string fileName)
        {
            if (File.Exists(fileName))
            {
                _fileInfo = new FileInfo(fileName);
            }
            else
            {
                throw new ArgumentException("Файл не знадено!");
            }
        }

        internal void Process(Dictionary<string, string> items)
        {
            Word.Application app = null;
            try
            {
                app = new Word.Application();
                Object file = _fileInfo.FullName;
                Object missing = Type.Missing;
                app.Documents.Open(file);
                foreach (var item in items)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;
                    Object wrap = Word.WdFindWrap.wdFindContinue;
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                        MatchCase: false,
                        Wrap: wrap,
                        ReplaceWith: missing, Replace: replace);
                }

                Object newFileName = Path.Combine(_fileInfo.DirectoryName, DateTime.Now.ToString("yyMMdd HHmm ") + _fileInfo.Name);
                app.ActiveDocument.SaveAs2(newFileName);
                app.ActiveDocument.Close();
                app.Quit();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                }
            }

        }
        public void Open()
        {
            var app = new Word.Application();
            Object file = _fileInfo.FullName;
            app.Documents.Open(file);
            app.Visible = true;
        }
    }
}
