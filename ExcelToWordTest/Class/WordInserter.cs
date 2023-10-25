using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//ВАЖНО в ссылках указать InteropWord (Object Library)
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows;

namespace ExcelToWordTest.Class
{
    internal class WordInserter
    {
        private FileInfo _fileInfo;

        public WordInserter(string filepath) 
        {
            if (File.Exists(filepath)) 
            {
                _fileInfo = new FileInfo(filepath);
            }
            else
            {
                throw new ArgumentException("Файл не найден");
            }
        }

        internal bool Process(Dictionary<string, string> items, string[,] students)
        {
            Word.Application app = null;
            try
            {
                app = new Word.Application();
                Object file = _fileInfo.FullName;
                Object missing = Type.Missing;

                app.Documents.Open(file);

                foreach (var item in items )
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;

                    Object wrap = Word.WdFindWrap.wdFindContinue;
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(
                        FindText: Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing, Replace: replace
                        );
                }

                for (int i = 0; i < students.GetLength(0); i++)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = "<Student" + (i+1).ToString() + "_FIO>";
                    find.Replacement.Text = students[i,0];
                    MessageBox.Show("<Student" + (i + 1).ToString() + "_FIO>" + " ЗАМЕНИТЬ " + students[i, 0]);

                    Object wrap = Word.WdFindWrap.wdFindContinue;
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(
                        FindText: Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing, Replace: replace
                        );
                }

                Object newFileName = Path.Combine(_fileInfo.DirectoryName, DateTime.Now.ToString("yy.MM.HH") + _fileInfo.Name);
                app.ActiveDocument.SaveAs2(newFileName);
                app.ActiveDocument.Close();
                return true;
            }
            catch (Exception ex){ MessageBox.Show(ex.Message); }
            finally 
            {
                if (app != null) { app.Quit(); } ;
            };
            return false;
        }
    }
}
