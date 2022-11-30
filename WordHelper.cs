﻿using Microsoft.Office.Interop.Word;

using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordsChenger
{
    public class WordHelper
    {
        private FileInfo _fileInfo;

      public  WordHelper(string fileName)
        {
            if (File.Exists(fileName))
            {
                _fileInfo = new FileInfo(fileName);
            }
            else
            {
                throw new ArgumentException("file not found");
            }
        }

        internal bool Process(Dictionary <string,string>items)
        {
            Word.Application app = null;
            try
            {
                app = new Word.Application();
                Object file =_fileInfo.FullName;
                Object missing = Type.Missing;
                app.Documents.Open(file);
                foreach(var item in items)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text=item.Value;
                    Object wrap = Word.WdFindWrap.wdFindContinue;
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(FindText:
                        Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing, Replace: replace);
                }
                //Object newFilename = Path.Combine(_fileInfo.DirectoryName,
                // DateTime.Now.ToString("yyyyMMdd HHmmss") + _fileInfo.Name);
                Object newFilename = Path.Combine(_fileInfo.DirectoryName,"№"+
                    items["<first>"] + "_Испытания электродвигателя переменного тока напряжением до 1 кВ_" + 
                    "поз."+ items["<second>"]+".doc");
                app.ActiveDocument.SaveAs2(newFilename);
                app.ActiveDocument.Close();
                app.Quit();
                return true;
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (app!=null)
                {
                    
                    app.Quit();
                    MessageBox.Show("Файл создан!");
                }
            }
            return false;
        }
    }
}