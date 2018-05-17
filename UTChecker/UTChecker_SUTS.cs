using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace UTChecker
{
    public partial class UTChecker
    {

        static string sPrevFindString = "";
        static string sPrevResult = "";
       

        /// <summary>
        /// Clear all preivous Result
        /// </summary>
        public void SUTS_ClearPreviousResult()
        {
            sPrevFindString = "";
            sPrevResult = "";
        }

        string[] SUTS_Pattern_JAVA = {
                                    "Test Data Sheet",
                                    "Test cases and test information are detailed in ",
                                    "Default Test Procedure",
                                    "Execute the following commands in command line to run all test cases to test the target function:",
                                    "adb shell",
                                    "N/A",
                                    "PowerMockito Test Procedure",
                                    "Please refer to 1.4 for the detailed steps of the test procedure",
                                    "Test project:",
                                    "Package path:",
                                    };

        



        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_wordDoc"></param>
        /// <param name="a_classText"></param>
        /// <param name="a_TDSExcelName"></param>
        /// <returns></returns>
        public string SUTS_FindSectionOfClass_Java(Word.Document a_wordDoc, string a_classText, string a_TDSExcelName)
        {
            string sFuncName = "[FindSectionOfClassInSUTS_Java]";


            // if current text is same as previous, then just return the previous result.
            if (a_classText.Equals(sPrevFindString))
            {
                return sPrevResult;
            }

            // Check Word.Application
            if (null == g_wordApp)
            {
                Logger.Print(sFuncName, ErrorMessage.WORD_APP_IS_NULL);
                return Constants.StringTokens.ERROR;
            }

            // cHECK Word.Document
            if (null == a_wordDoc)
            {
                Logger.Print(sFuncName, ErrorMessage.SUTS_DOC_IS_NULL);
                return Constants.StringTokens.ERROR;
            }


            // flag for ToC
            bool bToCFound = false;

            // flag for SUTS
            bool bSUTSFound = false;

            // flag that indicate the suts has beed found is used for a break from searching loop.
            bool bForceBreak = false;


            string szTocString = String.Empty;


            try
            {


                object unit = Word.WdUnits.wdSentence;

                int tS = a_wordDoc.Application.ActiveDocument.Content.Start;
                int tE = a_wordDoc.Application.ActiveDocument.Content.End;
                //Logger.Print($"Range from {tS} to {tE}", Logger.PrintOption.Both);

                int loopStart = a_wordDoc.TablesOfContents[1].Range.End;
                int loopEnd = a_wordDoc.Application.ActiveDocument.Content.End;

                // find class Name in ToC
                foreach (Word.Paragraph p in a_wordDoc.TablesOfContents[1].Range.Paragraphs)
                {

                    char[] charSeparators = new char[] { ' ', '\r', '\t' };
                    string[] words = p.Range.Text.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);

                    //Logger.Print($"{p.Range.Start,05} - {p.Range.End,05} - {p.Range.Text}", Logger.PrintOption.Both);


                    if (a_classText.EndsWith(words[1]))
                    {
                        bToCFound = true;
                        szTocString = $"{words[0]} - {words[1]}";
                        break;
                    }

                }

                if (bToCFound == false)
                {
                    Logger.Print($"  - {a_classText} has not fould in the table of contents of SUTS", Logger.PrintOption.File);
                    return Constants.StringTokens.ERROR;
                }



                foreach (Word.Paragraph paragraph in a_wordDoc.Application.ActiveDocument.Paragraphs)
                {
                    int diff = 0;

                    Word.Range r = paragraph.Range;


                    string title = r.Text.Trim(new char[] { '\r', '\n', '\u0015', ' ', '\t' });

                    if (a_classText.EndsWith(title) &&
                        r.Paragraphs[1].OutlineLevel == Word.WdOutlineLevel.wdOutlineLevel2)
                    {

                        bForceBreak = true;

                        // get next one of paragraph
                        Word.Range n = r.Next(unit);


                        for (int i = 0; i < 10; i++)
                        {
                            string s = n.Text.Trim(new char[] { '\r', '\n', '\x15', '\x20', ' ' });

                            int pTag = n.Paragraphs.Count;

                            if (i == 1)
                            {
                                string _TDSExcelName = a_TDSExcelName;

                                // get file name 
                                int ind = _TDSExcelName.LastIndexOf('\\');
                                if (ind > 0)
                                {
                                    _TDSExcelName = _TDSExcelName.Remove(0, ind + 1);
                                }

                                string docS = SUTS_Pattern_JAVA[1] + _TDSExcelName;

                                if (docS.Equals(s))
                                {
                                    bSUTSFound = true;
                                }
                                else
                                {
                                    bSUTSFound = false;
                                    Logger.Print("  - There are some different words between Pattern and SUTS. ", Logger.PrintOption.Both);
                                    Logger.Print("    PATN: " + docS, Logger.PrintOption.Both);
                                    Logger.Print("    SUTS: " + s, Logger.PrintOption.Both);

                                    break;
                                }
                            }
                            else if (i == 4)
                            {
                                if (s.StartsWith("adb shell") || (s.StartsWith("N/A")))
                                {
                                    bSUTSFound = true;

                                }
                                else
                                {
                                    bSUTSFound = false;
                                    Logger.Print("  - The content of Test script shall be N/A or starting with ADB shell.", Logger.PrintOption.Both);
                                    break;
                                }
                            }
                            else if (i == 5 && diff == 0)
                            {
                                if (s.Equals(string.Empty))
                                {
                                    bSUTSFound = true;

                                    if (pTag >= 2)
                                    {
                                        Logger.Print("  - The tag of paragraph shall have one only.", Logger.PrintOption.Both);
                                        break;
                                    }
                                }
                                else
                                {
                                    diff = 1;

                                    if (s.Equals(SUTS_Pattern_JAVA[i + diff]))
                                    {
                                        bSUTSFound = true;
                                    }
                                    else
                                    {
                                        Logger.Print($"  - The contents of section {a_classText} shall be double confirmed.", Logger.PrintOption.Both);
                                        bSUTSFound = false;
                                        break;
                                    }
                                }
                            }
                            else
                            {

                                if (s.Contains(SUTS_Pattern_JAVA[i + diff]))
                                {
                                    bSUTSFound = true;

                                    // the final content of section
                                    if (s.Contains(SUTS_Pattern_JAVA[9]))
                                    {
                                        break;
                                    }

                                }
                                else
                                {
                                    bSUTSFound = false;
                                    Logger.Print($"  - The contents of section {a_classText} shall be double confirmed.", Logger.PrintOption.Both);
                                    break;
                                }

                            }

                            // get next
                            n = n.Next(unit);

                        }

                        // forec break from foreach
                        if (bForceBreak)
                        {
                            break;
                        }
                    }
                    else
                    {
                        bSUTSFound = false;

                    }



                }


            }
            catch (Exception ex)
            {
                bSUTSFound = false;
                Logger.Print(sFuncName, ex.Message, Logger.PrintOption.File);
                return Constants.StringTokens.ERROR;
            }


            if (false == bSUTSFound && false == bForceBreak)
            {
                Logger.Print($"  - The content of section {a_classText} have not found.", Logger.PrintOption.Both);
                szTocString = Constants.StringTokens.ERROR;
            }

            sPrevResult = szTocString;
            sPrevFindString = a_classText;


            return szTocString;

        }





        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_wordDoc"></param>
        /// <param name="findString"></param>
        /// <returns></returns>
        public string SUTS_FindSectionOfClass_C(Word.Document a_wordDoc, string a_methodName, string a_testCase)
        {
            string sFuncName = "[SUTS_FindSectionOfClass_C]";

           

            string testCaseName = $"{a_methodName}.{a_testCase}";


            if (testCaseName.Equals(sPrevFindString))
            {
                return sPrevResult;
            }

            // Check Word.Application
            if (null == g_wordApp)
            {
                Logger.Print(sFuncName, ErrorMessage.WORD_APP_IS_NULL);
                return Constants.StringTokens.ERROR;
            }

            // Check Word.Document
            if (null == a_wordDoc)
            {
                Logger.Print(sFuncName, ErrorMessage.SUTS_DOC_IS_NULL);
                return Constants.StringTokens.ERROR;
            }

            string szToCString = "";
            bool bToCFound = false;

            try
            {

                // find class Name in ToC
                foreach (Word.Paragraph p in a_wordDoc.TablesOfContents[1].Range.Paragraphs)
                {
                    if (p.Range.Text.Contains(a_methodName))
                    {
                        szToCString = p.Range.Text;
                        bToCFound = true;
                        break;
                    }
                }

                if (bToCFound == false)
                {
                    Logger.Print($"  - {a_methodName} has not fould in the table of contents of SUTS", Logger.PrintOption.File);
                    return Constants.StringTokens.ERROR;
                }

                // for Find.Execute
                object findText = testCaseName;


                Word.Range rng = a_wordDoc.Application.ActiveDocument.Content;
                rng.Find.ClearFormatting();

                if (false == rng.Find.Execute(
                    ref findText,
                    true, 
                    Type.Missing, Type.Missing, Type.Missing, 
                    Type.Missing, Type.Missing, Type.Missing, 
                    Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing))
                {
                    Logger.Print($"  - {findText} have not found in STUS.", Logger.PrintOption.File);
                    szToCString = Constants.StringTokens.ERROR;
                }


            }
            catch (Exception e)
            {
                Logger.Print(e.Message, Logger.PrintOption.File);
                szToCString = Constants.StringTokens.ERROR;
            }

            
            sPrevResult = szToCString;
            sPrevFindString = testCaseName;

            return szToCString;
        }




        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_item"></param>
        /// <param name="a_SUTSPath"></param>
        /// <returns></returns>
        public string SearchSUTSDocumentPath(string a_item, string a_SUTSPath)
        {
            string sFuncName = "[SearchSUTSDocumentPath]";
            string suts_name = Constants.SUTS_FILENAME_PREFIX + a_item.Replace('_', ' ') + ".doc";
            string path = "";

            Logger.Print(sFuncName, $"Search SUTS {suts_name}");


            // Check the existence of the specified path.
            if (!Directory.Exists(a_SUTSPath))
            {
                Logger.Print(sFuncName, "Cannot find path \"" + a_SUTSPath + "\"; skipped.", Logger.PrintOption.Both);
                return "";
            }

            // Collect the considered files stored in current folder.
            string[] FileList = Directory.GetFiles(a_SUTSPath, Constants.SUTS_FILENAME_EXT);
            foreach (string f in FileList)
            {
                string docName = Path.GetFileName(f);

                if (docName.Equals(suts_name))
                {
                    path = f;

                    Logger.Print(sFuncName, $"SUTS is found in {path}");
                    Logger.Print(sFuncName, $"SUTS found", Logger.PrintOption.Logger);
                    break;
                }
            }

            return path;
        }






        public void testWord(string wordPath)
        {
            string sFuncName = "testWord";


            Word.Document wordDoc = null;

            try
            {
                wordDoc = OpenWordDocument(g_wordApp, wordPath);

                if (null == wordDoc)
                {
                    Logger.Print(sFuncName, $"Cannot found SUTS.");
                }

                //Word.Range rng = wordDoc.Application.ActiveDocument.Content;

                Logger.Print("Paragraphs " + wordDoc.Application.ActiveDocument.Paragraphs.Count.ToString(), Logger.PrintOption.File);
                Logger.Print("Sentences " + wordDoc.Application.ActiveDocument.Sentences.Count.ToString(), Logger.PrintOption.File);
                Logger.Print("Bookmarks " + wordDoc.Application.ActiveDocument.Bookmarks.Count.ToString(), Logger.PrintOption.File);
                Logger.Print("Sections " + wordDoc.Application.ActiveDocument.Sections.Count.ToString(), Logger.PrintOption.File);



                object unit = Word.WdUnits.wdSentence;
                object count = 1;

                foreach (Word.Paragraph r in wordDoc.Application.ActiveDocument.Content.Paragraphs)
                {

                }



            }
            catch (Exception ex)
            {
                Logger.Print(ex.Message);
            }
            wordDoc.Close();
        }



        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="a_wordDoc"></param>
        ///// <param name="findString"></param>
        ///// <returns></returns>
        //public string SUTS_FindSectionOfClass_C(Word.Document a_wordDoc, string findString)
        //{
        //    string sFuncName = "[FindSectionOfClassInSUTS_C]";

        //    string sResult = Constants.StringTokens.ERROR;

        //    if (findString.Equals(sPrevFindString))
        //    {
        //        return sPrevResult;
        //    }

        //    // Check the EXCEL app.
        //    if (null == g_wordApp)
        //    {
        //        Logger.Print(sFuncName, ErrorMessage.WORD_APP_IS_NULL);
        //        return Constants.StringTokens.ERROR;
        //    }

        //    if (null == a_wordDoc)
        //    {
        //        Logger.Print(sFuncName, ErrorMessage.SUTS_DOC_IS_NULL);
        //        return Constants.StringTokens.ERROR;
        //    }


        //    try
        //    {
        //        object findText = findString;

        //        Word.Range rng = a_wordDoc.Application.ActiveDocument.Content;
        //        rng.Find.ClearFormatting();

        //        if (rng.Find.Execute(ref findText,
        //            true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //            Type.Missing, Type.Missing))
        //        {

        //            sResult = "Found";
        //        }
        //        else
        //        {
        //            Logger.Print($"  - {findText} have not found in STUS.", Logger.PrintOption.File);
        //        }


        //    }
        //    catch (Exception e)
        //    {
        //        Logger.Print(e.Message, Logger.PrintOption.File);
        //    }


        //    sPrevResult = sResult;
        //    sPrevFindString = findString;

        //    return sResult;
        //}


        public void testWord2(string wordPath, string a_classText)
        {
            string sFuncName = "testWord2";
            Word.Document a_wordDoc = null;


            if (null == a_wordDoc)
            {
                Logger.Print(sFuncName, $"Cannot found SUTS.");
            }


            try
            {
                a_wordDoc = OpenWordDocument(g_wordApp, wordPath);


                object unit = Word.WdUnits.wdSentence;
                object count = 1;

                //Word.Range rng = wordDoc.Application.ActiveDocument.Content;

                Logger.Print("Paragraphs " + a_wordDoc.Application.ActiveDocument.Paragraphs.Count.ToString(), Logger.PrintOption.Both);
                Logger.Print("Sentences " + a_wordDoc.Application.ActiveDocument.Sentences.Count.ToString(), Logger.PrintOption.Both);
                Logger.Print("Bookmarks " + a_wordDoc.Application.ActiveDocument.Bookmarks.Count.ToString(), Logger.PrintOption.Both);
                Logger.Print("Sections " + a_wordDoc.Application.ActiveDocument.Sections.Count.ToString(), Logger.PrintOption.Both);
                Logger.Print("Sections " + a_wordDoc.Paragraphs.Count.ToString(), Logger.PrintOption.File);
                Logger.Print("TsOCs " + a_wordDoc.TablesOfContents.Count.ToString(), Logger.PrintOption.Both);


                int ParagraphsCount = a_wordDoc.Application.ActiveDocument.Paragraphs.Count;

                int tocStart = a_wordDoc.TablesOfContents[1].Range.Start;
                int tocEnd = a_wordDoc.TablesOfContents[1].Range.End;
                Logger.Print($"ToC from {tocStart} to {tocEnd}", Logger.PrintOption.File);


                //// TOC
                //foreach (Word.Paragraph p in a_wordDoc.TablesOfContents[1].Range.Paragraphs)
                //{

                //    Logger.Print($"{p.Range.Start,05} - {p.Range.End,05} - {p.Range.Text}", Logger.PrintOption.File);
                //}


                bool bToCFound = false;
                string szTocString = String.Empty;
                int startChapter = -1;

                for (int i = 1; i<= ParagraphsCount; i++)
                {
                    Word.Paragraph paragraph = a_wordDoc.Application.ActiveDocument.Paragraphs[i];
                    Word.Range r = paragraph.Range;
                    Logger.Print($"{i,04} - {r.Start,05} - {r.End,05} - {r.Text}", Logger.PrintOption.File);


                    if (r.Start >= tocStart && r.End <= tocEnd)
                    {

                        char[] charSeparators = new char[] { ' ', '\r', '\t' };
                        string[] words = r.Text.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);


                        if (a_classText.EndsWith(words[1]))
                        {
                            bToCFound = true;
                            szTocString = $"{words[0]} - {words[1]}";

                            Logger.Print($"Found {szTocString} in ToC", Logger.PrintOption.Both);
                            startChapter = i;
                            break;
                        }

                    }
                }

                for (int j = startChapter; j <= ParagraphsCount; j++)
                {

                    int diff = 0;
                    Word.Paragraph paragraph = a_wordDoc.Application.ActiveDocument.Paragraphs[j];
                    Word.Range r = paragraph.Range;


                    string title = r.Text.Trim(new char[] { '\r', '\n', '\u0015', ' ', '\t' });

                    if (a_classText.EndsWith(title) &&
                        r.Paragraphs[1].OutlineLevel == Word.WdOutlineLevel.wdOutlineLevel2)
                    {

                        Logger.Print($"Found {szTocString} in SUTS", Logger.PrintOption.Both);
                    }

                }

            }
            catch (Exception ex)
            {
                Logger.Print(ex.Message, Logger.PrintOption.File);
            }


            a_wordDoc.Close();

            ReleaseOfficeApps();

        }












    }
}
