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
                                    String.Empty,
                                    String.Empty,
                                    "PowerMockito Test Procedure",
                                    "Please refer to 1.4 for the detailed steps of the test procedure",
                                    "Test project:",
                                    "Package path:",
                                    };



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
            bool bSyntaxCheckingBreak = false;


            string szTocString = String.Empty;

            Word.Paragraphs lParagraphs = null;

            try
            {

                object unit = Word.WdUnits.wdParagraph;
                a_wordDoc.TablesOfContents[1].IncludePageNumbers = false;
                a_wordDoc.TablesOfContents[1].HidePageNumbersInWeb = true;


                // find class Name in ToC
                foreach (Word.Hyperlink hl in a_wordDoc.TablesOfContents[1].Range.Hyperlinks)
                {

                    Word.Bookmark wb = a_wordDoc.Bookmarks[hl.SubAddress];

                    //Logger.Print($"1:{hl.Name}, 2:{hl.Range.Text} 4:{wb.Range.Text}", Logger.PrintOption.File);

                    string replaceString = hl.Range.Text.Replace('\t', ' ').Replace('\r', ' ');

                    //// get chapter
                    string chapter = replaceString.Substring(0, replaceString.IndexOf(' ')).Trim();
                    string header = replaceString.Substring(chapter.Length).Trim();


                    string wbString = new string(wb.Range.Text.Where(c => !char.IsControl(c)).ToArray()).Trim();



                    if (a_classText.Equals(header) && wbString.Equals(header))
                    {
                        bToCFound = true;
                        szTocString = $"{chapter} - {header}";
                        lParagraphs = wb.Range.Paragraphs;
                        break;
                    }


                }



                // return Error if the class cannot be found in ToC
                if (bToCFound == false)
                {
                    Logger.Print($"  - {a_classText} has not fould in the table of contents of SUTS", Logger.PrintOption.Both);

                    sPrevResult = Constants.StringTokens.ERROR;
                    sPrevFindString = a_classText;

                    return Constants.StringTokens.ERROR;
                }


                // get paragraph from above result.
                Word.Range range = lParagraphs.First.Range;

                // revmoe control character.
                string sz = new string(range.Text.Where(c => !char.IsControl(c)).ToArray());
                string title = sz.Trim();

 

                if (a_classText.Equals(title) &&
                    lParagraphs.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevel2)
                {

                    bSyntaxCheckingBreak = true;

                    // get next one of paragraph
                    Word.Range nextPara = range.Next(unit);

                    //Logger.Print($"({r.Paragraphs.Count}) - \"{r.Text}\"", Logger.PrintOption.File);

                    for (int index = 0; index < SUTS_Pattern_JAVA.Length; index++)
                    {

                        string input = nextPara.Text;

                        //Logger.Print($"({index,02}) - \"{input}\"", Logger.PrintOption.File);

                        //char[] charA = input.ToCharArray();

                        //StringBuilder sb = new StringBuilder();
                        //foreach (char ch in charA)
                        //{
                        //    //Logger.Print($"{Convert.ToByte(ch).ToString("X2")} \'{ch:c}\'", Logger.PrintOption.File);
                        //    sb.Append(Convert.ToByte(ch).ToString("X2") + " ");
                        //}
                        //Logger.Print($"{sb.ToString()}", "");
                        //Logger.Print("", "");


                        string output = new string(input.Where(c => !char.IsControl(c)).ToArray());
                        string s = output.Trim();


                        if (index == 1)
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
                                Logger.Print($" - There are some different words between Pattern and SUTS. ({index})", Logger.PrintOption.Both);
                                //Logger.Print($"   PATN: {docS}", Logger.PrintOption.File);
                                //Logger.Print($"   SUTS: {s}", Logger.PrintOption.File);

                                break;
                            }
                        }
                        else if (index == 4)
                        {
                            if (s.StartsWith("adb shell") || (s.StartsWith("N/A")))
                            {
                                bSUTSFound = true;

                            }
                            else
                            {
                                bSUTSFound = false;
                                Logger.Print($" - The content of Test script shall be N/A or starting with ADB shell. ({index})", Logger.PrintOption.Both);
                                //Logger.Print($"   SUTS: {s}", Logger.PrintOption.File);

                                break;
                            }
                        }
                        else if (index == 6)
                        {

                            if (String.IsNullOrEmpty(s))
                            {
                                bSUTSFound = true;
                            }
                            else
                            {
                                index++;

                                if (s.Contains(SUTS_Pattern_JAVA[index]))
                                {
                                    bSUTSFound = true;
                                }
                                else
                                {
                                    bSUTSFound = false;
                                    Logger.Print($" - The format of Section {a_classText} shall be checked. ({index})", Logger.PrintOption.Both);
                                    //Logger.Print($"   PATN: {SUTS_Pattern_JAVA[index]}", Logger.PrintOption.File);
                                    //Logger.Print($"   SUTS: {s}", Logger.PrintOption.File);
                                    break;
                                }
                            }

                        }
                        else
                        {

                            if (s.Contains(SUTS_Pattern_JAVA[index]))
                            {
                                bSUTSFound = true;

                                // the final content of section
                                if (s.StartsWith(SUTS_Pattern_JAVA[SUTS_Pattern_JAVA.Length - 1]))
                                {
                                    break;
                                }

                            }
                            else
                            {
                                bSUTSFound = false;
                                Logger.Print($" - The format of Section {a_classText} shall be check. ({index})", Logger.PrintOption.Both);
                                //Logger.Print($"   PATN: {SUTS_Pattern_JAVA[index]}", Logger.PrintOption.File);
                                //Logger.Print($"   SUTS: {s}", Logger.PrintOption.File);
                                break;
                            }

                        }

                        // get next
                        nextPara = nextPara.Next(unit);

                    }


                }
                else
                {
                    bSUTSFound = false;

                }

            }
            catch (Exception ex)
            {
                bSUTSFound = false;
                Logger.Print(sFuncName, ex.Message, Logger.PrintOption.File);
                return Constants.StringTokens.ERROR;
            }


            if (false == bSUTSFound && false == bSyntaxCheckingBreak)
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

            string szToCString = String.Empty;
            bool bToCFound = false;
            Word.Paragraphs lParagraphs = null;

            try
            {

                a_wordDoc.TablesOfContents[1].IncludePageNumbers = false;
                a_wordDoc.TablesOfContents[1].HidePageNumbersInWeb = true;

                foreach (Word.Hyperlink hl in a_wordDoc.TablesOfContents[1].Range.Hyperlinks)
                {

                    Word.Bookmark wb = a_wordDoc.Bookmarks[hl.SubAddress];

                    //Logger.Print($"1:{hl.Name}, 2:{hl.Range.Text} 4:{wb.Range.Text}", Logger.PrintOption.File);

                    string replaceString = hl.Range.Text.Replace('\t', ' ').Replace('\r', ' ');

                    //// get chapter
                    string chapter = replaceString.Substring(0, replaceString.IndexOf(' ')).Trim();
                    string header = replaceString.Substring(chapter.Length).Trim();


                    string wbString = new string(wb.Range.Text.Where(c => !char.IsControl(c)).ToArray()).Trim();



                    if (a_methodName.Equals(header) && wbString.Equals(header))
                    {
                        bToCFound = true;
                        szToCString = $"{chapter} - {header}";
                        lParagraphs = wb.Range.Paragraphs;

                        break;
                    }


                }


                if (bToCFound == false)
                {
                    Logger.Print($"  - {a_methodName} has not fould in the table of contents of SUTS", Logger.PrintOption.File);

                    sPrevResult = Constants.StringTokens.ERROR;
                    sPrevFindString = testCaseName;

                    return Constants.StringTokens.ERROR;
                }



                // for Find.Execute
                object findText = testCaseName;


                Word.Range rng = lParagraphs.First.Range;
                rng.Find.ClearFormatting();

                rng.SetRange(rng.Start, a_wordDoc.Application.ActiveDocument.Content.End);
                rng.Select();


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
        /// Get the full path of SUTS
        /// </summary>
        /// <param name="a_name">The name of module</param>
        /// <param name="a_path">The path of SUTS</param>
        public string SearchSUTSDocumentPath(string a_name, string a_path)
        {
            string sFuncName = "[SearchSUTSDocumentPath]";

            string suts_name = Constants.SUTS_FILENAME_PREFIX + a_name.Replace('_', ' ') + ".doc";
            string path = string.Empty;

            // Check the existence of the specified path.
            if (!Directory.Exists(a_path))
            {
                Logger.Print(sFuncName, "Cannot find path \"" + a_path + "\"; skipped.", Logger.PrintOption.Both);
                return string.Empty;
            }

            path = Path.GetFullPath(a_path + "\\" + suts_name);

            if (File.Exists(path))
            {
                Logger.Print(sFuncName, $"SUTS was found in {suts_name}");
                
            }
            else
            {
                path = String.Empty;

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







        private static string _gSUTSChapter = string.Empty;
        private static Object sutsLock = new object();


        public static string gSUTSChapter
        {

            get
            {
                lock (sutsLock)
                {
                    return _gSUTSChapter;
                }
            }

            set
            {
                lock (sutsLock)
                {
                    _gSUTSChapter = value;
                }
            }

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_wordDoc"></param>
        /// <param name="a_classText"></param>
        /// <param name="a_TDSExcelName"></param>
        /// <returns></returns>
        public void SUTS_FindSectionOfClass_JavaByThread(object obj)
        {
            string sFuncName = "[FindSectionOfClassInSUTS_Java]";

            Tuple<Word.Document, string, string> items = obj as Tuple<Word.Document, string, string>;

            Word.Document a_wordDoc = items.Item1 as Word.Document;
            string a_classText = items.Item2 as string;
            string a_TDSExcelName = items.Item3 as string;


            // if current text is same as previous, then just return the previous result.
            if (a_classText.Equals(sPrevFindString))
            {
                gSUTSChapter = sPrevResult;
                return;
            }

            // Check Word.Application
            if (null == g_wordApp)
            {
                Logger.Print(sFuncName, ErrorMessage.WORD_APP_IS_NULL);
                gSUTSChapter = Constants.StringTokens.ERROR;
                return;
            }

            // cHECK Word.Document
            if (null == a_wordDoc)
            {
                Logger.Print(sFuncName, ErrorMessage.SUTS_DOC_IS_NULL);
                gSUTSChapter = Constants.StringTokens.ERROR;
                return;
            }


            // flag for ToC
            bool bToCFound = false;

            // flag for SUTS
            bool bSUTSFound = false;

            // flag that indicate the suts has beed found is used for a break from searching loop.
            bool bSyntaxCheckingBreak = false;


            string szTocString = String.Empty;

            Word.Paragraphs lParagraphs = null;

            try
            {

                object unit = Word.WdUnits.wdParagraph;
                a_wordDoc.TablesOfContents[1].IncludePageNumbers = false;
                a_wordDoc.TablesOfContents[1].HidePageNumbersInWeb = true;


                // find class Name in ToC
                foreach (Word.Hyperlink hl in a_wordDoc.TablesOfContents[1].Range.Hyperlinks)
                {

                    Word.Bookmark wb = a_wordDoc.Bookmarks[hl.SubAddress];

                    //Logger.Print($"1:{hl.Name}, 2:{hl.Range.Text} 4:{wb.Range.Text}", Logger.PrintOption.File);

                    string replaceString = hl.Range.Text.Replace('\t', ' ').Replace('\r', ' ');

                    //// get chapter
                    string chapter = replaceString.Substring(0, replaceString.IndexOf(' ')).Trim();
                    string header = replaceString.Substring(chapter.Length).Trim();


                    string wbString = new string(wb.Range.Text.Where(c => !char.IsControl(c)).ToArray()).Trim();



                    if (a_classText.Equals(header) && wbString.Equals(header))
                    {
                        bToCFound = true;
                        szTocString = $"{chapter} - {header}";
                        lParagraphs = wb.Range.Paragraphs;
                        break;
                    }


                }



                // return Error if the class cannot be found in ToC
                if (bToCFound == false)
                {
                    Logger.Print($"  - {a_classText} has not fould in the table of contents of SUTS", Logger.PrintOption.Both);

                    sPrevResult = Constants.StringTokens.ERROR;
                    sPrevFindString = a_classText;
                    gSUTSChapter = Constants.StringTokens.ERROR;
                    return;
                }



                // get paragraph from above result.
                Word.Range range = lParagraphs.First.Range;

                // revmoe control character.
                string sz = new string(range.Text.Where(c => !char.IsControl(c)).ToArray());
                string title = sz.Trim();


                if (a_classText.Equals(title) &&
                    lParagraphs.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevel2)
                {

                    bSyntaxCheckingBreak = true;

                    // get next one of paragraph
                    Word.Range nextPara = range.Next(unit);

                    //Logger.Print($"({r.Paragraphs.Count}) - \"{r.Text}\"", Logger.PrintOption.File);

                    for (int index = 0; index < SUTS_Pattern_JAVA.Length; index++)
                    {

                        string input = nextPara.Text;

                        string output = new string(input.Where(c => !char.IsControl(c)).ToArray());
                        string s = output.Trim();


                        if (index == 1)
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
                                Logger.Print($" - There are some different words between Pattern and SUTS. ({index})", Logger.PrintOption.Both);
                                //Logger.Print($"   PATN: {docS}", Logger.PrintOption.File);
                                //Logger.Print($"   SUTS: {s}", Logger.PrintOption.File);

                                break;
                            }
                        }
                        else if (index == 4)
                        {
                            if (s.StartsWith("adb shell") || (s.StartsWith("N/A")))
                            {
                                bSUTSFound = true;

                            }
                            else
                            {
                                bSUTSFound = false;
                                Logger.Print($" - The content of Test script shall be N/A or starting with ADB shell. ({index})", Logger.PrintOption.Both);
                                //Logger.Print($"   SUTS: {s}", Logger.PrintOption.File);

                                break;
                            }
                        }
                        else if (index == 6)
                        {

                            if (String.IsNullOrEmpty(s))
                            {
                                bSUTSFound = true;
                            }
                            else
                            {
                                index++;

                                if (s.Contains(SUTS_Pattern_JAVA[index]))
                                {
                                    bSUTSFound = true;
                                }
                                else
                                {
                                    bSUTSFound = false;
                                    Logger.Print($" - The format of Section {a_classText} shall be checked. ({index})", Logger.PrintOption.Both);
                                    //Logger.Print($"   PATN: {SUTS_Pattern_JAVA[index]}", Logger.PrintOption.File);
                                    //Logger.Print($"   SUTS: {s}", Logger.PrintOption.File);
                                    break;
                                }
                            }

                        }
                        else
                        {

                            if (s.Contains(SUTS_Pattern_JAVA[index]))
                            {
                                bSUTSFound = true;

                                // the final content of section
                                if (s.StartsWith(SUTS_Pattern_JAVA[SUTS_Pattern_JAVA.Length - 1]))
                                {
                                    break;
                                }

                            }
                            else
                            {
                                bSUTSFound = false;
                                Logger.Print($" - The format of Section {a_classText} shall be check. ({index})", Logger.PrintOption.Both);
                                //Logger.Print($"   PATN: {SUTS_Pattern_JAVA[index]}", Logger.PrintOption.File);
                                //Logger.Print($"   SUTS: {s}", Logger.PrintOption.File);
                                break;
                            }

                        }

                        // get next
                        nextPara = nextPara.Next(unit);

                    }

                }
                else
                {
                    bSUTSFound = false;

                }



            }
            catch (Exception ex)
            {
                bSUTSFound = false;
                Logger.Print(sFuncName, ex.Message, Logger.PrintOption.File);
                gSUTSChapter = Constants.StringTokens.ERROR;
                return;

            }


            if (false == bSUTSFound && false == bSyntaxCheckingBreak)
            {
                Logger.Print($"  - The content of section {a_classText} have not found.", Logger.PrintOption.Both);
                szTocString = Constants.StringTokens.ERROR;
            }

            sPrevResult = szTocString;
            sPrevFindString = a_classText;
            gSUTSChapter = szTocString;

            return;

        }


    }
}
