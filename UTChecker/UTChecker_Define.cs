using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace UTChecker
{

    /// <summary>
    /// 
    /// </summary>
    public struct EnvrionmentSetting
    {
        /// <summary>
        /// The path of list file
        /// </summary>
        public string listFile;

        /// <summary>
        /// The path of TDS
        /// </summary>
        public string tdsPath;

        /// <summary>
        /// The path for outcome
        /// </summary>
        public string outputPath;

        /// <summary>
        /// The path of report template (include file name)
        /// </summary>
        public string reportTemplate;

        /// <summary>
        /// The path of summary template (include file name)
        /// </summary>
        public string summaryTemplate;

        /// <summary>
        /// The path of test log
        /// </summary>
        public string testlogPath;

        /// <summary>
        /// The path of SUTS
        /// </summary>
        public string sutsPath;

        /// <summary>
        /// reserved
        /// </summary>
        public string referenceListsPath;

        /// <summary>
        ///  reserved
        /// </summary>
        public string sutrrPath;

    }


    /// <summary>
    /// 
    /// </summary>
    public enum RunMode
    {
        // run TDS_Parser by CommandLine
        CommandLine,

        // run TDS_Parser by Window/User
        User,

        // None
        None,
    }


    public partial class UTChecker
    {
        public static class Constants
        {

            /// <summary>
            /// Prefix of the SUTS.
            /// </summary>
            public const string SUTS_FILENAME_PREFIX = "NUGEN Software Unit Test Specification Document of ";

            /// <summary>
            /// Ext file name of SUTS
            /// </summary>
            public const string SUTS_FILENAME_EXT = "*.doc";


            /// <summary>
            /// Prefix of the TDS.
            /// </summary>
            public const string TDS_FILENAME_PREFIX = "NUGEN Test Data Sheet - ";

            /// <summary>
            /// Ext file name of TDS
            /// </summary>
            public const string TDS_FILENAME_EXT = "*.xlsx";

            /// <summary>
            /// Ext file name for test log.
            /// </summary>
            public const string TESTLOG_FILENAME_EXT = "*.txt";

            /// <summary>
            /// Prefix file name for Report.
            /// </summary>
            public const string REPORT_PREFIX = "UT_CHECK_";

            /// <summary>
            /// The sheet name of Lookup table in TDS
            /// </summary>
            public const string SHEET_NAME = "LookupTable";

            /// <summary>
            /// /The sheet name of Summary in TDS
            /// </summary>
            public const string SHEET_SUMMARY = "Summary";


            /// <summary>
            /// The numbers of Argument (Command Line)
            /// </summary>
            public static class CommandArguments
            {
                public const int Minium = 1;   // UTChecker self.


                // list, tds path, output path, template, summary, test log path
                public const int Args = 7;
                public const int Match = Args  + Minium;
            }

            /// <summary>
            /// String Tokens
            /// </summary>
            public static class StringTokens
            {
                public const string DESIGN_ID_PREFIX = "NUSWDD";
                public const string NA = "N/A";
                public const string ERROR = "Error";
                public const string ERROR_MSG_HEADER = "Error: ";
                public const string MSG_BULLET = "  *";
                public const string MSG_SUB_BULLET = "    -";
                public const string DUPLICATE_TC_LABEL = "TC label is repeated.";
                public const string X = "X";
                public const string DEFAULT_INVALID_VALUE = ERROR_MSG_HEADER + "Unknown";
            }

            /// <summary>
            /// Color
            /// </summary>
            public class Color
            {
                static public System.Int32 RED = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            }

        }

        /// <summary>
        /// Definitions for UTChecker.setting
        /// </summary>
        public static class UTCheckerSetting
        {
            public const string FileName = "UTChecker.setting";
            public const string Prefix = "@SET";
            public const string ListFile = "LIST_FILE";
            public const string TDSPath = "TDS_PATH";
            public const string OutputPath = "OUTPUT_PATH";
            public const string ReportTemplate = "REPORT_TEMPLATE";
            public const string SummaryTemplate = "SUMMARY_TEMPLATE";
            public const string TestLogPath = "TESTLOG_PATH";
            public const string ReferenceListsPath = "REFERENCE_LISTS_PATH";
            public const string SUTS_PATH = "SUTS_PATH";
            public const string SURR_PATH = "SUTRR_PATH";
        }


        /// <summary>
        /// An enum for Test Type
        /// </summary>
        public enum TestType
        {
            [Description("N/A")]
            ByMockito = 1,

            [Description("By PowerMockito")]
            ByPowerMockito,

            [Description("VectorCast")]
            ByVectorCast,

            [Description("Getter/Setter")]
            GetterSetter,

            [Description("Empty method")]
            Empty,

            [Description("Abstract method")]
            Abstract,

            [Description("Interface method")]
            Interface,

            [Description("Native method")]
            Native,

            [Description("Pure function calls")]
            PureFunctionCalls,

            [Description("Pure UI function calls")]
            PureUIFunctionCalls,

            [Description("By code analysis")]
            ByCodeAnalysis,

            [Description("Unknow")]
            Unknow,
        }


        /// <summary>
        /// Get the string value of TestType.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string GetStringValue(TestType value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());

            DescriptionAttribute[] attributes =
                (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);

            if (attributes != null && attributes.Length > 0)
            {
                return attributes[0].Description;
            }
            else
            {
                return value.ToString();
            }
        }



        /// <summary>
        /// The meaing of Test
        /// </summary>
        public enum TestMeans
        {
            // Unknow 
            UNKNOWN = 0, // e.g. invalid test case defined in TDS

            // tested by a script (Mockito / PowerMockito)
            TEST_SCRIPT,

            // tested by code analysis (code review)
            CODE_ANALYSIS,

            // no need test, e.q. empty method, setter/getter
            NA,
        }




        public struct TestCaseItem
        {
            public string sSourceFileName;
            public string sMethodName;
            public string sTDSFileName;
            public string sTCSourceFileName;
            public string sTCLabelName;
            public string sTCFuncName;
            public string sTCNote;

            public bool bIsRepeated;        // Only set it as true if the TC is testing the same method.
                                            // That is, it will be false if the TC is for testing multiplt methods.
            public TestMeans eTestMeans;
            public TestLog eTestlog;
            public TestType eType;

            public string sChapterInSUTS;

        };



        //
        // A struct which contents the infomration about the numbers of test type
        //
        public struct TestTypeStatistic
        {

            // by test scripts
            public int mockito;
            public int powermockito;

            // no test needed
            public int gettersetter;
            public int emptymethod;
            public int abstractmethod;
            public int interfacemethod;
            public int nativemethod;


            // by code analysis
            public int codeanalysis;
            public int purefunctioncalls;
            public int pureUIfunctioncalls;


            public int vectorcast;

            // unknow item
            public int unknow;

        };



        public struct TestCaseTable
        {
            public int dSourceFileCount;    // total # of non-repeated source files
            public int dMethodCount;        // total # of non-repeated methods
            public int dTestCaseFuncCount;  // total # of non-repeated TC functions

            // Test entry counters:
            public int dNormalEntryCount;   // total # of non-repeated TC labels
            public int dRepeatedEntryCount; // total # of repeated (TC label) entries
            public int dErrorEntryCount;    // total # of entries whose TC label is "ERROR..."
            // Note: Sum of above 3 counters shall == # of entries read from EXCEL tables.

            // TC execution means counters:
            public int dByNACount;          // no test needed
            public int dByTestScriptCount;  // tested via scripts
            public int dByCodeAnalysisCount;// tested via code analysis
            public int dByUnknownCount;     // tested via other means
            // Note: Sum of above 4 counters shall == dNormalEntryCount

            public int dErrorCount;         // total # of errors found
            public int dNGEntryCount;       // total # of NG entries (EXCEL rows)

            public int dTestLogIssueCount;
            public int dSUTSIssueCount;

            public List<TestCaseItem> ltItems;

            public TestTypeStatistic stTestTypeStatistic;

        };


        

        /// <summary>
        /// return Code
        /// </summary>
        public int ReturnCode { get; private set; } = -100;


        /// <summary>
        /// BackgroundWorker handler for UTChecker
        /// </summary>
        private BackgroundWorker g_bwUTChecker;


        public EnvrionmentSetting g_FilePathSetting { get; internal set; }



        public RunMode Mode { get; private set; } = RunMode.None;

        /// <summary>
        /// A handler for Excel
        /// </summary>
        private static Excel.Application g_excelApp = null;

        /// <summary>
        /// A handler for Word
        /// </summary>
        private static Word.Application g_wordApp = null;


        static public TestCaseTable g_tTestCaseTable;


        static string g_sModuleListFile = string.Empty;
        static string g_sTDSPath = string.Empty;
        static string g_sOutputPath = string.Empty;
        static string g_sTemplateFile = string.Empty;
        static string g_sSummaryReport = string.Empty;

        static string g_sTestLogPath = string.Empty;
        static string g_sSUTSPath = string.Empty;
        static string g_sSUTSDocumentPath = string.Empty;
        static string g_sReferenceListsPath = string.Empty;
        static string g_sSUTRRPath = string.Empty;

        /// <summary>
        /// A string that is used to be a path where the Log file saved at.
        /// </summary>
        static string g_sErrorLogFile = "";

        static List<string> g_lsModules = null;
        static List<string> g_lsTDSFiles = null;


        static List<TestLog> g_lsTestLogs = null;

    }
}
