using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace UTChecker
{
    public partial class TDS_Parser
    {
        public static class Constants
        {
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

            public class Color
            {
                static public System.Int32 RED = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            }

            public const int UTCheckerSelf = 1;   // UTChecker self.
            public const int ArgumentsCount = 5;
            public const int ArgumentsMatchLength = ArgumentsCount + UTCheckerSelf;



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
            public const string TestLogs = "TESTLOGS_PATH";

        }


        /// <summary>
        /// The definition of Note in TDS
        /// </summary>
        public static class TestType
        {
            public const string ByPowerMocktio = "By PowerMockito";
            public const string ByMockito = "N/A";
            public const string GetterSetter = "Getter/Setter";
            public const string Empty = "Empty method";
            public const string Abstract = "Abstract method";
            public const string Interface = "Interface method";
            public const string Native = "Native method";
            public const string PureFunctionCalls = "Pure function calls";
            public const string PureUIFunctionCalss = "Pure UI function calls";
            public const string ByCodeAnalysis = "By code analysis";
        }


        public enum TestMeans
        {
            UNKNOWN = 0, // e.g. invalid test case defined in TDS
            TEST_SCRIPT,
            CODE_ANALYSIS,
            NA // empty method, setter/getter
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
            public TestMeans eTestMeans;
            public bool bIsRepeated;        // Only set it as true if the TC is testing the same method.
                                            // That is, it will be false if the TC is for testing multiplt methods.
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

            public List<TestCaseItem> ltItems;
        };


        //
        // A struct which contents the infomration about the numbers of test type
        //
        public struct ModuleInfo
        {
            public string name;

            // no test needed
            public int gettersetter;
            public int emptymethod;
            public int abstractmethod;
            public int interfacemethod;
            public int nativemethod;

            // by test scripts
            public int mockito;
            public int powermockito;

            // by code analysis
            public int codeanalysis;
            public int purefunctioncalls;
            public int pureUIfunctioncalls;

            // unknow item
            public int unknow;

            public int count;
            public TestCaseTable testCase;
        }


        /// <summary>
        /// 
        /// </summary>
        public struct PathSetting
        {
            public string listFile;
            public string tdsPath;
            public string outputPath;
            public string reportTemplate;
            public string summaryTemplate;
            public string testlogsPath;

        }

        /// <summary>
        /// 
        /// </summary>
        public enum RunBy
        {
            // run TDS_Parser by CommandLine
            CommandLine,

            // run TDS_Parser by Window/User
            User,
        }

        public PathSetting g_FilePathSetting { get; internal set; }



        public RunBy RunUTCheckerBy { get; private set; }


        private BackgroundWorker g_bwTDSParse;


        static public TestCaseTable g_tTestCaseTable;

        static Excel.Application g_excelApp = null;

        static string g_sTDSPath = "";
        static string g_sOutputPath = "";
        static string g_sModuleListFile = "";
        static string g_sTemplateFile = "";
        static string g_sSummaryReport = "";

        static string g_sErrorLogFile = "";



        static List<string> g_lsModules = null;
        static List<string> g_lsTDSFiles = null;
        static List<ModuleInfo> g_lsModuleInfo = null;


        /// <summary>
        ///  variables which content the numbers for each test type
        /// </summary>
        static int gn_ByMockito = 0;
        static int gn_ByPowerMockito = 0;
        static int gn_Bycodeanalysis = 0;
        static int gn_GetterSetter = 0;
        static int gn_Emptymethod = 0;
        static int gn_Abstractmethod = 0;
        static int gn_Interfacemethod = 0;
        static int gn_Nativemethod = 0;
        static int gn_Purefunctioncalls = 0;
        static int gn_PureUIfunctioncalls = 0;

        static int gn_Unknow = 0;
    }
}
