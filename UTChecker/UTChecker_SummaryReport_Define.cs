using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UTChecker
{
    public partial class UTChecker
    {

        static class SummaryReport
        {
            public const string FILE_NAME = "UT_CHECK_Summary.xlsx";
            public const string SHEET_NAME = "Summary";

            public const int COLUMN_COUNT = (int)ColumnIndex.ERROR_COUNT;
            public const int HEADER_HEIGHT = 2;
            public const int COUNT_ROW = HEADER_HEIGHT;
            public const int FIRST_ROW = HEADER_HEIGHT + 1;

            public enum ColumnIndex
            {
                INDEX = 1,
                MODULE_NAME,

                SOURCE_COUNT,
                METHOD_COUNT,
                TESTCASE_COUNT,

                MOCKITO,
                POWERMOCKIT,

                VECTORCAST,

                GETTER_SETTER,
                EMPTY,
                INTERFACE,
                ABSTRACE,
                NATIVE,

                BY_CODE_ANALYSIS,
                PURE_CALL,
                PURE_UI_CALL,

                UNKNOW,

                TOTAL_TESTCASE_COUNT,
                NORMAL_ENTRY,
                REPEATED_ENTRY,
                ERROR_ENTRY,

                ERROR_COUNT,
                NG_COUNT,
                TESTLOG_ISSUE_COUNT,
                SUTS_ISSUE_COUNT,
            }
        }

    }
}
