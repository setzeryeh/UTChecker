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
            public const string FILE_NAME = "Method_TC_Lookup_Table_Summary.xlsx";
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

                GETTER_SETTER,
                EMPTY,
                INTERFACE,
                ABSTRACE,
                NATIVE,

                MOCKITO,
                POWERMOCKIT,

                BY_CODE_ANALYSIS,
                PURE_CALL,
                PURE_UI_CALL,

                UNKNOW,

                TOTAL_TESTCASE_COUNT,
                NORMAL_ENTRY,
                REPEATED_ENTRY,
                ERROR_ENTRY,
                NG_COUNT,

                ERROR_COUNT,

                LOGS_ERROR_COUNT,

            }
        }

    }
}
