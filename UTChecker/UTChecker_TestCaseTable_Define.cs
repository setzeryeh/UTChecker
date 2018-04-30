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

        static class TestCaseTableConstants
        {
            public const string INPUT_FILENAME_PREFIX = "NUGEN Test Data Sheet - ";
            public const string INPUT_FILE_EXT_NAME = "*.xlsx";
            public const string FILENAME_PREFIX = "Method_TC_Lookup_Table_of_";
            public const string SHEET_NAME = "LookupTable";

            // Indices for the summary sheet
            public enum RowIndex
            {
                DOC_NAME = 1,
                DATE_TIME,

                SOURCE_FILE_COUNT,
                METHOD_COUNT,
                TC_COUNT,

                TC_TEST_VIA_NA_COUNT,
                TC_TEST_VIA_SCRIPT_COUNT,
                TC_TEST_VIA_ANALYSIS_COUNT,
                TC_TEST_VIA_OTHERS_COUNT,

                TC_FUNC_COUNT,
                ERROR_COUNT,
            }

            // Indices for the lookup sheet
            public enum ColumnIndex
            {
                NG_MARKER = 1, // for marking & filtering NG entries
                SOURCE_FILE,
                METHOD_NAME,
                TC_LABEL,
                TC_NAME,
                TDS_FILE,
                TC_SOURCE_FILE,
                NOTE,
            }

            public const int HEADER_HEIGHT = 2;
            public const int COUNT_ROW = HEADER_HEIGHT;
            public const int FIRST_ROW = HEADER_HEIGHT + 1;
        }

    }
}
