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

            // Indices for the summary sheet in Template file.
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

            // Indices for the lookup sheet in Template file.
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
                TEST_LOG,
                TEST_RESULT,
                TEST_LOG_PATH,
                SUTS,
            }

            public const int HEADER_HEIGHT = 2;
            public const int COUNT_ROW = HEADER_HEIGHT;
            public const int FIRST_ROW = HEADER_HEIGHT + 1;
        }

    }
}
