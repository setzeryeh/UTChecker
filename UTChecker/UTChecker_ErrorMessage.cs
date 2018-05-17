using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UTChecker
{
    public partial class UTChecker
    {
        public static class ErrorMessage
        {
            // for office
            public const string EXCEL_APP_IS_NULL = "The handler of EXCEL Application is null.";
            public const string WORD_APP_IS_NULL = "The handler of WORD Application is null.";
            public const string SUTS_DOC_IS_NULL = "The handler of SUTS Document is null.";


            public const string NO_ENTRY_TO_BE_SAVED = "No item to be saved.";
            public const string OUTPUT_FILE_IS_NULL = "Output file is null.";

            public const string CANNOT_FIND_TEMPLATE = "Cannot find template";

            public const string VALUE_UNKNOWN = "Unknown";

            public const string INVLAID_SOURCE_FILE_NAME = "Invalid source file name";

            public const string INVLAID_METHOD_NAME = "Invalid method name";
            public const string METHOD_NAME_SHALL_NOT_BE_NA = "Method name shall not be N/A.";
            public const string METHOD_NAME_SHALL_NOT_BE_EMPTY = "Method name shall not be empty.";
            public const string METHOD_NAME_SHALL_NOT_CONTAIN_SPACE = "Method name shall not contain any space";

            public const string INVLAID_TC_LABEL = "Invalid TC label";
            public const string TC_LABEL_SHALL_NOT_BE_NA = "TC label shall not be N/A.";
            public const string TC_LABEL_SHALL_NOT_BE_EMPTY = "TC label shall not be empty.";
            public const string TC_LABEL_SHALL_NOT_CONTAIN_SPACE = "TC label shall not contain any space";
            public const string DUPLICATE_TC_LABEL_FOUND = "Duplicate TC label found";

            public const string INVLAID_TC_FUNC_NAME = "Invalid TC func name";
            public const string TC_FUNC_NAME_SHALL_NOT_CONTAIN_SPACE = "TC func name shall not contain any space";
            public const string NO_TC_FUNC_NAME_CAN_BE_READ = "No TC func name can be read.";
            public const string REASON_SHALL_BE_GIVEN_FOR_NA_TC_FUNC = "Reason shall be given for N/A TC func.";

            public const string TC_TEST_MEANS_SHALL_NOT_BE_UNKNOWN = "TC test means shall not be unknown";


            public const string TESTLOG_IS_MISSING = "The Test Log is missing";

            public const string AMBIGUOUS_BETWEEN_TCFUN_TCNOT = "Ambiguous test between function and Note.";
            public const string CLASS_SHALL_BE_DEFINED_IN_SUTS = "Class shall be defined in SUTS";

        }
    }
}
