#pragma once

using namespace System;
using namespace System::Windows::Forms;
using namespace System::Collections::Generic;
using namespace System::Data;
using namespace System::Data::SQLite;


namespace unsaintedWinAppLib {
    public enum class DataFormat {
        JSON,
        String,
        ListStr,
    };

    public enum class MinMax {
        Min,
        Max,
    };

    public ref class SQLiteDbHelper {

    private:
        String^ connectionString;
        String^ tmp_query;
        String^ tmp_table;
        String^ tmp_column;
        String^ tmp_value;
        String^ tmp_conditionColumn;
        String^ tmp_conditionValue;

        void GenerateNonQuery();
        void ResetQuery();

    public:
        SQLiteDbHelper(String^ dbPath);

        String^ SetQueryById(String^ table, String^ column, int id);
        String^ SetQueryByTitle(String^ talble, String^ column, String^ title);
        String^ SetQueryByCondition(String^ table, String^ column, String^ conditionColumn, String^ conditionValue, DataFormat format);
        List<String^>^ SetQueryByConditionLike(String^ table, String^ column, String^ conditionColumn, String^ conditionValue);
        List<String^>^ SetQueryByCondition(String^ table, String^ column, String^ conditionColumn, String^ conditionValue);
        String^ GetJsonString();
        String^ GetMinMaxColumnData(String^ table, String^ column, MinMax min_max);

        List<String^>^ GetColumnData();
        List<String^>^ GetColumnData(String^ tableName, String^ columnName);
        List<String^>^ GetSortedColumnData(String^ tableName, String^ columnName, int sortType);

        Dictionary<String^, String^>^ ExtractColumnsToDictionary(String^ table, String^ keyColumn, String^ valueColumn);
        Dictionary<String^, Object^>^ ExtractAnalyzesBlankToDictionary(String^ title);
        void SetNonQuery(String^ table, String^ destinyColumn, String^ destinyValue, String^ conditionColumn, String^ conditionValue);

        void ImportDictToDb(Dictionary<String^, String^>^ analyzesDict);
        void ImportRtfToDb(String^ key, String^ rtf);
        void ImportRtfToDb(String^ table, String^ dectColumn, String^ keyColumn, String^ rtf);
    };
}