#include "pch.h"
namespace unsaintedWinAppLib {
    SQLiteDbHelper::SQLiteDbHelper(String^ dbPath)
    {
        connectionString = "Data Source=" + dbPath + ";Version=3;";
    }
    // Функция создает строку sql запроса с условием подходящего id таблицы 
    String^ SQLiteDbHelper::SetQueryById(String^ table, String^ column, int id)
    {
        tmp_column = column;
        tmp_table = table;
        tmp_query = " SELECT DISTINCT " + column +
            " FROM " + table + " WHERE id = " + id;

        return tmp_query;
    }
    // Функция создает строку sql запроса с условием подходящего title таблицы 
    String^ SQLiteDbHelper::SetQueryByTitle(String^ table, String^ column, String^ title)
    {
        tmp_column = column;
        tmp_table = table;
        tmp_query = " SELECT DISTINCT " + column +
            " FROM " + table + " WHERE title = " + title;

        return tmp_query;
    }
    // Функция 
    String^ SQLiteDbHelper::SetQueryByCondition(String^ table, String^ column, String^ conditionColumn, String^ conditionValue, DataFormat format)
    {
        String^ result = gcnew String("");
        if (table == "" || column == "" || conditionColumn == "" || conditionValue == "") {
            return result;
        }
        tmp_column = column;
        tmp_table = table;
        tmp_query = " SELECT DISTINCT " + column +
            " FROM " + table + " WHERE " + conditionColumn + " = " + "'" + conditionValue + "'";
        switch (format)
        {
        case DataFormat::JSON:
            return GetJsonString();
            break;
        case DataFormat::String:

            break;
        case DataFormat::ListStr:
            break;
        default:
            break;
        }
        return result;
    }
    // Функция создает строку sql запроса с условием подходящего значения выбранного столбца и выполняет его
    // Возвращает полученные значения в виде списка строк 
    List<String^>^ SQLiteDbHelper::SetQueryByCondition(String^ table, String^ column, String^ conditionColumn, String^ conditionValue)
    {
        if (table == "" || column == "" || conditionColumn == "" || conditionValue == "") {
            List<String^>^ result = gcnew List<String^>();
            result->Add("");
            return result;

        }
        tmp_column = column;
        tmp_table = table;
        tmp_query = " SELECT DISTINCT " + column +
            " FROM " + table + " WHERE " + conditionColumn + " = " + "'" + conditionValue + "'";

        return GetColumnData();
    }
    // Функция извлекает строку в формате json из таблицы
    String^ SQLiteDbHelper::GetJsonString()
    {
        SQLiteConnection^ connection = gcnew SQLiteConnection(connectionString);
        String^ jsonData;
        try {
            connection->Open();
            SQLiteCommand^ cmd = gcnew SQLiteCommand(tmp_query, connection);
            SQLiteDataReader^ reader = cmd->ExecuteReader();

            if (reader->Read()) {
                jsonData = reader[tmp_column]->ToString();
            }
            reader->Close();
        }
        catch (Exception^ ex) {
            Console::WriteLine("Error: " + ex->Message);
            return jsonData;
        }
        finally {
            connection->Close();
            ResetQuery();
        }
        return jsonData;
    }
    // Функция создает tmp_query типа "SELECT DISTINCT column FROM table WHERE conditionalColumn LIKE 'conditionValue%'"
    List<String^>^ SQLiteDbHelper::SetQueryByConditionLike(String^ table, String^ column, String^ conditionColumn, String^ conditionValue)
    {
        List<String^>^ result;
        if (table == "" || column == "" || conditionColumn == "" || conditionValue == "") {
            result = gcnew List<String^>();
            result->Add("");
            return result;
        }
        tmp_column = column;
        tmp_table = table;
        tmp_query = " SELECT DISTINCT " + column +
            " FROM " + table + " WHERE " + conditionColumn + " LIKE " + "'" + conditionValue + "%" + "'";

        return GetColumnData();
    }
    // Функция выполняет sql запрос tmp_query, которую надо предварительно инициализировать
    // Возвращает список строк из колонки таблицы
    List<String^>^ SQLiteDbHelper::GetColumnData()
    {
        List<String^>^ results = gcnew List<String^>();

        SQLiteConnection^ connection = gcnew SQLiteConnection(connectionString);
        try
        {
            connection->Open();


            SQLiteCommand^ command = gcnew SQLiteCommand(tmp_query, connection);
            SQLiteDataReader^ reader = command->ExecuteReader();

            while (reader->Read())
            {
                results->Add(reader[tmp_column]->ToString());
            }

            reader->Close();
        }
        catch (Exception^ ex)
        {
            Console::WriteLine("Error: " + ex->Message);
        }
        finally
        {
            connection->Close();
            ResetQuery();
        }
        if (results->Count == 0) results->Add("");
        return results;
    }
    // Функция извлекает данные из столбца(columnName) таблицы(tableName) и возвращает список строк.
    List<String^>^ SQLiteDbHelper::GetColumnData(String^ tableName, String^ columnName)
    {
        List<String^>^ results = gcnew List<String^>();

        SQLiteConnection^ connection = gcnew SQLiteConnection(connectionString);
        try
        {
            connection->Open();

            String^ query = " SELECT DISTINCT " + columnName +
                " FROM " + tableName;


            SQLiteCommand^ command = gcnew SQLiteCommand(query, connection);
            SQLiteDataReader^ reader = command->ExecuteReader();

            while (reader->Read())
            {
                results->Add(reader[columnName]->ToString());
            }

            reader->Close();
        }
        catch (Exception^ ex)
        {
            Console::WriteLine("Error: " + ex->Message);
        }
        finally
        {
            connection->Close();
            ResetQuery();
        }
        return results;
    }
    // Функция возвращает отсортированный список строк по убыванию или возрастанию
    List<String^>^ SQLiteDbHelper::GetSortedColumnData(String^ tableName, String^ columnName, int sortOrder)
    {
        List<String^>^ results = gcnew List<String^>();

        SQLiteConnection^ connection = gcnew SQLiteConnection(connectionString);
        try
        {
            connection->Open();
            String^ query;
            switch (sortOrder) {
            case 0:
                query = " SELECT DISTINCT " + columnName +
                    " FROM " + tableName;
                break;
            case 1:
                query = " SELECT DISTINCT " + columnName +
                    " FROM " + tableName +
                    " ORDER BY " + columnName + " ASC ";
                break;
            case 2:
                query = " SELECT DISTINCT " + columnName +
                    " FROM " + tableName +
                    " ORDER BY " + columnName + " DESC ";
                break;
            default:
                break;
            }


            SQLiteCommand^ command = gcnew SQLiteCommand(query, connection);
            SQLiteDataReader^ reader = command->ExecuteReader();

            while (reader->Read())
            {
                results->Add(reader[columnName]->ToString());
            }

            reader->Close();
        }
        catch (Exception^ ex)
        {
            Console::WriteLine("Error: " + ex->Message);
        }
        finally
        {
            connection->Close();
            ResetQuery();
        }
        return results;
    }
    Dictionary<String^, String^>^ SQLiteDbHelper::ExtractColumnsToDictionary(String^ table, String^ keyColumn, String^ valueColumn)
    {
        Dictionary<String^, String^>^ dict = gcnew Dictionary<String^, String^>();
        String^ query = "SELECT " + keyColumn + ", " + valueColumn + " FROM " + table;
        SQLiteConnection^ connection = gcnew SQLiteConnection(connectionString);
        try
        {
            connection->Open();
            SQLiteCommand^ command = gcnew SQLiteCommand(query, connection);
            SQLiteDataReader^ reader = command->ExecuteReader();

            while (reader->Read())
            {
                String^ key = reader[keyColumn]->ToString();
                String^ value = reader[valueColumn]->ToString();
                dict->Add(key, value);
            }
            reader->Close();
        }
        catch (Exception^ ex)
        {
            Console::WriteLine("Error: " + ex->Message);
        }
        finally
        {
            connection->Close();
        }
        return dict;
    }
    Dictionary<String^, Object^>^ SQLiteDbHelper::ExtractAnalyzesBlankToDictionary(String^ title)
    {
        Dictionary<String^, Object^>^ result = gcnew Dictionary<String^, Object^>();
        String^ query = "SELECT id, title, position FROM analyzes WHERE title = @title";

        SQLiteConnection^ connection = gcnew SQLiteConnection(connectionString);
        try
        {
            connection->Open();
            auto cmd = gcnew SQLiteCommand(query, connection);
            cmd->Parameters->AddWithValue("@title", title);
            SQLiteDataReader^ reader = cmd->ExecuteReader();
            while (reader->Read()) {
                result->Add("id", reader["id"]->ToString());
                result->Add("title", reader["title"]->ToString());
                result->Add("position", reader["position"]->ToString());

            }
        }
        catch (Exception^ e)
        {
            Console::WriteLine("Error message: " + e);
        }
        finally {
            connection->Close();
        }
        return result;
    }
    void SQLiteDbHelper::SetNonQuery(String^ table, String^ destinyColumn, String^ destinyValue,
        String^ conditionColumn, String^ conditionValue)
    {

        tmp_table = table;
        tmp_column = destinyColumn;
        tmp_value = destinyValue;
        tmp_conditionColumn = conditionColumn;
        tmp_conditionValue = conditionValue;
        tmp_query = "UPDATE " + tmp_table + " SET " + tmp_column + " = @value WHERE " + tmp_conditionColumn + " = @conditionValue";
        GenerateNonQuery();
    }
    void SQLiteDbHelper::ImportRtfToDb(String^ key, String^ rtf)
    {
        SQLiteConnection^ connection = gcnew SQLiteConnection(connectionString);
        String^ query = "UPDATE analyzes SET rtfValue = @rtf WHERE title = @key";

        try
        {
            connection->Open();
            SQLiteCommand^ cmd = gcnew SQLiteCommand(query, connection);
            cmd->Parameters->AddWithValue("@rtf", rtf);
            cmd->Parameters->AddWithValue("@key", key);
            cmd->ExecuteNonQuery();

        }
        catch (Exception^ ex) {
            Console::WriteLine("Error: " + ex->Message);
            return;
        }
        finally {
            connection->Close();
        }
    }
    void SQLiteDbHelper::ImportRtfToDb(String^ table, String^ dectColumn, String^ keyColumn, String^ rtf)
    {
    }
    void SQLiteDbHelper::GenerateNonQuery() {
        SQLiteConnection^ connection = gcnew SQLiteConnection(connectionString);

        try
        {
            connection->Open();
            SQLiteCommand^ cmd = gcnew SQLiteCommand(tmp_query, connection);
            cmd->Parameters->AddWithValue("@value", tmp_value);
            cmd->Parameters->AddWithValue("@conditionValue", tmp_conditionValue);
            cmd->ExecuteNonQuery();

        }
        catch (Exception^ ex) {
            Console::WriteLine("Error: " + ex->Message);
            return;
        }
        finally {
            connection->Close();
            ResetQuery();
        }
    }
    void SQLiteDbHelper::ImportDictToDb(Dictionary<String^, String^>^ analyzesDict) {
        for each (KeyValuePair<String^, String^> ^ kvp in analyzesDict) {
            ImportRtfToDb(kvp->Key, kvp->Value);
        }
    }
    void SQLiteDbHelper::ResetQuery() {
        tmp_column = "";
        tmp_query = "";
        tmp_table = "";
        tmp_column = "";
        tmp_conditionColumn = "";
        tmp_conditionValue = "";
    }
    String^ SQLiteDbHelper::GetMinMaxColumnData(String^ table, String^ column, MinMax min_max)
    {
        SQLiteConnection^ connection = gcnew SQLiteConnection(connectionString);
        String^ Data;

        tmp_table = table;
        switch (min_max)
        {
        case MinMax::Min:
            tmp_query = "SELECT DISTINCT MIN(" + column + ")" + " FROM " + table;
            tmp_column = "MIN(" + column + ")";
            break;
        case MinMax::Max:
            tmp_query = "SELECT DISTINCT MAX(" + column + ")" + " FROM " + table;
            tmp_column = "MAX(" + column + ")";
            break;
        default:
            break;
        }

        try {
            connection->Open();
            SQLiteCommand^ cmd = gcnew SQLiteCommand(tmp_query, connection);
            SQLiteDataReader^ reader = cmd->ExecuteReader();

            if (reader->Read()) {
                Data = reader[tmp_column]->ToString();
            }
            reader->Close();
        }
        catch (Exception^ ex) {
            Console::WriteLine("Error: " + ex->Message);
            return Data;
        }
        finally {
            connection->Close();
            ResetQuery();
        }
        return Data;
    }
};