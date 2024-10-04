#include "pch.h"

namespace unsaintedWinAppLib {
    Epicris::Epicris()
    {
        analyzesList = gcnew List<String^>();
        doctorsLooked = gcnew List<String^>();
        therapy = gcnew List<String^>();
    }

    bool Epicris::CheckEpicrisFields() {
        // Строка для сбора информации о незаполненных полях
        String^ missingFields = "Следующие поля не заполнены:\n";

        // Проверка всех приватных полей класса Epicris               
        if (historyNumber == 0) {
            missingFields += "historyNumber\n";
        }
        if (String::IsNullOrEmpty(historyYear)) {
            missingFields += "historyYear\n";
        }
        if (String::IsNullOrEmpty(name)) {
            missingFields += "name\n";
        }
        if (String::IsNullOrEmpty(surname)) {
            missingFields += "surname\n";
        }
        if (String::IsNullOrEmpty(patronymic)) {
            missingFields += "patronymic\n";
        }
        if (String::IsNullOrEmpty(rank)) {
            missingFields += "rank\n";
        }
        if (String::IsNullOrEmpty(militaryUnit)) {
            missingFields += "militaryUnit\n";
        }
        if (String::IsNullOrEmpty(birthday)) {
            missingFields += "birthday\n";
        }
        if (String::IsNullOrEmpty(incomeDate)) {
            missingFields += "incomeDate\n";
        }
        if (String::IsNullOrEmpty(outcomeDate)) {
            missingFields += "outcomeDate\n";
        }
        if (String::IsNullOrEmpty(mkb)) {
            missingFields += "mkb\n";
        }
        if (String::IsNullOrEmpty(diagnosis)) {
            missingFields += "diagnosis\n";
        }
        if (String::IsNullOrEmpty(relatedDiagnosis)) {
            missingFields += "relatedDiagnosis\n";
        }
        if (String::IsNullOrEmpty(complications)) {
            missingFields += "complications\n";
        }
        if (String::IsNullOrEmpty(anamnesisJson)) {
            missingFields += "anamnesisJson\n";
        }
        if (String::IsNullOrEmpty(anamnesisText)) {
            missingFields += "anamnesisText\n";
        }
        if (analyzesList == nullptr || analyzesList->Count == 0) {
            missingFields += "analyzesList\n";
        }
        if (String::IsNullOrEmpty(additionalData)) {
            missingFields += "additionalData\n";
        }
        if (therapy->Count == 0) {
            missingFields += "therapy\n";
        }
        if (doctorsLooked->Count == 0) {
            missingFields += "doctorsLooked\n";
        }
        if (String::IsNullOrEmpty(sideData)) {
            missingFields += "sideData\n";
        }
        if (String::IsNullOrEmpty(recommendations)) {
            missingFields += "recommendations\n";
        }
        if (String::IsNullOrEmpty(unworkableList)) {
            missingFields += "unworkableList\n";
        }
        if (String::IsNullOrEmpty(illBeginDate)) {
            missingFields += "illBeginDate\n";
        }

        // Если есть незаполненные поля, выводим их в MessageBox
        if (missingFields != "Следующие поля не заполнены:\n") {
            MessageBox::Show(missingFields, "Незаполненные поля", MessageBoxButtons::OK, MessageBoxIcon::Warning);
            return false;
        }
        else {
            MessageBox::Show("Все поля заполнены.", "Информация", MessageBoxButtons::OK, MessageBoxIcon::Information);
            return true;
        }
    }

    void Epicris::CheckProperty()
    {

    }

    void Epicris::AddAnalysisToAnalyzesList(Dictionary<String^, Object^>^ analyzesDict) {
        JsonSerializerSettings^ settings = gcnew JsonSerializerSettings();
        settings->NullValueHandling = NullValueHandling::Ignore;
        String^ json = JsonConvert::SerializeObject(analyzesDict, settings);
        analyzesList->Add(json);
    }
    void Epicris::Clear()
    {
        historyNumber++;
        historyYear = nullptr;
        name = nullptr;
        surname = nullptr;
        patronymic = nullptr;
        rank = nullptr;
        militaryUnit = nullptr;
        birthday = nullptr;
        incomeDate = nullptr;
        outcomeDate = nullptr;
        mkb = nullptr;
        diagnosis = nullptr;
        relatedDiagnosis = nullptr;
        complications = nullptr;
        anamnesisJson = nullptr;
        anamnesisText = nullptr;
        analyzesList->Clear();
        additionalData = nullptr;
        therapy = nullptr;
        doctorsLooked = nullptr;
        sideData = nullptr;
        recommendations = nullptr;
        unworkableList = nullptr;
        illBeginDate = nullptr;
    }
}