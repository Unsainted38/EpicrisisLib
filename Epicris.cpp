#include "pch.h"

namespace unsaintedWinAppLib {
    Epicris::Epicris()
    {
        analyzesList = gcnew List<JObject^>();
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
            missingFields += "Номер истории болезни\n";
        }
        if (String::IsNullOrEmpty(name)) {
            missingFields += "Имя\n";
        }
        if (String::IsNullOrEmpty(surname)) {
            missingFields += "Фамилия\n";
        }
        if (String::IsNullOrEmpty(patronymic)) {
            missingFields += "Отчество\n";
        }
        if (String::IsNullOrEmpty(rank)) {
            missingFields += "Звание\n";
        }
        if (String::IsNullOrEmpty(militaryUnit)) {
            missingFields += "Войсковая часть\n";
        }
        if (String::IsNullOrEmpty(birthday)) {
            missingFields += "Дата рождения\n";
        }
        if (String::IsNullOrEmpty(incomeDate)) {
            missingFields += "Дата поступления\n";
        }
        if (String::IsNullOrEmpty(outcomeDate)) {
            missingFields += "Дата выписки\n";
        }
        if (String::IsNullOrEmpty(mkb)) {
            missingFields += "МКБ\n";
        }
        if (String::IsNullOrEmpty(diagnosis)) {
            missingFields += "Диагноз\n";
        }
        if (String::IsNullOrEmpty(relatedDiagnosis)) {
            missingFields += "Сопутствующий диагноз\n";
        }
        if (String::IsNullOrEmpty(complications)) {
            missingFields += "Осложнения\n";
        }
        /*if (String::IsNullOrEmpty(anamnesisJson)) {
            missingFields += "anamnesisJson\n";
        }*/
        if (String::IsNullOrEmpty(anamnesisText)) {
            missingFields += "Анамнез\n";
        }
        if (String::IsNullOrEmpty(analyzesListJson)) {
            missingFields += "Анализы\n";
        }
        /*if (String::IsNullOrEmpty(additionalData)) {
            missingFields += "Дополнительные данные (лечение и осмотр врачами)\n";
        }*/
        if (therapy->Count == 0) {
            missingFields += "Лечение\n";
        }
        if (doctorsLooked->Count == 0) {
            missingFields += "Осмотрен\n";
        }
        if (String::IsNullOrEmpty(sideData)) {
            missingFields += "Дополнительно\n";
        }
        if (String::IsNullOrEmpty(recommendations)) {
            missingFields += "Рекомендации\n";
        }
        if (String::IsNullOrEmpty(unworkableList)) {
            missingFields += "Лист нетрудоспособности\n";
        }
        if (String::IsNullOrEmpty(illBeginDate)) {
            missingFields += "Дата начала болезни\n";
        }
        if (String::IsNullOrEmpty(vvk)) {
            missingFields += "ВВК";
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

    void Epicris::AddAnalysisToAnalyzesList(JObject^ analyzes) {
        JsonSerializerSettings^ settings = gcnew JsonSerializerSettings();
        settings->NullValueHandling = NullValueHandling::Ignore;        
        List<JObject^>^ jobList = gcnew List<JObject^>();
        jobList->AddRange(analyzesList);
        if (jobList->Count == 0)
            analyzesList->Add(analyzes);
        bool flag = false;
        for each (JObject ^ ob in jobList) {
            if (analyzes->Value<String^>("id") == ob->Value<String^>("id")) {
                analyzesList->Remove(ob);
                analyzesList->Add(analyzes);
                flag = false;
                break;
            }
            else {
                flag = true;
            }
        }
        if (flag)
            analyzesList->Add(analyzes);
        analyzesListJson = Newtonsoft::Json::JsonConvert::SerializeObject(analyzesList, settings);
    }
    void Epicris::AddAnalysisToAnalyzesList(Dictionary<String^, JObject^>^ analyzes) {
        JsonSerializerSettings^ settings = gcnew JsonSerializerSettings();
        settings->NullValueHandling = NullValueHandling::Ignore;
        analyzesListJson = JsonConvert::SerializeObject(analyzes, settings);
    }
    void Epicris::Clear()
    {
        historyNumber++;
        historyYear = String::Empty;
        name = String::Empty;
        surname = String::Empty;
        patronymic = String::Empty;
        rank = String::Empty;
        militaryUnit = String::Empty;
        birthday = String::Empty;
        incomeDate = String::Empty;
        outcomeDate = String::Empty;
        mkb = String::Empty;
        diagnosis = String::Empty;
        relatedDiagnosis = String::Empty;
        complications = String::Empty;
        anamnesisJson = String::Empty;
        anamnesisText = String::Empty;
        analyzesList->Clear();
        additionalData = String::Empty;
        therapy->Clear();
        doctorsLooked->Clear();
        sideData = String::Empty;
        recommendations = String::Empty;
        unworkableList = String::Empty;
        illBeginDate = String::Empty;
        analyzesListJson = String::Empty;   
        vvk = String::Empty;
    }
}