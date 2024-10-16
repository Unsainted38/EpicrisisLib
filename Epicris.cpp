#include "pch.h"

namespace unsaintedWinAppLib {
    Epicris::Epicris()
    {
        analyzesList = gcnew List<JObject^>();
        doctorsLooked = gcnew List<String^>();
        therapy = gcnew List<String^>();
    }

    bool Epicris::CheckEpicrisFields() {
        // ������ ��� ����� ���������� � ������������� �����
        String^ missingFields = "��������� ���� �� ���������:\n";

        // �������� ���� ��������� ����� ������ Epicris               
        if (historyNumber == 0) {
            missingFields += "historyNumber\n";
        }
        if (String::IsNullOrEmpty(historyYear)) {
            missingFields += "����� ������� �������\n";
        }
        if (String::IsNullOrEmpty(name)) {
            missingFields += "���\n";
        }
        if (String::IsNullOrEmpty(surname)) {
            missingFields += "�������\n";
        }
        if (String::IsNullOrEmpty(patronymic)) {
            missingFields += "��������\n";
        }
        if (String::IsNullOrEmpty(rank)) {
            missingFields += "������\n";
        }
        if (String::IsNullOrEmpty(militaryUnit)) {
            missingFields += "��������� �����\n";
        }
        if (String::IsNullOrEmpty(birthday)) {
            missingFields += "���� ��������\n";
        }
        if (String::IsNullOrEmpty(incomeDate)) {
            missingFields += "���� �����������\n";
        }
        if (String::IsNullOrEmpty(outcomeDate)) {
            missingFields += "���� �������\n";
        }
        if (String::IsNullOrEmpty(mkb)) {
            missingFields += "���\n";
        }
        if (String::IsNullOrEmpty(diagnosis)) {
            missingFields += "�������\n";
        }
        if (String::IsNullOrEmpty(relatedDiagnosis)) {
            missingFields += "������������� �������\n";
        }
        if (String::IsNullOrEmpty(complications)) {
            missingFields += "����������\n";
        }
        /*if (String::IsNullOrEmpty(anamnesisJson)) {
            missingFields += "anamnesisJson\n";
        }*/
        if (String::IsNullOrEmpty(anamnesisText)) {
            missingFields += "�������\n";
        }
        if (String::IsNullOrEmpty(analyzesListJson)) {
            missingFields += "�������\n";
        }
        /*if (String::IsNullOrEmpty(additionalData)) {
            missingFields += "�������������� ������ (������� � ������ �������)\n";
        }*/
        if (therapy->Count == 0) {
            missingFields += "�������\n";
        }
        if (doctorsLooked->Count == 0) {
            missingFields += "��������\n";
        }
        if (String::IsNullOrEmpty(sideData)) {
            missingFields += "�������������\n";
        }
        if (String::IsNullOrEmpty(recommendations)) {
            missingFields += "������������\n";
        }
        if (String::IsNullOrEmpty(unworkableList)) {
            missingFields += "���� ������������������\n";
        }
        if (String::IsNullOrEmpty(illBeginDate)) {
            missingFields += "���� ������ �������\n";
        }

        // ���� ���� ������������� ����, ������� �� � MessageBox
        if (missingFields != "��������� ���� �� ���������:\n") {
            MessageBox::Show(missingFields, "������������� ����", MessageBoxButtons::OK, MessageBoxIcon::Warning);
            return false;
        }
        else {
            MessageBox::Show("��� ���� ���������.", "����������", MessageBoxButtons::OK, MessageBoxIcon::Information);
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
        therapy->Clear();
        doctorsLooked->Clear();
        sideData = nullptr;
        recommendations = nullptr;
        unworkableList = nullptr;
        illBeginDate = nullptr;
        analyzesListJson = nullptr;
    }
}