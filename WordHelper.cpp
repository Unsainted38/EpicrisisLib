#include "pch.h"

namespace unsaintedWinAppLib {
    
    WordHelper::WordHelper(String^ templateFilePath, String^ outputDirPath, Epicris^ epicris)
    {
        m_epicris = epicris;
        m_templateFile = templateFilePath;
        m_outputDir = outputDirPath;   
    }

    WordHelper::WordHelper(String^ templateFilePath, String^ outputDirPath, DoctorsRecords^ docRec)
    {
        throw gcnew System::NotImplementedException();
    }

    WordHelper::WordHelper(String^ templateFilePath, String^ outputDirPath, FirstList^ firstList)
    {
        throw gcnew System::NotImplementedException();
    }

    WordHelper::~WordHelper()
    {
    }

    WordHelper::WordHelper()
    {
        
    }

    void WordHelper::OpenTemplate()
    {
        m_wordApp = gcnew Word::ApplicationClass();
        m_wordApp->Visible = true;
        m_wordDocs = m_wordApp->Documents;
        Object^ newTemplate = false;
        Object^ documentType = Word::WdNewDocumentType::wdNewBlankDocument;
        Object^ visible = true;
        try
        {
            m_wordDoc = m_wordDocs->Add(m_templateFile, newTemplate, documentType, visible);
        }
        catch (Exception^ e)
        {
            MessageBox::Show("Не удалось открыть файл шаблона!\nПроверьте заданный путь в настройках.\nТекущий путь:\n" + m_templateFile + "\n" + e->Message, "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);

            return;
        }
        // Открытие исходного шаблона, а нужно новый документ, созданный по шаблону
        //Object^ confirmConversions = true;
        //Object^ readOnly = false;
        //Object^ addToRecentFiles = true;
        //Object^ passwordDocument = missing;
        //Object^ passwordTemplate = missing;
        //Object^ revert = false;
        //Object^ writePasswordDocument = missing;
        //Object^ writePasswordTemplate = missing;
        //Object^ format = missing;
        //Object^ encoding = missing;;
        //Object^ oVisible = missing;
        //Object^ openConflictDocument = missing;
        //Object^ openAndRepair = missing;
        //Object^ documentDirection = missing;
        //Object^ noEncodingDialog = false;
        //Object^ xmlTransform = missing;
        //m_wordDoc = m_wordDocs->Open(m_templateFile,
        //    confirmConversions,     // ConfirmConversions
        //    readOnly,     // ReadOnly
        //    addToRecentFiles,     // AddToRecentFiles
        //    passwordDocument,     // PasswordDocument
        //    passwordTemplate,     // PasswordTemplate
        //    revert,     // Revert
        //    writePasswordDocument,     // WritePasswordDocument
        //    writePasswordTemplate,     // WritePasswordTemplate
        //    format,     // Format
        //    encoding,     // Encoding
        //    oVisible,     // Visible
        //    openAndRepair,     // OpenAndRepair
        //    documentDirection,     // DocumentDirection
        //    noEncodingDialog,     // NoEncodingDialog
        //    xmlTransform);
    }
    
    void WordHelper::InsertEpicrisToTemplate()
    {
        if (m_wordDoc == nullptr)
            return;
        Object^ matchCase = true;
        Object^ matchWholeWord = true;
        Object^ matchWildcards = false;
        Object^ matchSoundsLike = false;
        Object^ matchAllWordForms = false;
        Object^ forward = true;
        Object^ wrap = Word::WdFindWrap::wdFindContinue;    
        Object^ format = false;
        Object^ replaceAll = Word::WdReplace::wdReplaceAll;
        Object^ matchKashida = missing;
        Object^ matchDiacritics = missing;
        Object^ matchAlefHamza = missing;
        Object^ matchControl = missing;
        
        Object^ bmHistoryNum = (Object^)"НомерБолезни";
        Object^ bmHistoryYear = (Object^)"ГодБолезни";
        Object^ bmAnamnes = (Object^)"Анамнез";
        Object^ bmVVK = (Object^)"ВВК";
        Object^ bmOutcomeDate = (Object^)"ДатаВыписки";
        Object^ bmIncomeDate = (Object^)"ДатаПоступления";
        Object^ bmBirthday = (Object^)"ДатаРождения";
        Object^ bmSideInfo = (Object^)"Дополнительно";
        Object^ bmRank = (Object^)"Звание";
        Object^ bmName = (Object^)"Имя";
        Object^ bmTherapy = (Object^)"Лечение";
        Object^ bmMilitaryUnit = (Object^)"НомерЧасти";
        Object^ bmSurgery = (Object^)"Операции";
        Object^ bmComplications = (Object^)"Осложнения";
        Object^ bmDiagnosis = (Object^)"ОсновнойДиагноз";
        Object^ bmPatronymic = (Object^)"Отчество";
        Object^ bmAnalyzes = (Object^)"РезультатыАнализов";
        Object^ bmRecommendations = (Object^)"Рекомендации";
        Object^ bmRelatedDiagnosis = (Object^)"СопутствующиеЗаболевания";
        Object^ bmSurname = (Object^)"Фамилия";

        Word::Range^ range;
        Word::Bookmark^ bookmark;
        // Номер истории болезни
        if (m_wordDoc->Bookmarks->Exists((String^)bmHistoryNum)) {
            bookmark = m_wordDoc->Bookmarks[bmHistoryNum];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->HistoryNumber);
        }        
        /*Object^ findText = "{{НомерБолезни}}";
        Object^ replaceWith = m_epicris->HistoryNumber;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Год болезни
        if (m_wordDoc->Bookmarks->Exists((String^)bmHistoryYear)) {
            bookmark = m_wordDoc->Bookmarks[bmHistoryYear];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->HistoryYear);
        }
        /*findText = "{{ГодБолезни}}";
        replaceWith = m_epicris->HistoryYear;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Имя
        if (m_wordDoc->Bookmarks->Exists((String^)bmName)) {
            bookmark = m_wordDoc->Bookmarks[bmName];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->Name);
        }
        /*findText = "{{Имя}}";
        replaceWith = m_epicris->Name;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike,matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Фамилия
        if (m_wordDoc->Bookmarks->Exists((String^)bmSurname)) {
            bookmark = m_wordDoc->Bookmarks[bmSurname];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->Surname);
        }
        /*findText = "{{Фамилия}}";
        replaceWith = m_epicris->Surname;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Отчество
        if (m_wordDoc->Bookmarks->Exists((String^)bmPatronymic)) {
            bookmark = m_wordDoc->Bookmarks[bmPatronymic];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->Patronymic);
        }
        /*findText = "{{Отчество}}";
        replaceWith = m_epicris->Patronymic;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Дата поступления
        if (m_wordDoc->Bookmarks->Exists((String^)bmIncomeDate)) {
            bookmark = m_wordDoc->Bookmarks[bmIncomeDate];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->IncomeDate);
        }
        /*findText = "{{ДатаПоступления}}";
        replaceWith = m_epicris->IncomeDate;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Дата выписки
        if (m_wordDoc->Bookmarks->Exists((String^)bmOutcomeDate)) {
            bookmark = m_wordDoc->Bookmarks[bmOutcomeDate];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->OutcomeDate);
        }
        /*findText = "{{ДатаВыписки}}";
        replaceWith = m_epicris->OutcomeDate;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Дата рождения
        if (m_wordDoc->Bookmarks->Exists((String^)bmBirthday)) {
            bookmark = m_wordDoc->Bookmarks[bmBirthday];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->Birthday);
        }
        /*findText = "{{ДатаРождения}}";
        replaceWith = m_epicris->Birthday;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Звание
        if (m_wordDoc->Bookmarks->Exists((String^)bmRank)) {
            bookmark = m_wordDoc->Bookmarks[bmRank];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->Rank);
        }
        /*findText = "{{Звание}}";
        replaceWith = m_epicris->Rank;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Номер части
        if (m_wordDoc->Bookmarks->Exists((String^)bmMilitaryUnit)) {
            bookmark = m_wordDoc->Bookmarks[bmMilitaryUnit];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->MilitaryUnit);
        }
        /*findText = "{{НомерЧасти}}";
        replaceWith = m_epicris->MilitaryUnit;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Основной диагноз
        if (m_wordDoc->Bookmarks->Exists((String^)bmDiagnosis)) {
            bookmark = m_wordDoc->Bookmarks[bmDiagnosis];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->Diagnosis);
        }
        /*findText = "{{ОснДиагноз}}";
        replaceWith = m_epicris->Diagnosis;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Осложнения
        if (m_wordDoc->Bookmarks->Exists((String^)bmComplications)) {
            bookmark = m_wordDoc->Bookmarks[bmComplications];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->Complications);
        }
        /*findText = "{{Осложнения}}";
        replaceWith = m_epicris->Complications;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Сопутствующие заболевания
        if (m_wordDoc->Bookmarks->Exists((String^)bmRelatedDiagnosis)) {
            bookmark = m_wordDoc->Bookmarks[bmRelatedDiagnosis];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->RelatedDiagnosis);
        }
        /*findText = "{{СопутЗабол}}";
        replaceWith = m_epicris->RelatedDiagnosis;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Анамнез
        if (m_wordDoc->Bookmarks->Exists((String^)bmAnamnes)) {
            bookmark = m_wordDoc->Bookmarks[bmAnamnes];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->AnamnesisText);
        }
        /*findText = "{{Анамнез}}";
        replaceWith = m_epicris->AnamnesisText;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Лечение
        if (m_wordDoc->Bookmarks->Exists((String^)bmTherapy)) {
            bookmark = m_wordDoc->Bookmarks[bmTherapy];
            range = bookmark->Range;
            range->Text = String::Join(", ", m_epicris->Therapy);
        }
        /*findText = "{{Лечение}}";       
        replaceWith = String::Join(", ", m_epicris->Therapy);
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Дополнительно
        if (m_wordDoc->Bookmarks->Exists((String^)bmSideInfo)) {
            bookmark = m_wordDoc->Bookmarks[bmSideInfo];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->SideData);
        }
        /*findText = "{{Дополнительно}}";
        replaceWith = m_epicris->SideData;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/
        // Рекомендации
        if (m_wordDoc->Bookmarks->Exists((String^)bmRecommendations)) {
            bookmark = m_wordDoc->Bookmarks[bmRecommendations];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->Recommendations);
        }
        /*findText = "{{Рекомендации}}";
        replaceWith = m_epicris->Recommendations;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);*/ 
        // ВВК
        if (m_wordDoc->Bookmarks->Exists((String^)bmVVK)) {
            bookmark = m_wordDoc->Bookmarks[bmVVK];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->VVK);
        }
    }

    void WordHelper::CloseTemplate()
    {
        Object^ saveChanges = Word::WdSaveOptions::wdPromptToSaveChanges;
        Object^ originalFormat = Word::WdOriginalFormat::wdWordDocument;
        Object^ routeDocument = missing;
        ((Word::_Application^)m_wordApp)->Quit(saveChanges, originalFormat, routeDocument);
        m_wordApp = nullptr;
    }

    void WordHelper::InsertFirstListToTemplate()
    {
        throw gcnew System::NotImplementedException();
    }

    void WordHelper::SaveTemplate()
    {
        if (m_wordDoc == nullptr)
            return;
        String^ diagnosis;
        if (m_epicris->Diagnosis->ToLower()->Contains("пневмония"))
            diagnosis = "пневмония ";
        else if (m_epicris->Diagnosis->ToLower()->Contains("бронхит"))
            diagnosis = "бронхит ";
        Object^ fileName = m_outputDir + "\\" + m_epicris->Surname + " эпикриз " + diagnosis + DateTime::Parse(m_epicris->OutcomeDate).AddDays(-1).ToShortDateString() + ".doc";
        Object^ fileFormat = Word::WdSaveFormat::wdFormatDocument;
        Object^ lockComments = false;
        Object^ password = "";
        Object^ addToRecentFiles = false;
        Object^ writePassword = "";
        Object^ readOnlyRecommended = false;
        Object^ embedTrueTypeFonts = false;
        Object^ saveNativePictureFormat = false;
        Object^ saveFormsData = false;
        Object^ saveAsAOCELetter = missing;
        Object^ encoding = missing;
        Object^ insertLineBreaks = missing;
        Object^ allowSubstitutions = missing;
        Object^ lineEnding = missing;
        Object^ addBiDiMarks = missing;
        m_wordDoc->SaveAs(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts,
            saveNativePictureFormat,saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks);
    }

    void WordHelper::InsertDoctorsRecords()
    {
        throw gcnew System::NotImplementedException();
    }
    void WordHelper::InsertTable(Table^ table, Word::Range^% range) {
        int numRows = table->children->Count;
        int numColumns = table->columns->Count;
        Object^ defaultTableBehavior = Word::WdDefaultTableBehavior::wdWord9TableBehavior;
        Object^ autoFitBehavior = Word::WdAutoFitBehavior::wdAutoFitWindow;
        Word::Table^ wordTable = m_wordDoc->Tables->Add(range, numRows, numColumns, defaultTableBehavior, autoFitBehavior);
        int i = 1;
        for each (TableRow ^ row in table->children) {
            int j = 1;
            for each(TableCell^ cell in row->children) {
                if (cell->paragraphs != nullptr) {
                    for each (Paragraph ^ para in cell->paragraphs) {
                        for each (Child ^ child in para->children) {
                            Word::Cell^ wordCell = wordTable->Cell(i, j);
                            wordCell->VerticalAlignment = Word::WdCellVerticalAlignment::wdCellAlignVerticalCenter;
                            wordCell->Range->ParagraphFormat->Alignment = Word::WdParagraphAlignment::wdAlignParagraphCenter;
                            wordCell->TopPadding = 1;
                            wordCell->RightPadding = 1;
                            wordCell->BottomPadding = 1;
                            wordCell->LeftPadding = 1;
                            GetChildFormatting(wordCell, child);
                        }
                    }
                }
                else {
                    for each (Child ^ child in cell->children) {
                        Word::Cell^ wordCell = wordTable->Cell(i, j);
                        wordCell->VerticalAlignment = Word::WdCellVerticalAlignment::wdCellAlignVerticalCenter;
                        wordCell->Range->ParagraphFormat->Alignment = Word::WdParagraphAlignment::wdAlignParagraphCenter;
                        wordCell->TopPadding = 1;
                        wordCell->RightPadding = 1;
                        wordCell->BottomPadding = 1;
                        wordCell->LeftPadding = 1;
                        GetChildFormatting(wordCell, child);
                    }
                }
                j++;
            }
            i++;
        }
    }
    Word::Range^ WordHelper::GetChildFormatting(Word::Cell^ cell, Child^ child) {
        Word::Range^ range = cell->Range;        
        if (child->bold.HasValue)           
            range->Bold = Convert::ToInt32(child->bold.Value);
        if (child->underline.HasValue)
            if (child->underline.Value)
                range->Underline = Word::WdUnderline::wdUnderlineSingle;
        range->Font->Size = (float)child->fontSize.Value;
        range->Text = child->text;
        return range;
    }
    Word::Range^ WordHelper::GetChildFormatting(Word::Range^% range, Child^ child) {
        if (child->bold.HasValue)
            range->Bold = Convert::ToInt32(child->bold.Value);
        if (child->underline.HasValue)
            if (child->underline.Value)
                range->Underline = Word::WdUnderline::wdUnderlineSingle;
        range->Font->Size = (float)child->fontSize.Value;
        range->Text = child->text;
        return range;
    }
    void WordHelper::InsertParagraph(Paragraph^ paragraph, Word::Range^% range) {
        if (paragraph->align == "center")
            range->ParagraphFormat->Alignment = Word::WdParagraphAlignment::wdAlignParagraphCenter;
        if (paragraph->align == "right")
            range->ParagraphFormat->Alignment = Word::WdParagraphAlignment::wdAlignParagraphRight;
        if (paragraph->align == "left")
            range->ParagraphFormat->Alignment = Word::WdParagraphAlignment::wdAlignParagraphLeft;
        if (paragraph->align == "justify")
            range->ParagraphFormat->Alignment = Word::WdParagraphAlignment::wdAlignParagraphJustify;
        for each (Child ^ child in paragraph->children) {
            range = GetChildFormatting(range, child);
        }
    }
    void WordHelper::InsertAnalyzes() {
        Word::Application^ wordApp = gcnew Word::ApplicationClass();
        wordApp->Visible = true;
        Object^ newTemplate = false;
        Object^ documentType = Word::WdNewDocumentType::wdNewBlankDocument;
        Object^ visible = true;
        Object^ filePath = R"(C:\Users\user\Desktop\newdoc.docx)";
        Word::Document^ doc = wordApp->Documents->Add(filePath, newTemplate, documentType, visible);
        Object^ begin = (Object^)0;
        Object^ end = (Object^)0;
        Object^ defaultTableBehavior = Word::WdDefaultTableBehavior::wdWord9TableBehavior;
        Object^ autoFitBehavior = Word::WdAutoFitBehavior::wdAutoFitWindow;
        List<JObject^>^ items = JsonConvert::DeserializeObject<List<JObject^>^>(m_epicris->AnalyzesListJson);       
        SortedList<int, JToken^>^ sortedItems = gcnew SortedList<int, JToken^>();        
        Dictionary<int, JToken^>^ dict = gcnew Dictionary<int, JToken^>();
        for each (JObject ^ item in items) {
            dict->Add(Convert::ToInt32(item["position"]), item["value"]);
        }
        sortedItems = gcnew SortedList<int, JToken^>(dict,nullptr);
        
        /*for each (Object ^ item in parser->DeserializedItems) {
            if (dynamic_cast<Paragraph^>(item)) {
                InsertParagraph((Paragraph^)item);
            }
            else if (dynamic_cast<Table^>(item)) {
                InsertTable((Table^));
            }
        }*/
    }
    void WordHelper::InsertAnalyzes(String^ analyzes) {
        Word::Application^ wordApp = gcnew Word::ApplicationClass();
        wordApp->Visible = true;
        Object^ newTemplate = false;
        Object^ documentType = Word::WdNewDocumentType::wdNewBlankDocument;
        Object^ visible = true;
        Object^ filePath = R"(C:\Users\user\Desktop\newdoc.docx)";
        m_wordDoc = wordApp->Documents->Add(filePath, newTemplate, documentType, visible);
        Object^ defaultTableBehavior = Word::WdDefaultTableBehavior::wdWord9TableBehavior;
        Object^ autoFitBehavior = Word::WdAutoFitBehavior::wdAutoFitWindow;
        List<JObject^>^ items = JsonConvert::DeserializeObject<List<JObject^>^>(analyzes);
        SortedList<int, JToken^>^ sortedItems = gcnew SortedList<int, JToken^>();
        Dictionary<int, JToken^>^ dict = gcnew Dictionary<int, JToken^>();
        for each (JObject ^ item in items) {
            dict->Add(Convert::ToInt32(item["position"]), item["value"]);
        }
        sortedItems = gcnew SortedList<int, JToken^>(dict, nullptr);
        List<JToken^>^ sortedList = gcnew List<JToken^>();
        for each (JToken ^ token in sortedItems->Values) {
            for each (JToken ^ item in token) {
                sortedList->Add(item);
            }
        }
        String^ json = JsonConvert::SerializeObject(sortedList);
        RtfDocumentCreator^ rtfCreator = gcnew RtfDocumentCreator();
        rtfCreator->GenerateParser(json);
        Parser^ parser = rtfCreator->GetParser();
        Object^ bmAnalyzes = (Object^)"Анализы";
        if (!m_wordDoc->Bookmarks->Exists((String^)bmAnalyzes))
            return;
        Word::Bookmark^ bookmark = m_wordDoc->Bookmarks->default[bmAnalyzes];
        Word::Range^ range = bookmark->Range;
        for each (Object ^ item in parser->DeserializedItems) {
            if (dynamic_cast<Paragraph^>(item)) {
                InsertParagraph((Paragraph^)item, range);
                range->InsertParagraphAfter();
                range = range->Paragraphs->Last->Range;
            }
            else if (dynamic_cast<Table^>(item)) {
                InsertTable((Table^)item, range);
                range->InsertParagraphAfter();
                range = range->Tables->default[range->Tables->Count]->Range;
                
            }
        }
    }
    
    

};