#include "pch.h"

//
//  TODO
// ������� ���������� ������ �������� ������
// ������� ������ � ������ ������� �� �������
// ������� ���������� json �� �������� �����
//

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
            MessageBox::Show("�� ������� ������� ���� �������!\n��������� �������� ���� � ����������.\n������� ����:\n" + m_templateFile + "\n" + e->Message, "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);

            return;
        }
        // �������� ��������� �������, � ����� ����� ��������, ��������� �� �������
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
        
        Object^ bmHistoryNum = (Object^)"������������";
        Object^ bmHistoryYear = (Object^)"����������";
        Object^ bmAnamnes = (Object^)"�������";
        Object^ bmVVK = (Object^)"���";
        Object^ bmOutcomeDate = (Object^)"�����������";
        Object^ bmIncomeDate = (Object^)"���������������";
        Object^ bmBirthday = (Object^)"������������";
        Object^ bmSideInfo = (Object^)"�������������";
        Object^ bmRank = (Object^)"������";
        Object^ bmName = (Object^)"���";
        Object^ bmTherapy = (Object^)"�������";
        Object^ bmMilitaryUnit = (Object^)"����������";
        Object^ bmSurgery = (Object^)"��������";
        Object^ bmComplications = (Object^)"����������";
        Object^ bmDiagnosis = (Object^)"���������������";
        Object^ bmPatronymic = (Object^)"��������";
        Object^ bmAnalyzes = (Object^)"������������������";
        Object^ bmRecommendations = (Object^)"������������";
        Object^ bmRelatedDiagnosis = (Object^)"������������������������";
        Object^ bmSurname = (Object^)"�������";
        Object^ bmMkb = (Object^)"���";

        Word::Range^ range;
        Word::Bookmark^ bookmark;
        // ����� ������� �������
        if (m_wordDoc->Bookmarks->Exists((String^)bmHistoryNum)) {
            bookmark = m_wordDoc->Bookmarks[bmHistoryNum];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->HistoryNumber);
        }
        // ��� �������
        if (m_wordDoc->Bookmarks->Exists((String^)bmHistoryYear)) {
            bookmark = m_wordDoc->Bookmarks[bmHistoryYear];
            range = bookmark->Range;
            range->Text = Convert::ToString(m_epicris->HistoryYear);
        }
        // ���
        if (m_wordDoc->Bookmarks->Exists((String^)bmName)) {
            bookmark = m_wordDoc->Bookmarks[bmName];
            range = bookmark->Range;
            range->Text = m_epicris->Name;
        }
        // �������
        if (m_wordDoc->Bookmarks->Exists((String^)bmSurname)) {
            bookmark = m_wordDoc->Bookmarks[bmSurname];
            range = bookmark->Range;
            range->Text = m_epicris->Surname;
        }
        // ��������
        if (m_wordDoc->Bookmarks->Exists((String^)bmPatronymic)) {
            bookmark = m_wordDoc->Bookmarks[bmPatronymic];
            range = bookmark->Range;
            range->Text = m_epicris->Patronymic;
        }
        // ���� �����������
        if (m_wordDoc->Bookmarks->Exists((String^)bmIncomeDate)) {
            bookmark = m_wordDoc->Bookmarks[bmIncomeDate];
            range = bookmark->Range;
            range->Text = m_epicris->IncomeDate;
        }
        // ���� �������
        if (m_wordDoc->Bookmarks->Exists((String^)bmOutcomeDate)) {
            bookmark = m_wordDoc->Bookmarks[bmOutcomeDate];
            range = bookmark->Range;
            range->Text = m_epicris->OutcomeDate;
        }
        // ���� ��������
        if (m_wordDoc->Bookmarks->Exists((String^)bmBirthday)) {
            bookmark = m_wordDoc->Bookmarks[bmBirthday];
            range = bookmark->Range;
            range->Text = m_epicris->Birthday;
        }
        // ������
        if (m_wordDoc->Bookmarks->Exists((String^)bmRank)) {
            bookmark = m_wordDoc->Bookmarks[bmRank];
            range = bookmark->Range;
            range->Text = m_epicris->Rank->ToLower();
        }
        // ����� �����
        if (m_wordDoc->Bookmarks->Exists((String^)bmMilitaryUnit)) {
            bookmark = m_wordDoc->Bookmarks[bmMilitaryUnit];
            range = bookmark->Range;
            range->Text = m_epicris->MilitaryUnit;
        }
        // �������� �������
        if (m_wordDoc->Bookmarks->Exists((String^)bmDiagnosis)) {
            bookmark = m_wordDoc->Bookmarks[bmDiagnosis];
            range = bookmark->Range;
            range->Text = m_epicris->Diagnosis;
        }
        // ���
        if (m_wordDoc->Bookmarks->Exists((String^)bmMkb)) {
            bookmark = m_wordDoc->Bookmarks[bmMkb];
            range = bookmark->Range;
            range->Text = m_epicris->Mkb;
        }
        // ����������
        if (m_wordDoc->Bookmarks->Exists((String^)bmComplications)) {
            bookmark = m_wordDoc->Bookmarks[bmComplications];
            range = bookmark->Range;
            range->Text = m_epicris->Complications;
        }
        // ������������� �����������
        if (m_wordDoc->Bookmarks->Exists((String^)bmRelatedDiagnosis)) {
            bookmark = m_wordDoc->Bookmarks[bmRelatedDiagnosis];
            range = bookmark->Range;
            range->Text = m_epicris->RelatedDiagnosis;
        }
        // �������
        if (m_wordDoc->Bookmarks->Exists((String^)bmAnamnes)) {
            bookmark = m_wordDoc->Bookmarks[bmAnamnes];
            range = bookmark->Range;
            range->Text = m_epicris->AnamnesisText;
        }
        // �������
        if (m_wordDoc->Bookmarks->Exists((String^)bmAnalyzes)) {
            InsertAnalyzes(bmAnalyzes);
        }
        // �������
        if (m_wordDoc->Bookmarks->Exists((String^)bmTherapy)) {
            bookmark = m_wordDoc->Bookmarks[bmTherapy];
            range = bookmark->Range;
            range->Text = String::Join(", ", m_epicris->Therapy);
        }
        // �������������
        if (m_wordDoc->Bookmarks->Exists((String^)bmSideInfo)) {
            bookmark = m_wordDoc->Bookmarks[bmSideInfo];
            range = bookmark->Range;
            range->Text = m_epicris->SideData;
        }
        // ������������
        if (m_wordDoc->Bookmarks->Exists((String^)bmRecommendations)) {
            bookmark = m_wordDoc->Bookmarks[bmRecommendations];
            range = bookmark->Range;
            range->Text = m_epicris->Recommendations;
        } 
        // ���
        if (m_wordDoc->Bookmarks->Exists((String^)bmVVK)) {
            bookmark = m_wordDoc->Bookmarks[bmVVK];
            range = bookmark->Range;
            range->Text = m_epicris->VVK;
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
        String^ diagnosis = "";
        if (m_epicris->Diagnosis->ToLower()->Contains("���������"))
            diagnosis = "��������� ";
        else if (m_epicris->Diagnosis->ToLower()->Contains("�������"))
            diagnosis = "������� ";
        Object^ fileName = m_outputDir + "\\" + m_epicris->Surname + " ������� " + diagnosis + DateTime::Parse(m_epicris->OutcomeDate).AddDays(-1).ToShortDateString() + ".doc";
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
        try {
            m_wordDoc->SaveAs(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts,
                saveNativePictureFormat, saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks);
        }
        catch(Exception^ ex) {
            MessageBox::Show(ex->Message);
            m_wordDoc->Save();
        }        
    }

    void WordHelper::InsertDoctorsRecords()
    {
        throw gcnew System::NotImplementedException();
    }
    void WordHelper::InsertTable(Table^ table, Word::Range^% range) {
        table = DeleteEmptyColumns(table);
        int numRows = table->children->Count;
        int numColumns = table->columns->Count;
        Object^ defaultTableBehavior = Word::WdDefaultTableBehavior::wdWord9TableBehavior;
        Object^ autoFitBehavior = Word::WdAutoFitBehavior::wdAutoFitWindow;
        //range->ParagraphFormat->LeftIndent = range->PageSetup->LeftMargin / 28.3465f;
        Word::Table^ wordTable = m_wordDoc->Tables->Add(range, numRows, numColumns, defaultTableBehavior, autoFitBehavior);
        wordTable->Range->Cells->VerticalAlignment = Word::WdCellVerticalAlignment::wdCellAlignVerticalCenter;
        //wordTable->Range->Cells->AutoFit();
        wordTable->BottomPadding = 0;
        wordTable->LeftPadding = 1;
        wordTable->TopPadding = 0;
        wordTable->RightPadding = 1;
        wordTable->Range->ParagraphFormat->Alignment = Word::WdParagraphAlignment::wdAlignParagraphCenter;
        int i = 1;
        for each (TableRow ^ row in table->children) {
            int j = 1;            
            for each(TableCell^ cell in row->children) {
                Word::Cell^ wordCell = wordTable->Cell(i, j);
                Word::Range^ CellRange = wordCell->Range;
                /*wordCell->TopPadding = 0;
                wordCell->RightPadding = 1;
                wordCell->BottomPadding = 0;
                wordCell->LeftPadding = 1;*/
                if (cell->paragraphs != nullptr) {
                    int paraCount = 1;
                    for each (Paragraph ^ para in cell->paragraphs) {                        
                        Word::Paragraph^ wordPara = CellRange->Paragraphs[paraCount];
                        Word::Range^ paraRange = wordPara->Range;
                        for each (Child ^ child in para->children) {                                                        
                            GetChildFormatting(paraRange, child);
                        }
                        if (paraCount != 1)
                            CellRange->InsertParagraphAfter();
                        paraCount++;
                    }
                }
                else {
                    for each (Child ^ child in cell->children) {
                        GetChildFormatting(wordCell, child);
                    }
                }
                j++;
            }
            i++;
        }
        range->SetRange(wordTable->Range->End, wordTable->Range->End);
        /*range = wordTable->Range;
        range->InsertParagraphAfter();        
        Object^ unit = Word::WdUnits::wdParagraph;
        Object^ count = 1;
        range = range->Next(unit, count);*/
    }
    Word::Range^ WordHelper::GetChildFormatting(Word::Cell^ cell, Child^ child) {
        Word::Range^ range = cell->Range;
        range->InsertAfter(child->text);
        if (child->bold.HasValue)           
            range->Bold = child->bold ? 1 : 0;
        else
            range->Font->Bold = 0;
        if (child->underline.HasValue)
            range->Underline = child->underline ? Word::WdUnderline::wdUnderlineSingle : Word::WdUnderline::wdUnderlineNone;
        else
            range->Font->Underline = Word::WdUnderline::wdUnderlineNone;
        range->Font->Size = (float)child->fontSize.Value;
        Object^ collapseDirection = Word::WdCollapseDirection::wdCollapseEnd;
        range->Collapse(collapseDirection);
        return range;
    }
    Word::Range^ WordHelper::GetChildFormatting(Word::Range^ %range, Child^ child) {
        range->InsertAfter(child->text);
        if (child->bold.HasValue)
            range->Font->Bold = child->bold ? 1 : 0;
        else
            range->Font->Bold = 0;
        if (child->underline.HasValue)
            range->Font->Underline = child->underline ? Word::WdUnderline::wdUnderlineSingle : Word::WdUnderline::wdUnderlineNone;
        else
            range->Font->Underline = Word::WdUnderline::wdUnderlineNone;
        range->Font->Size = (float)child->fontSize.Value;
        Object^ collapseDirection = Word::WdCollapseDirection::wdCollapseEnd;
        range->Collapse(collapseDirection);
        return range;
    }
    Table^ WordHelper::DeleteEmptyColumns(Table^ table)
    {
        if (table->columns->Count == 0 || table->children->Count == 0)
            return table;

        // ������� �������, ������� ����� �������
        List<int>^ emptyColumnIndices = gcnew List<int>();

        // �������� �� ������� �������
        for (int col = 0; col < table->columns->Count; col++)
        {
            bool isEmpty = true;

            // �������� �� ������ ������, ������� �� ������ (������ 1)
            for (int row = 1; row < table->children->Count; row++)
            {
                TableRow^ tableRow = table->children[row];
                TableCell^ cell = tableRow->children[col];

                // ���������, ���� �� ����� � ������
                if (cell->paragraphs->Count > 0)
                {
                    for each (Paragraph ^ paragraph in cell->paragraphs)
                    {
                        if (paragraph->children->Count > 0)
                        {                           
                            for each (Child ^ child in paragraph->children) {
                                if (!String::IsNullOrEmpty(child->text)) {
                                    isEmpty = false;
                                    break;
                                }
                            }
                            if (!isEmpty)
                                break;
                        }
                    }
                }
                else {
                    for each (Child ^ child in cell->children) {
                        if (!String::IsNullOrEmpty(child->text)) {
                            isEmpty = false;
                            break;
                        }
                    }
                }

                if (!isEmpty)
                    break; // ���������� �������� ���� ����� ����� � ������
            }

            // ���� ������� ������, ��������� ��� ������ � ������ ��� ��������
            if (isEmpty)
            {
                emptyColumnIndices->Add(col);
            }
        }

        // ������� ������ �������, ������� � �����, ����� �� �������� �������
        for (int i = emptyColumnIndices->Count - 1; i >= 0; i--)
        {
            int colIndex = emptyColumnIndices[i];

            // ������� �������
            table->columns->RemoveAt(colIndex);

            // ������� ��������������� ������ � ������ ������
            for each (TableRow ^ row in table->children)
            {
                row->children->RemoveAt(colIndex);
            }
        }
        return table;
    }
    void WordHelper::InsertParagraph(Paragraph^ paragraph, Word::Range^% range) {
        //range->ParagraphFormat->LeftIndent = m_wordDoc->PageSetup->LeftMargin / 28.3465f;
        //range->ParagraphFormat->FirstLineIndent = m_wordDoc->PageSetup->LeftMargin / 28.3465f;
        if (paragraph->align == "center")
            range->ParagraphFormat->Alignment = Word::WdParagraphAlignment::wdAlignParagraphCenter;
        else if (paragraph->align == "right")
            range->ParagraphFormat->Alignment = Word::WdParagraphAlignment::wdAlignParagraphRight;
        else if (paragraph->align == "left")
            range->ParagraphFormat->Alignment = Word::WdParagraphAlignment::wdAlignParagraphLeft;
        else if (paragraph->align == "justify")
            range->ParagraphFormat->Alignment = Word::WdParagraphAlignment::wdAlignParagraphJustify;
        bool flag = false;
        for each (Child ^ child in paragraph->children) {
            range = GetChildFormatting(range, child);
            if (child->text->Contains("�-���������� �����"))
                flag = true;
        }
        if (flag) {
            range->SetRange(range->End, range->End);
            range->InsertAfter(" ");
            flag = false;
            return;
        }            
        range->InsertParagraphAfter();
        range->SetRange(range->End, range->End);
        //range->ParagraphFormat->LeftIndent = range->PageSetup->LeftMargin / 28.3465f;
        /*range->InsertParagraphAfter();
        Object^ unit = Word::WdUnits::wdParagraph;
        Object^ count = 1;
        range = range->Next(unit, count); */
    }
    void WordHelper::InsertAnalyzes(Object^ bmAnalyzes) {
        /*Word::Application^ wordApp = gcnew Word::ApplicationClass();
        wordApp->Visible = true;
        Object^ newTemplate = false;
        Object^ documentType = Word::WdNewDocumentType::wdNewBlankDocument;
        Object^ visible = true;
        Object^ filePath = R"(C:\Users\user\Desktop\newdoc.docx)";
        m_wordDoc = wordApp->Documents->Add(filePath, newTemplate, documentType, visible);
        Object^ defaultTableBehavior = Word::WdDefaultTableBehavior::wdWord9TableBehavior;
        Object^ autoFitBehavior = Word::WdAutoFitBehavior::wdAutoFitWindow;*/
        // ������ ���������� json �� �������� position
        String^ analyzes = m_epicris->AnalyzesListJson;
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
        if (!m_wordDoc->Bookmarks->Exists((String^)bmAnalyzes))
            return;
        Word::Bookmark^ bookmark = m_wordDoc->Bookmarks->default[bmAnalyzes];
        Word::Range^ range = bookmark->Range;
        range->Text = "";
        //range->ParagraphFormat->LeftIndent = m_wordDoc->PageSetup->LeftMargin / 28.3465f;
        //range->ParagraphFormat->FirstLineIndent = m_wordDoc->PageSetup->LeftMargin / 28.3465f;
        m_wordApp->ScreenUpdating = false;
        for each (Object ^ item in parser->DeserializedItems) {
            if (dynamic_cast<Paragraph^>(item)) {
                InsertParagraph((Paragraph^)item, range);
            }
            else if (dynamic_cast<Table^>(item)) {
                InsertTable((Table^)item, range);                
            }
        }
        m_wordApp->ScreenUpdating = true;
    }
    
    

};