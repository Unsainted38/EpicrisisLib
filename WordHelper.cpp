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
        
        Word::Range^ range = m_wordDoc->Content;
        // ����� ������� �������
        Object^ findText = "{{������������}}";
        Object^ replaceWith = m_epicris->HistoryNumber;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // ��� �������
        findText = "{{����������}}";
        replaceWith = m_epicris->HistoryYear;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // ���
        findText = "{{���}}";
        replaceWith = m_epicris->Name;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike,matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // �������
        findText = "{{�������}}";
        replaceWith = m_epicris->Surname;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // ��������
        findText = "{{��������}}";
        replaceWith = m_epicris->Patronymic;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // ���� �����������
        findText = "{{���������������}}";
        replaceWith = m_epicris->IncomeDate;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // ���� �������
        findText = "{{�����������}}";
        replaceWith = m_epicris->OutcomeDate;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // ���� ��������
        findText = "{{������������}}";
        replaceWith = m_epicris->Birthday;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // ������
        findText = "{{������}}";
        replaceWith = m_epicris->Rank;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // ����� �����
        findText = "{{����������}}";
        replaceWith = m_epicris->MilitaryUnit;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // �������� �������
        findText = "{{����������}}";
        replaceWith = m_epicris->Diagnosis;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // ����������
        findText = "{{����������}}";
        replaceWith = m_epicris->Complications;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // ������������� �����������
        findText = "{{����������}}";
        replaceWith = m_epicris->RelatedDiagnosis;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // �������
        findText = "{{�������}}";
        replaceWith = m_epicris->AnamnesisText;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // �������
        findText = "{{�������}}";       
        replaceWith = String::Join(", ", m_epicris->Therapy);
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // �������������
        findText = "{{�������������}}";
        replaceWith = m_epicris->SideData;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);
        // ������������
        findText = "{{������������}}";
        replaceWith = m_epicris->Recommendations;
        range->Find->Execute(findText, matchCase, matchWholeWord, matchWildcards,
            matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replaceAll,
            matchKashida, matchDiacritics, matchAlefHamza, matchControl);      
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
        m_wordDoc->SaveAs(fileName, fileFormat, lockComments, password, addToRecentFiles, writePassword, readOnlyRecommended, embedTrueTypeFonts,
            saveNativePictureFormat,saveFormsData, saveAsAOCELetter, encoding, insertLineBreaks, allowSubstitutions, lineEnding, addBiDiMarks);
    }

    void WordHelper::InsertDoctorsRecords()
    {
        throw gcnew System::NotImplementedException();
    }
    void WordHelper::InsertTable(Table^ table)
    {
        
    }
    void WordHelper::InsertParagraph(Paragraph^ parapraph)
    {
    }
    void WordHelper::InsertAnalyzes()
    {
    }
    void WordHelper::InsertTables(Word::Document^ doc, Dictionary<String^, Object^>^ json) {
        // �������� ������� �������� ��� ������� �������
        if (doc->Bookmarks->Exists("TablePlaceholder")) {
            Word::Bookmark^ bookmark = m_wordDoc->Bookmarks->default[(Object^%)"Tables"];
            
            Word::Range^ range = bookmark->Range;
            

            // ��������� ������ �� JSON
            List<String^>^ columns = safe_cast<List<String^>^>(json["columns"]);
            List<List<Object^>^>^ rows = safe_cast<List<List<Object^>^>^>(json["rows"]);

            // ������� �������
            Word::Table^ table = doc->Tables->Add(range, rows->Count + 1, columns->Count, missing, missing);
            
            //// ���������� ���������� ��������
            for (int col = 0; col < columns->Count; ++col) {
                table->Cell(1, col + 1)->Range->Text = columns[col];
            }
            
            //// ���������� ����� ������
            //for (int row = 0; row < rows->Count; ++row) {
            //    for (int col = 0; col < columns->Count; ++col) {
            //        table->Cell(row + 2, col + 1)->Range->Text = rows[row][col]->ToString();
            //    }
            //}
        }
        else {
            throw gcnew Exception("Bookmark 'TablePlaceholder' not found in the document.");
        }
    }
    

};