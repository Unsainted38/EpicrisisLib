#include "pch.h"
namespace unsaintedWinAppLib {
    RtfDocumentParser::RtfDocumentParser()
    {
        
    }
    RtfDocumentParser::RtfDocumentParser(String^ rtfString) {
        array<Byte>^ rtfBytes = Encoding::UTF8->GetBytes(rtfString);
        auto stream = gcnew MemoryStream(rtfBytes);
        m_doc = gcnew Aspose::Words::Document(stream);
        ParseRtfDocument();
        CleanParser();
        GenerateJsonDocument();
    }

    void RtfDocumentParser::ResetRtfDocumentParser(String^ rtfString) {
        array<Byte>^ rtfBytes = Encoding::UTF8->GetBytes(rtfString);
        auto stream = gcnew MemoryStream(rtfBytes);
        m_doc = gcnew Aspose::Words::Document(stream);
        ParseRtfDocument();
        CleanParser();
        GenerateJsonDocument();
    }

    Parser^ RtfDocumentParser::GetParser() {
        return m_parser;
    }

    String^ RtfDocumentParser::GetJsonDocument() {
        return m_jsonDocument;
    }

    void RtfDocumentParser::ParseRtfDocument() {
        m_parser = gcnew Parser();
        NodeCollection^ nodes = m_doc->GetChildNodes(NodeType::Any, true);
        for each (Node ^ node in nodes) {
            if (node->NodeType == NodeType::Table) {
                ParseTable(node);
            }
            else if (node->NodeType == NodeType::Paragraph) {
                ParseParagraph(node);
            }
        }
    }

    void RtfDocumentParser::ParseTable(Node^ node) {
        Table^ parsed_table = gcnew Table();
        Aspose::Words::Tables::Table^ table = safe_cast<Aspose::Words::Tables::Table^>(node);

        for each (Aspose::Words::Tables::Row ^ row in table->Rows) {
            TableRow^ parsed_row = gcnew TableRow();
            for each (Aspose::Words::Tables::Cell ^ cell in row->Cells) {
                TableCell^ parsed_cell = gcnew TableCell();
                if (row->IsFirstRow) {
                    Column^ parsed_column = gcnew Column();
                    parsed_cell->type = "tableHeaderCell";
                    parsed_cell->children = gcnew List<Child^>();
                    String^ cellText = cell->GetText()->Trim()->Replace("\a", "");
                    DateTime dateValue;
                    parsed_column->title = cellText;
                    if (cellText->ToLower()->Contains("дата")) {
                        parsed_column->type = "date";
                        DateTime::TryParse(cellText, dateValue);
                        for each (Run ^ run in cell->GetChildNodes(NodeType::Run, true)) {
                            Child^ child = gcnew Child();
                            if (run->Font->Bold)
                                child->bold = true;
                            if (run->Font->Underline == Underline::Single)
                                child->underline = true;
                            child->fontSize = Convert::ToInt16(System::Math::Round(run->Font->Size, MidpointRounding::AwayFromZero));
                            child->text = run->Text->Replace("\a", "");
                            parsed_cell->children->Add(child);
                        }
                        if (cell->GetChildNodes(NodeType::Run, true)->Count == 0) {
                            Child^ child = gcnew Child();
                            child->text = "";
                            parsed_cell->children->Add(child);
                        }
                    }
                    else {
                        parsed_column->type = "text";
                        for each (Run ^ run in cell->GetChildNodes(NodeType::Run, true)) {
                            Child^ child = gcnew Child();
                            if (run->Font->Bold)
                                child->bold = true;
                            if (run->Font->Underline == Underline::Single)
                                child->underline = true;
                            child->fontSize = Convert::ToInt16(System::Math::Round(run->Font->Size, MidpointRounding::AwayFromZero));
                            child->text = run->Text->Replace("\a", "");
                            parsed_cell->children->Add(child);
                        }
                        if (cell->GetChildNodes(NodeType::Run, true)->Count == 0) {
                            Child^ child = gcnew Child();
                            child->text = "";
                            parsed_cell->children->Add(child);
                        }
                    }
                    parsed_table->columns->Add(parsed_column);
                }
                else {
                    parsed_cell->type = "tableDataCell";
                    parsed_cell->paragraphs = gcnew List<Paragraph^>();
                    Paragraph^ parsed_parapraph = gcnew Paragraph();
                    String^ cellText = cell->GetText()->Trim()->Replace("\a", "");
                    DateTime dateValue;
                    if (DateTime::TryParse(cellText, dateValue)) {
                        parsed_cell->columnType = "date";
                        parsed_parapraph->type = "dateInput";
                        for each (Run ^ run in cell->GetChildNodes(NodeType::Run, true)) {
                            Child^ child = gcnew Child();
                            if (run->Font->Bold)
                                child->bold = true;
                            if (run->Font->Underline == Underline::Single)
                                child->underline = true;
                            child->fontSize = Convert::ToInt16(System::Math::Round(run->Font->Size, MidpointRounding::AwayFromZero));
                            child->text = dateValue.ToShortDateString();
                            parsed_parapraph->children->Add(child);
                        }
                        if (cell->GetChildNodes(NodeType::Run, true)->Count == 0) {
                            Child^ child = gcnew Child();
                            child->text = "";
                            parsed_parapraph->children->Add(child);
                        }
                    }
                    else {
                        parsed_cell->columnType = "text";
                        parsed_parapraph->type = "paragraph";
                        for each (Run ^ run in cell->GetChildNodes(NodeType::Run, true)) {
                            Child^ child = gcnew Child();
                            if (run->Font->Bold)
                                child->bold = true;
                            if (run->Font->Underline == Underline::Single)
                                child->underline = true;
                            child->fontSize = Convert::ToInt16(System::Math::Round(run->Font->Size, MidpointRounding::AwayFromZero));
                            child->text = run->Text->Replace("\a", "");
                            parsed_parapraph->children->Add(child);
                        }
                        if (cell->GetChildNodes(NodeType::Run, true)->Count == 0) {
                            Child^ child = gcnew Child();
                            child->text = "";
                            parsed_parapraph->children->Add(child);
                        }
                    }
                    parsed_cell->paragraphs->Add(parsed_parapraph);

                }
                parsed_row->children->Add(parsed_cell);

            }
            parsed_table->children->Add(parsed_row);
        }
        m_parser->DeserializedItems->Add(parsed_table);
    }

    void RtfDocumentParser::ParseParagraph(Node^ node) {
        Paragraph^ parsed_paragraph = gcnew Paragraph();
        Aspose::Words::Paragraph^ para = safe_cast<Aspose::Words::Paragraph^>(node);

        if (!para->IsInCell && !String::IsNullOrWhiteSpace(para->GetText())) {
            Paragraph^ parsed_paragraph = gcnew Paragraph();
            parsed_paragraph->type = "paragraph";
            if (para->ParagraphFormat->Alignment == ParagraphAlignment::Center)
                parsed_paragraph->align = "center";
            if (para->ParagraphFormat->Alignment == ParagraphAlignment::Left)
                parsed_paragraph->align = "left";
            if (para->ParagraphFormat->Alignment == ParagraphAlignment::Right)
                parsed_paragraph->align = "right";
            for each (Run ^ run in para->GetChildNodes(NodeType::Run, true)) {
                Child^ child = gcnew Child();
                if (run->Font->Bold)
                    child->bold = true;
                if (run->Font->Underline == Underline::Single)
                    child->underline = true;
                child->fontSize = Convert::ToInt16(System::Math::Round(run->Font->Size, MidpointRounding::AwayFromZero));
                child->text = run->Text->Replace("\a", "");
                parsed_paragraph->children->Add(child);
            }
            if (para->GetChildNodes(NodeType::Run, true) == nullptr) {
                Child^ child = gcnew Child();
                child->text = "";
                parsed_paragraph->children->Add(child);
            }
            m_parser->DeserializedItems->Add(parsed_paragraph);
        }
    }

    void RtfDocumentParser::CleanParser() {
        List<Object^>^ obList = gcnew List<Object^>(m_parser->DeserializedItems);
        for each (Object ^ ob in obList) {
            if (Paragraph^ para = dynamic_cast<Paragraph^>(ob)) {
                bool flag = false;
                for each (Child ^ child in para->children) {
                    if (child->text == "Created with an evaluation copy of Aspose.Words. To remove all limitations, you can use Free Temporary License " ||
                        child->text == " HYPERLINK \"https://products.aspose.com/words/temporary-license/\" " ||
                        child->text == "https://products.aspose.com/words/temporary-license/" ||
                        child->text == "Evaluation Only. Created with Aspose.Words. Copyright 2003-2024 Aspose Pty Ltd.") {
                        flag = true;
                    }
                }
                if (flag)
                    m_parser->DeserializedItems->Remove(ob);
            }
        }
    }

    void RtfDocumentParser::GenerateJsonDocument()
    {
        String^ json;
        JsonSerializerSettings^ settings = gcnew JsonSerializerSettings();
        settings->NullValueHandling = NullValueHandling::Ignore;
        json = JsonConvert::SerializeObject(m_parser->DeserializedItems, settings);
        json = json->Replace("paragraphs", "children");
        m_jsonDocument = json;
    }
}