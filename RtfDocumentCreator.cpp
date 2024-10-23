#include "pch.h"

namespace unsaintedWinAppLib {
    RtfDocumentCreator::RtfDocumentCreator() {
        AnalyzesRtfDict = gcnew Dictionary<String^, String^>();
    }

    RtfDocumentCreator::RtfDocumentCreator(String^ json) {
        InitializeFromJson(json);
    }

    RtfDocumentCreator::RtfDocumentCreator(Parser^ parser) {
        InitializeFromParser(parser);
    }

    RtfDocumentCreator::RtfDocumentCreator(Parser^ parser, RichTextBox^ rtb) {
        ParentRichTextBox = rtb;
        InitializeFromParser(parser);
    }

    RtfDocumentCreator::RtfDocumentCreator(String^ json, RichTextBox^ rtb) {
        ParentRichTextBox = rtb;
        InitializeFromJson(json);
    }

    RtfDocumentCreator::RtfDocumentCreator(String^ json, RichTextBox^ rtb, String^ defaultDate) {
        m_defaultDate = defaultDate;
        ParentRichTextBox = rtb;
        InitializeFromJson(json);
    }

    void RtfDocumentCreator::AddRowToTable(RichTextBox^ richTextBox)
    {
        // Загружаем содержимое richTextBox как документ Aspose.Words
        auto doc = gcnew Document(gcnew IO::MemoryStream(System::Text::Encoding::UTF8->GetBytes(richTextBox->Rtf)));
        auto builder = gcnew DocumentBuilder(doc);

        // Находим последнюю таблицу в документе
        Aspose::Words::Tables::Table^ table = FindLastTable(doc);
        if (table == nullptr) return;  // Если таблица не найдена, ничего не делаем

        // Находим последнюю строку таблицы
        Row^ lastRow = table->LastRow;

        // Создаем новую строку
        Row^ newRow = gcnew Row(doc);
        
        table->Rows->Add(newRow);

        // Копируем структуру ячеек и заполняем их
        for (int i = 0; i < lastRow->Cells->Count; i++)
        {
            Cell^ newCell = gcnew Cell(doc);
            newRow->Cells->Add(newCell);

            // Проверяем, содержит ли предыдущая ячейка дату
            String^ previousCellText = lastRow->Cells[i]->ToString(SaveFormat::Text)->Trim();
            DateTime dateValue;
            if (DateTime::TryParse(previousCellText, dateValue))
            {
                // Если в предыдущей ячейке дата, добавляем в эту дату
                newCell->AppendChild(gcnew Aspose::Words::Paragraph(doc));
                builder->MoveTo(newCell->FirstParagraph);
                builder->Write(dateValue.ToShortDateString());
            }
            else
            {
                // Если нет, оставляем пустую ячейку
                newCell->AppendChild(gcnew Aspose::Words::Paragraph(doc));
            }
        }

        // Сохраняем обновленный документ обратно в RichTextBox
        IO::MemoryStream^ stream = gcnew IO::MemoryStream();
        doc->Save(stream, SaveFormat::Rtf);
        richTextBox->Rtf = System::Text::Encoding::UTF8->GetString(stream->ToArray());
    }

    Dictionary<String^, String^>^ RtfDocumentCreator::GetAnalyzesDict() {
        return AnalyzesRtfDict;
    }

    String^ RtfDocumentCreator::GetRtfDocument() {
        return m_rtfDocument;
    }

    String^ RtfDocumentCreator::GetRtfDocumentFromDict(String^ key) {
        return AnalyzesRtfDict[key];
    }

    Parser^ RtfDocumentCreator::GetParser() {
        return m_parser;
    }

    void RtfDocumentCreator::ResetRtfDocumentCreator(String^ json) {
        InitializeFromJson(json);
    }

    void RtfDocumentCreator::ResetRtfDocumentCreator(Parser^ parser) {
        InitializeFromParser(parser);
    }

    void RtfDocumentCreator::InitializeDict(Dictionary<String^, String^>^ dict) {
        for each (KeyValuePair<String^, String^> ^ kvp in dict) {
            InitializeFromJson(kvp->Value);
            AnalyzesRtfDict->Add(kvp->Key, m_rtfDocument);
        }
    }

    void RtfDocumentCreator::ResetRtfDocumentCreator(String^ json, RichTextBox^ rtb, String^ defaultDate) {
        m_defaultDate = defaultDate;
        ParentRichTextBox = rtb;
        InitializeFromJson(json);
    }

    void RtfDocumentCreator::GenerateParser(String^ jsonString) {
        Parser^ parser = gcnew Parser();
        parser->DeserializedItems = gcnew List<Object^>();
        List<JObject^>^ items = JsonConvert::DeserializeObject<List<JObject^>^>(jsonString);

        for each (JObject ^ item in items) {
            String^ type;
            String^ align;
            if (item->ContainsKey("type"))
                type = item["type"]->ToString();
            if (item->ContainsKey("align"))
                align = item["align"]->ToString();
            if (type == "paragraph") {
                Paragraph^ m_paragraph = gcnew Paragraph();
                List<JObject^>^ children;
                m_paragraph->type = type;
                m_paragraph->align = align;
                if (item->ContainsKey("children")) {
                    m_paragraph->children = gcnew List<Child^>();
                    children = JsonConvert::DeserializeObject<List<JObject^>^>(item["children"]->ToString());
                    for each (JObject ^ child in children) {
                        Child^ m_child;
                        if (child->ContainsKey("type")) {
                            String^ type = child["type"]->ToString();
                            List<Child^>^ m_children = gcnew List<Child^>();
                            int parentfontsize = 11;
                            if (child->ContainsKey("fontSize"))
                                parentfontsize = child["fontSize"]->ToObject<int>();
                            children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                            if (type == "dateInput") {
                                for each (JObject ^ small_child in children) {
                                    Child^ m_child = GenerateChild(small_child);
                                    m_child->fontSize = parentfontsize;
                                    if (small_child->ContainsKey("fontSize"))
                                        m_child->fontSize = small_child["fontSize"]->ToObject<int>();
                                    m_children->Add(m_child);
                                }
                            }
                            else if (m_child->type == "paragraph") {
                                for each (JObject ^ small_child in children) {
                                    Child^ m_child = GenerateChild(small_child);
                                    m_child->fontSize = parentfontsize;
                                    if (small_child->ContainsKey("fontSize"))
                                        m_child->fontSize = small_child["fontSize"]->ToObject<int>();
                                    m_children->Add(m_child);
                                }
                            }
                            m_child = GenerateChild(child);
                            m_child->type = type;
                            m_child->children = m_children;
                        }
                        else {
                            m_child = GenerateChild(child);
                        }

                        m_paragraph->children->Add(m_child);
                    }
                }
                parser->DeserializedItems->Add(m_paragraph);
            }
            else if (type == "table") {
                Table^ m_table = gcnew Table();
                List<JObject^>^ columns;
                List<JObject^>^ children;
                if (item->ContainsKey("columns")) {
                    columns = JsonConvert::DeserializeObject<List<JObject^>^>(item["columns"]->ToString());
                    m_table->columns = gcnew List<Column^>();
                    for each (JObject ^ column in columns) {
                        Column^ m_column = gcnew Column();
                        String^ type;
                        String^ title;
                        if (column->ContainsKey("type"))
                            type = column["type"]->ToString();
                        if (column->ContainsKey("title"))
                            title = column["title"]->ToString();
                        m_column->title = title;
                        m_column->type = type;
                        m_table->columns->Add(m_column);
                    }
                }
                if (item->ContainsKey("children")) {
                    children = JsonConvert::DeserializeObject<List<JObject^>^>(item["children"]->ToString());
                    m_table->children = gcnew List<TableRow^>();
                    for each (JObject ^ child in children) {
                        String^ type;
                        List<JObject^>^ children;
                        TableRow^ m_tableRow = gcnew TableRow();
                        m_tableRow->children = gcnew List<TableCell^>();
                        if (child->ContainsKey("type"))
                            type = child["type"]->ToString();
                        m_tableRow->type = type;

                        if (type == "tableRow") {
                            children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());

                            for each (JObject ^ child in children) {
                                TableCell^ m_tableCell = gcnew TableCell();
                                String^ type;
                                List<JObject^>^ children;
                                if (child->ContainsKey("type"))
                                    type = child["type"]->ToString();
                                m_tableCell->type = type;
                                if (type == "tableHeaderCell") {
                                    children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                                    m_tableCell->children = gcnew List<Child^>();
                                    for each (JObject ^ child in children) {
                                        Child^ m_child = GenerateChild(child);
                                        m_tableCell->children->Add(m_child);
                                    }
                                }
                                else if (type == "tableDataCell") {
                                    String^ columnType;
                                    List<JObject^>^ children;
                                    m_tableCell->paragraphs = gcnew List<Paragraph^>();
                                    if (child->ContainsKey("columnType"))
                                        columnType = child["columnType"]->ToString();
                                    m_tableCell->columnType = columnType;
                                    if (columnType == "date") {
                                        children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                                        for each (JObject ^ child in children) {
                                            String^ type;
                                            List<JObject^>^ children;
                                            Paragraph^ m_paragraph = gcnew Paragraph();
                                            m_paragraph->children = gcnew List<Child^>();
                                            if (child->ContainsKey("type"))
                                                type = child["type"]->ToString();
                                            m_paragraph->type = type;
                                            if (type == "dateInput") {
                                                children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                                                for each (JObject ^ child in children) {
                                                    Child^ m_child = GenerateChild(child);
                                                    m_paragraph->children->Add(m_child);
                                                }
                                            }
                                            else if (type == "paragraph") {
                                                children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                                                for each (JObject ^ child in children) {
                                                    Child^ m_child = GenerateChild(child);
                                                    m_paragraph->children->Add(m_child);
                                                }
                                            }
                                            m_tableCell->paragraphs->Add(m_paragraph);
                                        }
                                    }
                                    else if (columnType == "text") {
                                        children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                                        for each (JObject ^ child in children) {
                                            String^ type;
                                            List<JObject^>^ children;
                                            Paragraph^ m_paragraph = gcnew Paragraph();
                                            m_paragraph->children = gcnew List<Child^>();
                                            if (child->ContainsKey("type"))
                                                type = child["type"]->ToString();
                                            m_paragraph->type = type;
                                            if (type == "dateInput") {
                                                children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                                                for each (JObject ^ child in children) {
                                                    Child^ m_child = GenerateChild(child);
                                                    m_paragraph->children->Add(m_child);
                                                }
                                            }
                                            else if (type == "paragraph") {
                                                children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                                                for each (JObject ^ child in children) {
                                                    Child^ m_child = GenerateChild(child);
                                                    m_paragraph->children->Add(m_child);
                                                }
                                            }
                                            m_tableCell->paragraphs->Add(m_paragraph);
                                        }

                                    }
                                }
                                m_tableRow->children->Add(m_tableCell);
                            }
                        }
                        m_table->children->Add(m_tableRow);
                    }
                }
                parser->DeserializedItems->Add(m_table);
            }
        }
        m_parser = parser;
    }

    Aspose::Words::Tables::Table^ RtfDocumentCreator::FindLastTable(Document^ doc)
    {
        for each (Section ^ section in doc->Sections)
        {
            for each (Node ^ node in section->Body->GetChildNodes(NodeType::Any, true))
            {
                if (Aspose::Words::Tables::Table^ table = dynamic_cast<Aspose::Words::Tables::Table^>(node))
                {
                    return table;  // Возвращаем первую найденную таблицу
                }
            }
        }
        return nullptr;
    }

    void RtfDocumentCreator::InitializeFromJson(String^ json) {
        m_jsonDocument = json;
        GenerateParser();
        CleanParser();
        GenerateRtfDocument();
        CleanRtfDoc();
    }

    void RtfDocumentCreator::InitializeFromParser(Parser^ parser) {
        m_parser = parser;
        CleanParser();
        GenerateRtfDocument();
        CleanRtfDoc();
    }

    void RtfDocumentCreator::GenerateRtfDocument() {
        auto doc = gcnew Document();
        auto builder = gcnew DocumentBuilder(doc);
        for each (Object ^ ob in m_parser->DeserializedItems) {
            if (Table^ table = dynamic_cast<Table^>(ob)) {
                auto asposeTable = builder->StartTable();
                int TotalWidth;
                int dateColumnsCounter = 0;
                for each (Column ^ column in table->columns) {
                    if (column->type == "date") {
                        dateColumnsCounter++;
                    }
                }
                if (parentRichTextBox == nullptr)
                    TotalWidth = 18800;
                else if (dateColumnsCounter) {
                    TotalWidth = GetRichTextBoxWidthInTwips(parentRichTextBox) - 1000 * dateColumnsCounter - 500;
                }
                else {
                    TotalWidth = GetRichTextBoxWidthInTwips(parentRichTextBox) - 500;
                }
                Dictionary<String^, int>^ widths = CalculateColumnsWidths(table->columns, TotalWidth);
                if (widths == nullptr)
                    return;
                
                for each (TableRow ^ row in table->children) {
                    int  i = 0;
                    for each (TableCell ^ cell in row->children) {                        
                        auto asposeCell = builder->InsertCell();
                        builder->CurrentParagraph->ParagraphFormat->ClearFormatting();
                        if (i != 0)
                            builder->CurrentParagraph->ParagraphFormat->Alignment = ParagraphAlignment::Center;
                        asposeCell->CellFormat->VerticalAlignment = CellVerticalAlignment::Center;                     
                        asposeCell->CellFormat->Borders->Color = System::Drawing::Color::Black;
                        asposeCell->CellFormat->BottomPadding = 2;
                        asposeCell->CellFormat->TopPadding = 2;
                        asposeCell->CellFormat->RightPadding = 0;
                        asposeCell->CellFormat->LeftPadding = 0;
                        asposeCell->CellFormat->Width = widths[table->columns[i]->title] / 20;
                        if (cell->type == "tableHeaderCell") {
                            for each (Child ^ child in cell->children) {
                                ChildTextFormatting(child, builder);
                                if (child->text != nullptr)
                                    builder->Write(child->text);
                                builder->Font->ClearFormatting();
                            }
                        }
                        else if (cell->type == "tableDataCell") {
                            for each (Paragraph ^ para in cell->paragraphs) {
                                auto asposePara = builder->CurrentParagraph;
                                if (para->align == "center")
                                    asposePara->ParagraphFormat->Alignment = ParagraphAlignment::Center;
                                else if (para->align == "left")
                                    asposePara->ParagraphFormat->Alignment = ParagraphAlignment::Left;
                                else if (para->align == "right")
                                    asposePara->ParagraphFormat->Alignment = ParagraphAlignment::Right;

                                if (para->type == "paragraph") {
                                    for each (Child ^ child in para->children) {
                                        ChildTextFormatting(child, builder);
                                        if (child->text != nullptr)
                                            builder->Write(child->text);
                                        builder->Font->ClearFormatting();
                                    }
                                }
                                else if (para->type == "dateInput") {
                                    for each (Child ^ child in para->children) {
                                        ChildTextFormatting(child, builder);
                                        if (String::IsNullOrEmpty(m_defaultDate))
                                            builder->Write(DateTime::Now.AddDays(-12).ToShortDateString());
                                        else {
                                            builder->Write(m_defaultDate);
                                        }
                                        builder->Font->ClearFormatting();
                                    }
                                }
                            }
                        }
                        i++;
                    }
                    builder->EndRow();
                }
                builder->EndTable();
                //asposeTable->AutoFit(AutoFitBehavior::AutoFitToWindow);
                PageSetup^ pageSetup = doc->FirstSection->PageSetup;
                double pageWidth = pageSetup->PageWidth - pageSetup->LeftMargin - pageSetup->RightMargin;
                asposeTable->PreferredWidth = PreferredWidth::FromPoints(pageWidth);
                
                asposeTable->Alignment = TableAlignment::Center;
            }
            else if (Paragraph^ para = dynamic_cast<Paragraph^>(ob)) {
                auto asposePara = builder->InsertParagraph();
                if (para->align == "center")
                    asposePara->ParagraphFormat->Alignment = ParagraphAlignment::Center;
                else if (para->align == "left")
                    asposePara->ParagraphFormat->Alignment = ParagraphAlignment::Left;
                else if (para->align == "right")
                    asposePara->ParagraphFormat->Alignment = ParagraphAlignment::Right;
                if (para->type == "paragraph") {
                    for each (Child ^ child in para->children) {
                        if (child->type == "dateInput") {
                            for each (Child ^ small_child in child->children) {
                                ChildTextFormatting(small_child, builder, child);
                                if (String::IsNullOrEmpty(m_defaultDate))
                                    builder->Write(DateTime::Now.AddDays(-12).ToShortDateString());
                                else {
                                    builder->Write(m_defaultDate);
                                }
                                builder->Font->ClearFormatting();
                            }
                        }
                        else {
                            ChildTextFormatting(child, builder);
                            if (child->text != nullptr)
                                builder->Write(child->text);
                            builder->Font->ClearFormatting();
                        }                        
                    }
                }
                else if (para->type == "dateInput") {
                    for each (Child ^ child in para->children) {
                        ChildTextFormatting(child, builder);
                        builder->Write(DateTime::Now.AddDays(-12).ToShortDateString());
                        builder->Font->ClearFormatting();
                        //Aspose::Words::Drawing::Charts::Chart^ builder;
                        //builder->InsertChart(ChartType::Area, 0, 0);
                    }
                }
                //builder->InsertTextInput("date", Fields::TextFormFieldType::Date, "dd.mm.yyyy", DateTime::Now.ToShortDateString(), 20);
            }
        }
        auto stream = gcnew MemoryStream();
        doc->Save(stream, SaveFormat::Rtf);
        String^ rtfString = Encoding::UTF8->GetString(stream->ToArray());
        stream->Close();

        m_rtfDocument = rtfString;
    }

    void RtfDocumentCreator::ChildTextFormatting(Child^ child, DocumentBuilder^% builder) {
        if (child->bold.HasValue)
            builder->Font->Bold = child->bold.Value;
        if (child->underline.HasValue)
            builder->Font->Underline = child->bold.Value ? Underline::Single : Underline::None;
        if (child->fontSize.HasValue)
            builder->Font->Size = Convert::ToDouble(child->fontSize);
    }

    void RtfDocumentCreator::ChildTextFormatting(Child^ child, DocumentBuilder^% builder, Child^ parent) {
        if (child->bold.HasValue)
            builder->Font->Bold = child->bold.Value;
        else if (parent->bold.HasValue)
            builder->Font->Bold = parent->bold.Value;
        if (child->underline.HasValue)
            builder->Font->Underline = child->bold.Value ? Underline::Single : Underline::None;
        else if (parent->underline.HasValue)
            builder->Font->Underline = parent->underline.Value ? Underline::Single : Underline::None;
        if (child->fontSize.HasValue)
            builder->Font->Size = Convert::ToDouble(child->fontSize);
        else if (parent->fontSize.HasValue)
            builder->Font->Size = Convert::ToDouble(parent->fontSize);
    }

    void RtfDocumentCreator::GenerateParser() {
        Parser^ parser = gcnew Parser();
        parser->DeserializedItems = gcnew List<Object^>();
        List<JObject^>^ items = JsonConvert::DeserializeObject<List<JObject^>^>(m_jsonDocument);

        for each (JObject ^ item in items) {
            String^ type;
            String^ align;
            if (item->ContainsKey("type"))
                type = item["type"]->ToString();
            if (item->ContainsKey("align"))
                align = item["align"]->ToString();
            if (type == "paragraph") {
                Paragraph^ m_paragraph = gcnew Paragraph();
                List<JObject^>^ children;
                m_paragraph->type = type;
                m_paragraph->align = align;
                if (item->ContainsKey("children")) {
                    m_paragraph->children = gcnew List<Child^>();
                    children = JsonConvert::DeserializeObject<List<JObject^>^>(item["children"]->ToString());
                    for each (JObject ^ child in children) {
                        Child^ m_child;
                        if (child->ContainsKey("type")) {
                            String^ type = child["type"]->ToString();
                            List<Child^>^ m_children = gcnew List<Child^>();
                            int parentfontsize = 11;
                            if (child->ContainsKey("fontSize"))
                                parentfontsize = child["fontSize"]->ToObject<int>();
                            children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                            if (type == "dateInput") {                                
                                for each (JObject ^ small_child in children) {
                                    Child^ m_child = GenerateChild(small_child);
                                    m_child->fontSize = parentfontsize;
                                    if (small_child->ContainsKey("fontSize"))
                                        m_child->fontSize = small_child["fontSize"]->ToObject<int>();
                                    m_children->Add(m_child);
                                }
                            }
                            else if (m_child->type == "paragraph") {                                
                                for each (JObject ^ small_child in children) {
                                    Child^ m_child = GenerateChild(small_child);
                                    m_child->fontSize = parentfontsize;
                                    if (small_child->ContainsKey("fontSize"))
                                        m_child->fontSize = small_child["fontSize"]->ToObject<int>();
                                    m_children->Add(m_child);
                                }
                            }
                            m_child = GenerateChild(child);
                            m_child->type = type;
                            m_child->children = m_children;
                        }
                        else {
                            m_child = GenerateChild(child);
                        }
                        
                        m_paragraph->children->Add(m_child);
                    }
                }
                parser->DeserializedItems->Add(m_paragraph);
            }
            else if (type == "table") {
                Table^ m_table = gcnew Table();
                List<JObject^>^ columns;
                List<JObject^>^ children;
                if (item->ContainsKey("columns")) {
                    columns = JsonConvert::DeserializeObject<List<JObject^>^>(item["columns"]->ToString());
                    m_table->columns = gcnew List<Column^>();
                    for each (JObject ^ column in columns) {
                        Column^ m_column = gcnew Column();
                        String^ type;
                        String^ title;
                        if (column->ContainsKey("type"))
                            type = column["type"]->ToString();
                        if (column->ContainsKey("title"))
                            title = column["title"]->ToString();
                        m_column->title = title;
                        m_column->type = type;
                        m_table->columns->Add(m_column);
                    }
                }
                if (item->ContainsKey("children")) {
                    children = JsonConvert::DeserializeObject<List<JObject^>^>(item["children"]->ToString());
                    m_table->children = gcnew List<TableRow^>();
                    for each (JObject ^ child in children) {
                        String^ type;
                        List<JObject^>^ children;
                        TableRow^ m_tableRow = gcnew TableRow();
                        m_tableRow->children = gcnew List<TableCell^>();
                        if (child->ContainsKey("type"))
                            type = child["type"]->ToString();
                        m_tableRow->type = type;

                        if (type == "tableRow") {
                            children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());

                            for each (JObject ^ child in children) {
                                TableCell^ m_tableCell = gcnew TableCell();
                                String^ type;
                                List<JObject^>^ children;
                                if (child->ContainsKey("type"))
                                    type = child["type"]->ToString();
                                m_tableCell->type = type;
                                if (type == "tableHeaderCell") {
                                    children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                                    m_tableCell->children = gcnew List<Child^>();
                                    for each (JObject ^ child in children) {
                                        Child^ m_child = GenerateChild(child);
                                        m_tableCell->children->Add(m_child);
                                    }
                                }
                                else if (type == "tableDataCell") {
                                    String^ columnType;
                                    List<JObject^>^ children;
                                    m_tableCell->paragraphs = gcnew List<Paragraph^>();
                                    if (child->ContainsKey("columnType"))
                                        columnType = child["columnType"]->ToString();
                                    m_tableCell->columnType = columnType;
                                    if (columnType == "date") {
                                        children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                                        for each (JObject ^ child in children) {
                                            String^ type;
                                            List<JObject^>^ children;
                                            Paragraph^ m_paragraph = gcnew Paragraph();
                                            m_paragraph->children = gcnew List<Child^>();
                                            if (child->ContainsKey("type"))
                                                type = child["type"]->ToString();
                                            m_paragraph->type = type;
                                            if (type == "dateInput") {
                                                children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                                                for each (JObject ^ child in children) {
                                                    Child^ m_child = GenerateChild(child);
                                                    m_paragraph->children->Add(m_child);
                                                }
                                            }
                                            else if (type == "paragraph") {
                                                children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                                                for each (JObject ^ child in children) {
                                                    Child^ m_child = GenerateChild(child);
                                                    m_paragraph->children->Add(m_child);
                                                }
                                            }
                                            m_tableCell->paragraphs->Add(m_paragraph);
                                        }
                                    }
                                    else if (columnType == "text") {
                                        children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                                        for each (JObject ^ child in children) {
                                            String^ type;
                                            List<JObject^>^ children;
                                            Paragraph^ m_paragraph = gcnew Paragraph();
                                            m_paragraph->children = gcnew List<Child^>();
                                            if (child->ContainsKey("type"))
                                                type = child["type"]->ToString();
                                            m_paragraph->type = type;
                                            if (type == "dateInput") {
                                                children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                                                for each (JObject ^ child in children) {
                                                    Child^ m_child = GenerateChild(child);
                                                    m_paragraph->children->Add(m_child);
                                                }
                                            }
                                            else if (type == "paragraph") {
                                                children = JsonConvert::DeserializeObject<List<JObject^>^>(child["children"]->ToString());
                                                for each (JObject ^ child in children) {
                                                    Child^ m_child = GenerateChild(child);
                                                    m_paragraph->children->Add(m_child);
                                                }
                                            }
                                            m_tableCell->paragraphs->Add(m_paragraph);
                                        }

                                    }
                                }
                                m_tableRow->children->Add(m_tableCell);
                            }
                        }
                        m_table->children->Add(m_tableRow);
                    }
                }
                parser->DeserializedItems->Add(m_table);
            }
        }
        m_parser = parser;
    }

    Child^ RtfDocumentCreator::GenerateChild(JObject^ child) {
        Child^ m_child = gcnew Child();
        Nullable<bool> bold;
        Nullable<bool> underline;
        Nullable<bool> anchor;
        Nullable<bool> inLine;
        int fontSize = 11;
        String^ text;
        if (child->ContainsKey("bold"))
            bold = child["bold"]->ToObject<bool>();
        if (child->ContainsKey("text"))
            text = child["text"]->ToString();
        if (child->ContainsKey("underline"))
            underline = child["underline"]->ToObject<bool>();
        if (child->ContainsKey("fontSize"))
            fontSize = child["fontSize"]->ToObject<int>();
        if (child->ContainsKey("anchor"))
            anchor = child["anchor"]->ToObject<bool>();
        if (child->ContainsKey("inline"))
            inLine = child["inline"]->ToObject<bool>();

        m_child->bold = bold;
        m_child->fontSize = fontSize;
        m_child->underline = underline;
        m_child->text = text;
        m_child->Inline = inLine;
        m_child->anchor = anchor;
        return m_child;
    }

    void RtfDocumentCreator::CleanParser() {
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

    void RtfDocumentCreator::CleanRtfDoc() {
        if (m_rtfDocument == nullptr)
            return;
        RemoveBetween(m_rtfDocument, WATERMARK_START, WATERMARK_END);
        RemoveBetween(m_rtfDocument, R"({\footer\pard\plain)", R"(Pty Ltd.}{\rtlch\afs24\ltrch\fs24\par}})");
        for (int i = 0; i < 2; i++) {
            int start_index = m_rtfDocument->IndexOf(R"(\par}\pard)");
            if (start_index < 0)
                return;
            m_rtfDocument = m_rtfDocument->Remove(start_index, String(R"(\par}\pard)").Length);            
        }            
        
    }

    void RtfDocumentCreator::RemoveBetween(String^% input, String^ start, String^ end) {
        // Находим индекс начала подстроки start
        if (input == nullptr)
            return;
        int startIndex = input->IndexOf(start);

        // Если start не найдена, возвращаем исходную строку
        if (startIndex == -1) {
            return;
        }

        // Находим индекс конца подстроки end, начиная с позиции после подстроки start
        int endIndex = input->IndexOf(end, startIndex + start->Length);

        // Если end не найдена, возвращаем исходную строку
        if (endIndex == -1) {
            return;
        }

        // Удаляем всё от start до end включительно
        input = input->Remove(startIndex, (endIndex + end->Length) - startIndex);
    }

    void RtfDocumentCreator::SortJsonByPosition(String^ json)
    {
        // Парсим JSON строку в JArray (массив JSON объектов)
        JArray^ jsonArray = JArray::Parse(json);

        // Сортируем JArray по значению ключа "position"
        List<JToken^>^ sortedList = gcnew List<JToken^>();
        for each (JToken ^ token in jsonArray)
        {
            sortedList->Add(token);
        }

        // Используем лямбда-функцию для сортировки по ключу "position"
        /*sortedList->Sort(gcnew Comparison<JToken^>(
            [](JToken^ a, JToken^ b) {
                return (int)a->Value<int>("position") - (int)b->Value<int>("position");
            }
        ));*/

        // Преобразуем отсортированный список обратно в JArray
        JArray^ sortedJsonArray = gcnew JArray(sortedList);

        // Сериализуем JArray обратно в строку
        String^ sortedJsonString = sortedJsonArray->ToString();
    }

    Dictionary<String^, int>^ RtfDocumentCreator::CalculateColumnsWidths(List<Column^>^ columns, int TotalWidth) {
        Dictionary<String^, int>^ widths = gcnew Dictionary<String^, int>();

        double TotaltextLength = 0;
        double k;
        for each (Column ^ column in columns) {            
            int length = BiggestStringWordLength(column->title);
            k = K(length);
            TotaltextLength += length * k;                                  
        }
        // Ширина ячеек
        int currpos = 0;
        double s = 0;
        for each (Column ^ column in columns) {
            int CellWidth;                           
            int length = BiggestStringWordLength(column->title);
            k = K(length);
            CellWidth = Convert::ToInt32((k * length / TotaltextLength) * TotalWidth);
            //s += k * CelltextLength / TotaltextLength;
            if (column->type == "date")
                CellWidth += 1000;                        
            widths->Add(column->title, CellWidth);            
        }
        return widths;
    }

    double RtfDocumentCreator::K(int length) {
        double k;
        if (length <= 2)
            k = 1.35;
        else if (length <= 3)
            k = 1.15;
        else if (length <= 4)
            k = 1;
        else if (length <= 5)
            k = 0.8;
        else
            k = 0.7;
        return k;
    }

    int RtfDocumentCreator::BiggestStringWordLength(String^ str) {  
        array<String^>^ words = str->Split(' ');
        String^ biggestWord = words[0];
        for (int i = 1; i < words->Length; i++) {
            if (biggestWord->Length < words[i]->Length)
                biggestWord = words[i];
        }
        return biggestWord->Length;                
    }
    int RtfDocumentCreator::PixelsToTwipsX(int pixels, Graphics^ g)
    {
        // Получаем DPI по оси X (горизонтальная плотность пикселей)
        float dpiX = g->DpiX;
        // Переводим пиксели в дюймы, затем в твипы
        return (int)((pixels / dpiX) * 1440);
    }
    int RtfDocumentCreator::PixelsToTwipsY(int pixels, Graphics^ g)
    {
        // Получаем DPI по оси Y (вертикальная плотность пикселей)
        float dpiY = g->DpiY;
        // Переводим пиксели в дюймы, затем в твипы
        return (int)((pixels / dpiY) * 1440);
    }
    int RtfDocumentCreator::GetRichTextBoxWidthInTwips(RichTextBox^ richTextBox)
    {
        // Создаем объект Graphics для текущего RichTextBox
        Graphics^ g = richTextBox->CreateGraphics();

        // Размеры в пикселях
        int widthInPixels = richTextBox->Width;        

        // Преобразуем пиксели в твипы
        int widthInTwips = PixelsToTwipsX(widthInPixels, g);        

        return widthInTwips;
    }
}
