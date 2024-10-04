#include "pch.h"

namespace unsaintedWinAppLib {
    Table^ Parser::GenerateTableByColumns(List<Column^>^ columnsList)
    {
        Table^ table = gcnew Table();
        table->columns->AddRange(columnsList);        
        for (int i = 0; i < 2; i++) {
            TableRow^ row = gcnew TableRow();
            for each (Column ^ column in columnsList) {
                TableCell^ cell = gcnew TableCell();
                if (i == 0) {
                    cell->type = "tableHeaderCell";
                    cell->children = gcnew List<Child^>();
                    Child^ child = gcnew Child();
                    child->fontSize = 8;
                    child->text = column->title;
                    cell->children->Add(child);
                }
                else if (i == 1) {
                    cell->type = "tableDataCell";
                    cell->paragraphs = gcnew List<Paragraph^>();
                    Paragraph^ paragraph = gcnew Paragraph();
                    paragraph->align = "center";
                    if (column->type == "date") {
                        cell->columnType = "date";
                        paragraph->type == "dateInput";
                        Child^ child = gcnew Child();
                        child->text = "";
                        paragraph->children->Add(child);                        
                    }
                    else if (column->type == "text") {
                        cell->columnType = "text";
                        paragraph->type == "paragraph";
                        Child^ child = gcnew Child();
                        child->text = "";
                        paragraph->children->Add(child);
                    }
                    cell->paragraphs->Add(paragraph);
                }
                row->children->Add(cell);
            }
            table->children->Add(row);
        }
        return table;
    }
}