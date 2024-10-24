#pragma once
using namespace System;
using namespace System::Collections::Generic;
using namespace System::Text;

namespace unsaintedWinAppLib {
    public ref class Child
    {
    private:
        String^ Type;
        String^ Text;
        Nullable<bool> Underline;
        Nullable<bool> Bold;
        Nullable<int> FontSize;
        Nullable<bool> InLine;
        Nullable<bool> Anchor;
        List<Child^>^ Children;

    public:
        property List<Child^>^ children {
            void set(List<Child^>^ value) {
                Children = value;
            }
            List<Child^>^ get() {
                return Children;
            }
        }
        property String^ type {
            void set(String^ value) {
                Type = value;
            }
            String^ get() {
                return Type;
            }
        }
        property String^ text {
            void set(String^ value) {
                Text = value;
            }
            String^ get() {
                return Text;
            }
        }
        property Nullable<bool> bold {
            void set(Nullable<bool> value) {
                Bold = value;
            }
            Nullable<bool> get() {
                return Bold;
            }
        }
        property Nullable<bool> underline {
            void set(Nullable<bool> value) {
                Underline = value;
            }
            Nullable<bool> get() {
                return Underline;
            }
        }
        property Nullable<int> fontSize {
            void set(Nullable<int> value) {
                FontSize = value;
            }
            Nullable<int> get() {
                return FontSize;
            }
        }
        property Nullable<bool> Inline {
            void set(Nullable<bool> value) {
                InLine = value;
            }
            Nullable<bool> get() {
                return InLine;
            }
        }
        property Nullable<bool> anchor {
            void set(Nullable<bool> value) {
                Anchor = value;
            }
            Nullable<bool> get() {
                return Anchor;
            }
        }
    };

    public ref class Paragraph
    {
    private:
        String^ Type;
        List<Child^>^ Children;
        String^ Align;
    public:
        Paragraph() {
            Children = gcnew List<Child^>();
        }
        property String^ type {
            void set(String^ value) {
                Type = value;
            }
            String^ get() {
                return Type;
            }
        }
        property List<Child^>^ children {
            void set(List<Child^>^ value) {
                Children = value;
            }
            List<Child^>^ get() {
                return Children;
            }
        }
        property String^ align {
            void set(String^ value) {
                Align = value;
            }
            String^ get() {
                return Align;
            }
        }

    };

    public ref class Column
    {
    private:
        String^ Type;
        String^ Title;
    public:
        property String^ type {
            void set(String^ value) {
                Type = value;
            }
            String^ get() {
                return Type;
            }
        }
        property String^ title {
            void set(String^ value) {
                Title = value;
            }
            String^ get() {
                return Title;
            }
        }
    };

    public ref class TableCell
    {
        String^ Type;
        String^ Columntype;
        List<Paragraph^>^ Paragraphs;
        List<Child^>^ Children;
    public:
        
        property String^ type {
            void set(String^ value) {
                Type = value;
            }
            String^ get() {
                return Type;
            }
        }
        property String^ columnType {
            void set(String^ value) {
                Columntype = value;
            }
            String^ get() {
                return Columntype;
            }
        };
        property List<Paragraph^>^ paragraphs {
            void set(List<Paragraph^>^ value) {
                Paragraphs = value;
            }
            List<Paragraph^>^ get() {
                return Paragraphs;
            }
        }
        property List<Child^>^ children {
            void set(List<Child^>^ value) {
                Children = value;
            }
            List<Child^>^ get() {
                return Children;
            }
        }
    };

    public ref class TableRow
    {
    private:
        String^ Type;
        List<TableCell^>^ Children;
    public:
        TableRow() {
            Type = "tableRow";
            Children = gcnew List<TableCell^>();
        }
        property String^ type {
            void set(String^ value) {
                Type = value;
            }
            String^ get() {
                return Type;
            }
        }
        property List<TableCell^>^ children {
            void set(List<TableCell^>^ value) {
                Children = value;
            }
            List<TableCell^>^ get() {
                return Children;
            }
        }
    };

    public ref class Table
    {
    private:
        String^ Type;
        List<Column^>^ Columns;
        List<TableRow^>^ Children;

    public:
        Table() {
            Type = "table";
            Columns = gcnew List<Column^>();
            Children = gcnew List<TableRow^>();
        }
        property List<Column^>^ columns {
            void set(List<Column^>^ value) {
                Columns = value;
            }
            List<Column^>^ get() {
                return Columns;
            }
        }
        property List<TableRow^>^ children {
            void set(List<TableRow^>^ value) {
                Children = value;
            }
            List<TableRow^>^ get() {
                return Children;
            }
        }
        property String^ type {
            void set(String^ value) {
                Type = value;
            }
            String^ get() {
                return Type;
            }
        }
    };
    public ref class Parser {
    private:
        List<Object^>^ deserializeditems;
    public:
        Parser() {
            deserializeditems = gcnew List<Object^>();
        }
        Table^ GenerateTableByColumns(List<Column^>^ columnsList);
        property List<Object^>^ DeserializedItems {
            void set(List<Object^>^ value) {
                deserializeditems = value;
            }
            List<Object^>^ get() {
                return deserializeditems;
            }
        }

    };
}