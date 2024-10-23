#pragma once
using namespace System;
using namespace System::Text;
using namespace System::Collections::Generic;
using namespace System::IO;
using namespace System::Windows::Forms;
using namespace System::Drawing;
using namespace Aspose;
using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Drawing::Charts;
using namespace Newtonsoft::Json;
using namespace Newtonsoft::Json::Linq;


namespace unsaintedWinAppLib {
	public ref class RtfDocumentCreator
	{
	public:
		RtfDocumentCreator();
		RtfDocumentCreator(String^ json);
		RtfDocumentCreator(Parser^ parser);
		RtfDocumentCreator(Parser^ parser, RichTextBox^ rtb);
		RtfDocumentCreator(String^ json, RichTextBox^ rtb);
		RtfDocumentCreator(String^ json, RichTextBox^ rtb, String^ defaultDate);

		void AddRowToTable(RichTextBox^ richTextBox);
		Dictionary<String^, String^>^ GetAnalyzesDict();
		String^ GetRtfDocument();
		String^ GetRtfDocumentFromDict(String^ key);
		Parser^ GetParser();
		void ResetRtfDocumentCreator(String^ json);
		void ResetRtfDocumentCreator(Parser^ parser);
		void ResetRtfDocumentCreator(String^ json, RichTextBox^ rtb, String^ defaultDate);
		void InitializeDict(Dictionary<String^, String^>^ dict);
		property RichTextBox^ ParentRichTextBox {
			void set(RichTextBox^ value) {
				parentRichTextBox = value;
			}
			RichTextBox^ get() {
				return parentRichTextBox;
			}
		}
		void GenerateParser(String^ jsonString);
	private:
		RichTextBox^ parentRichTextBox;
		Aspose::Words::Tables::Table^ FindLastTable(Document^ doc);

		void InitializeFromJson(String^ json);
		void InitializeFromParser(Parser^ parser);
		void GenerateRtfDocument();
		void ChildTextFormatting(Child^ child, DocumentBuilder^% builder);
		void ChildTextFormatting(Child^ child, DocumentBuilder^% builder, Child^ parent);
		void CleanRtfDoc();
		void RemoveBetween(String^% input, String^ start, String^ end);
		void SortJsonByPosition(String^ json);

		void CleanParser();
		void GenerateParser();
		Child^ GenerateChild(JObject^ child);

		Dictionary<String^, int>^ CalculateColumnsWidths(List<Column^>^ columns, int TotalWidth);
		double K(int length);
		int BiggestStringWordLength(String^ str);
		int PixelsToTwipsX(int pixels, Graphics^ g);
		int PixelsToTwipsY(int pixels, Graphics^ g);
		int GetRichTextBoxWidthInTwips(RichTextBox^ richTextBox);

		String^ m_defaultDate;
		String^ m_jsonDocument;
		String^ m_rtfDocument;
		Parser^ m_parser;
		Dictionary<String^, String^>^ AnalyzesRtfDict;

		String^ WATERMARK_START = R"({\header\pard\plain)";
		String^ WATERMARK_END = R"(https://products.aspose.com/words/temporary-license/}}})";
	};
}
