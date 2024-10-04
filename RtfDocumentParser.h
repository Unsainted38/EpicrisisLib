#pragma once

using namespace System;
using namespace System::Text;
using namespace System::Collections::Generic;
using namespace System::IO;
using namespace Aspose;
using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Tables;
using namespace Newtonsoft::Json;
using namespace Newtonsoft::Json::Linq;
using namespace System::Threading::Tasks;

namespace unsaintedWinAppLib {
	public ref class RtfDocumentParser {
	public:
		RtfDocumentParser(String^ rtfString);

		Parser^ GetParser();
		String^ GetJsonDocument();
		void ResetRtfDocumentParser(String^ rtfString);
	private:
		Parser^ m_parser;
		Aspose::Words::Document^ m_doc;
		String^ m_jsonDocument;

		void ParseRtfDocument();
		void ParseTable(Node^ node);
		void ParseParagraph(Node^ node);
		void CleanParser();

		void GenerateJsonDocument();
	};
}
