#pragma once
#include "DoctorsRecords.h"
#include "FirstList.h"

using namespace System;
using namespace System::Collections::Generic;
using namespace Microsoft::Office::Interop;
using namespace Newtonsoft::Json;



namespace unsaintedWinAppLib {
	public ref class WordHelper {

		Epicris^ m_epicris;
		Word::Application^ m_wordApp;
		Word::Document^ m_wordDoc;
		Word::Documents^ m_wordDocs;
		Object^ m_templateFile;
		Object^ m_outputDir;
		Object^ missing = Type::Missing;

	public:
		WordHelper();
		WordHelper(String^ templateFilePath, String^ outputDirPath, Epicris^ epicris);
		WordHelper(String^ templateFilePath, String^ outputDirPath, DoctorsRecords^ docRec);
		WordHelper(String^ templateFilePath, String^ outputDirPath, FirstList^ firstList);
		~WordHelper();
				
		void OpenTemplate();
		void CloseTemplate();
		void SaveTemplate();
		void InsertEpicrisToTemplate();
		void InsertFirstListToTemplate();
		void InsertDoctorsRecords();
		
	private:
		
		void InsertTable(Table^ table);
		void InsertParagraph(Paragraph^ parapraph);
		void InsertAnalyzes();
		void InsertTables(Word::Document^ doc, Dictionary<String^, Object^>^ json);
	};
}