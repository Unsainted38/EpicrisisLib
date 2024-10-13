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
		~WordHelper();
				
		void OpenTemplate();
		void CloseTemplate();
		void SaveTemplate();
		void InsertEpicrisToTemplate(Epicris^ epicris);
		void InsertFirstListToTemplate(DoctorsRecords^ docRec);
		void InsertDoctorsRecords(FirstList^ firstList);
	private:
		

		void InsertTable(Word::Document^ doc, Dictionary<String^, Object^>^ json);
	};
}