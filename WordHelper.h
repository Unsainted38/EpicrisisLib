#pragma once
#include "DoctorsRecords.h"
#include "FirstList.h"

using namespace System;
using namespace System::Collections::Generic;
using namespace Microsoft::Office::Interop;
using namespace System::Runtime::InteropServices;
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
		property Object^ TemplateFile {
			void set(Object^ value) {
				m_templateFile = value;
			}
			Object^ get() {
				return m_templateFile;
			}
		}
		property Object^ OutputDir {
			void set(Object^ value) {
				m_outputDir = value;
			}
			Object^ get() {
				return m_outputDir;
			}
		}
		property Word::Application^ WordApp {
			Word::Application ^ get() {
				return m_wordApp;
			}
		}
		property Word::Document^ WordDoc {
			Word::Document^ get() {
				return m_wordDoc;
			}
		}
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
		void InsertTable(Table^ table, Word::Range^ %range);
		void InsertAnalyzes(String^ analyzes);
	private:
		
		Word::Range^ GetChildFormatting(Word::Cell^ cell, Child^ child);		
		Word::Range^ GetChildFormatting(Word::Range^% para, Child^ child);
		void InsertParagraph(Paragraph^ paragraph, Word::Range^ %range);
		void InsertAnalyzes();		
	};
}