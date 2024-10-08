#pragma once
#include "pch.h"

using namespace System;
using namespace System::Collections::Concurrent;
using namespace System::Threading;
using namespace System::Threading::Tasks;
using namespace System::Collections::Generic;

//namespace unsaintedWinApp {
//    // ���������� ������ RtfDocumentHandler
//    ref class RtfDocumentHandler {
//    private:
//        ConcurrentQueue<Action^>^ requestQueue;
//        Object^ queueLock;
//
//    public:
//        // �����������
//        RtfDocumentHandler() {
//            requestQueue = gcnew ConcurrentQueue<Action^>();
//            queueLock = gcnew Object();
//        }
//
//        // ����������� ����� ��� ���������� �������
//        Task^ AddAnalysisAsync(String^ rtfString, Epicris^ epicris, RichTextBox^ richTextBox, ComboBox^ AnalyzesResults_comboBox, DB_Helper^ dbHelper) {
//            return Task::Run(gcnew Action([=]() {
//                EnqueueAnalysisRequest(rtfString, epicris, richTextBox, AnalyzesResults_comboBox, dbHelper);
//                }));
//        }
//
//        // ����� ���������� ������� � �������
//        void EnqueueAnalysisRequest(String^ rtfString, Epicris^ epicris, RichTextBox^ richTextBox, ComboBox^ AnalyzesResults_comboBox, DB_Helper^ dbHelper) {
//            Monitor::Enter(queueLock);
//            try {
//                // ��������� ������ � �������
//                requestQueue->Enqueue(gcnew Action([=]() {
//                    ProcessAnalysisRequest(rtfString, epicris, richTextBox, AnalyzesResults_comboBox, dbHelper);
//                    }));
//            }
//            finally {
//                Monitor::Exit(queueLock);
//            }
//        }
//    private:
//        // �����, ������� ��������� �������� ������ ��������� �������
//        void ProcessAnalysisRequest(String^ rtfString, Epicris^ epicris, RichTextBox^ richTextBox, ComboBox^ AnalyzesResults_comboBox, DB_Helper^ dbHelper) {
//            // �������� ������
//            RtfDocumentParser^ rtfParser = gcnew RtfDocumentParser(rtfString);
//            Parser^ parser = rtfParser->GetParser();
//            Dictionary<String^, Object^>^ analysisDict = gcnew Dictionary<String^, Object^>();
//            analysisDict = dbHelper->ExtractAnalyzesBlankToDictionary(AnalyzesResults_comboBox->Text);
//            analysisDict->Add("value", parser->DeserializedItems);
//
//            // ���������� ������� � epicris
//            epicris->AddAnalysisToAnalyzesList(analysisDict);
//        }
//
//    };
//}
