#pragma once
#include "pch.h"

using namespace System;
using namespace System::Collections::Concurrent;
using namespace System::Threading;
using namespace System::Threading::Tasks;
using namespace System::Collections::Generic;

//namespace unsaintedWinApp {
//    // Объявление класса RtfDocumentHandler
//    ref class RtfDocumentHandler {
//    private:
//        ConcurrentQueue<Action^>^ requestQueue;
//        Object^ queueLock;
//
//    public:
//        // Конструктор
//        RtfDocumentHandler() {
//            requestQueue = gcnew ConcurrentQueue<Action^>();
//            queueLock = gcnew Object();
//        }
//
//        // Асинхронный метод для добавления анализа
//        Task^ AddAnalysisAsync(String^ rtfString, Epicris^ epicris, RichTextBox^ richTextBox, ComboBox^ AnalyzesResults_comboBox, DB_Helper^ dbHelper) {
//            return Task::Run(gcnew Action([=]() {
//                EnqueueAnalysisRequest(rtfString, epicris, richTextBox, AnalyzesResults_comboBox, dbHelper);
//                }));
//        }
//
//        // Метод добавления запроса в очередь
//        void EnqueueAnalysisRequest(String^ rtfString, Epicris^ epicris, RichTextBox^ richTextBox, ComboBox^ AnalyzesResults_comboBox, DB_Helper^ dbHelper) {
//            Monitor::Enter(queueLock);
//            try {
//                // Добавляем запрос в очередь
//                requestQueue->Enqueue(gcnew Action([=]() {
//                    ProcessAnalysisRequest(rtfString, epicris, richTextBox, AnalyzesResults_comboBox, dbHelper);
//                    }));
//            }
//            finally {
//                Monitor::Exit(queueLock);
//            }
//        }
//    private:
//        // Метод, который выполняет основную логику обработки запроса
//        void ProcessAnalysisRequest(String^ rtfString, Epicris^ epicris, RichTextBox^ richTextBox, ComboBox^ AnalyzesResults_comboBox, DB_Helper^ dbHelper) {
//            // Основная логика
//            RtfDocumentParser^ rtfParser = gcnew RtfDocumentParser(rtfString);
//            Parser^ parser = rtfParser->GetParser();
//            Dictionary<String^, Object^>^ analysisDict = gcnew Dictionary<String^, Object^>();
//            analysisDict = dbHelper->ExtractAnalyzesBlankToDictionary(AnalyzesResults_comboBox->Text);
//            analysisDict->Add("value", parser->DeserializedItems);
//
//            // Добавление анализа в epicris
//            epicris->AddAnalysisToAnalyzesList(analysisDict);
//        }
//
//    };
//}
