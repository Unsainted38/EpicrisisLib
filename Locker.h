#pragma once
using namespace System::Timers;


namespace unsaintedWinAppLib {
	public ref class Locker {
		Timers::Timer^ m_elapsedTimer;
		bool m_lock;
	public:
		Locker() {
			m_lock = false;
			m_elapsedTimer = gcnew Timers::Timer(200);
			m_elapsedTimer->Elapsed += gcnew System::Timers::ElapsedEventHandler(this, &unsaintedWinAppLib::Locker::OnElapsed);
		}
		Locker(double interval) {
			m_lock = false;
			m_elapsedTimer = gcnew Timers::Timer(interval);
			m_elapsedTimer->Elapsed += gcnew System::Timers::ElapsedEventHandler(this, &unsaintedWinAppLib::Locker::OnElapsed);
		}
		bool isLocked() {
			return m_lock;
		}
		void Lock() {
			m_lock = true;
			m_elapsedTimer->Start();
		}
		void Unlock() {
			m_lock = false;
			m_elapsedTimer->Stop();
		}
	private:
		void OnElapsed(System::Object^ sender, System::Timers::ElapsedEventArgs^ e) {
			Unlock();
		}
	};
}


