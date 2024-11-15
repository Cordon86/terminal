/*++
Copyright (c) Microsoft Corporation Licensed under the MIT license.

Class Name:
- WindowEmperor.h

Abstract:
- The WindowEmperor is our class for managing the single Terminal process
  with all our windows. It will be responsible for handling the commandline
  arguments. It will initially try to find another terminal process to
  communicate with. If it does, it'll hand off to the existing process.
- If it determines that it should create a window, it will set up a new thread
  for that window, and a message loop on the main thread for handling global
  state, such as hotkeys and the notification icon.

--*/

#pragma once

class AppHost;

class WindowEmperor : public std::enable_shared_from_this<WindowEmperor>
{
public:
    enum UserMessages : UINT
    {
        WM_CLOSE_TERMINAL_WINDOW = WM_USER,
        WM_IDENTIFY_ALL_WINDOWS,
        WM_NOTIFY_FROM_NOTIFICATION_AREA,
    };

    WindowEmperor() noexcept;

    HWND GetMainWindow() const noexcept;
    void ForcePersistence(bool force) noexcept;
    void HandleCommandlineArgs(int nCmdShow);

private:
    struct SummonWindowSelectionArgs
    {
        std::wstring_view WindowName;
        bool OnCurrentDesktop = false;
        winrt::TerminalApp::SummonWindowBehavior SummonBehavior;
        uint64_t WindowID;
    };

    [[nodiscard]] static LRESULT __stdcall _wndProc(HWND window, UINT message, WPARAM wparam, LPARAM lparam) noexcept;

    void _createNewWindowThread(const winrt::TerminalApp::WindowRequestedArgs& args);
    LRESULT _messageHandler(HWND window, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    void _numberOfWindowsChanged(const winrt::Windows::Foundation::IInspectable&, const winrt::Windows::Foundation::IInspectable&);
    safe_void_coroutine _windowIsQuakeWindowChanged(winrt::Windows::Foundation::IInspectable sender, winrt::Windows::Foundation::IInspectable args);
    void _createMessageWindow();
    void _hotkeyPressed(long hotkeyIndex);
    void _registerHotKey(int index, const winrt::Microsoft::Terminal::Control::KeyChord& hotkey) noexcept;
    void _unregisterHotKey(int index) noexcept;
    safe_void_coroutine _setupGlobalHotkeys();
    safe_void_coroutine _close();
    void _finalizeSessionPersistence() const;
    void _checkWindowsForNotificationIcon();
    void _summonWindow();

    wil::unique_hwnd _window;
    winrt::TerminalApp::App _app;
    winrt::Windows::System::DispatcherQueue _dispatcher{ nullptr };
    std::list<std::shared_ptr<::AppHost>> _windows;
    std::vector<winrt::Microsoft::Terminal::Settings::Model::GlobalSummonArgs> _hotkeys;
    NOTIFYICONDATA _notificationIcon{};
    bool _notificationIconShown = false;
    bool _requiresPersistenceCleanupOnExit = false;
    bool _currentSystemThemeIsDark = false;
    bool _quitting{ false };
};
