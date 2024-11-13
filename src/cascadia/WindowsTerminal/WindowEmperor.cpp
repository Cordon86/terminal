// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

#include "pch.h"
#include "WindowEmperor.h"

#include <icon.h>

#include "LibraryResources.h"
#include "../inc/WindowingBehavior.h"
#include "../../types/inc/utils.hpp"
#include "../WinRTUtils/inc/WtExeUtils.h"
#include "resource.h"
#include "ScopedResourceLoader.h"

enum class NotificationIconMenuItemAction
{
    FocusTerminal, // Focus the MRU terminal.
    SummonWindow
};

using namespace winrt;
using namespace winrt::Microsoft::Terminal;
using namespace winrt::Microsoft::Terminal::Settings::Model;
using namespace winrt::Windows::Foundation;
using namespace ::Microsoft::Console;
using namespace std::chrono_literals;
using VirtualKeyModifiers = winrt::Windows::System::VirtualKeyModifiers;

#if defined(WT_BRANDING_RELEASE)
#define TERMINAL_MESSAGE_CLASS_NAME L"Windows Terminal Release"
#elif defined(WT_BRANDING_PREVIEW)
#define TERMINAL_MESSAGE_CLASS_NAME L"Windows Terminal Preview"
#elif defined(WT_BRANDING_CANARY)
#define TERMINAL_MESSAGE_CLASS_NAME L"Windows Terminal Canary"
#else
#define TERMINAL_MESSAGE_CLASS_NAME L"Windows Terminal Dev"
#endif

static constexpr ULONG_PTR TERMINAL_HANDOFF_MAGIC = 0x5445524d494e414c; // 'TERMINAL'

extern "C" IMAGE_DOS_HEADER __ImageBase;

static std::vector<winrt::hstring> buildArgsFromCommandline(const wchar_t* commandLine)
{
    int argc = 0;
    const wil::unique_hlocal_ptr<LPWSTR> argv{ CommandLineToArgvW(commandLine, &argc) };

    std::vector<winrt::hstring> args;
    for (int i = 0; i < argc; i++)
    {
        args.emplace_back(argv.get()[i]);
    }
    return args;
}

// Returns the length of a double-null encoded string *excluding* the trailing double-null character.
static std::wstring_view stringFromDoubleNullTerminated(const wchar_t* beg)
{
    auto end = beg;

    for (; *end; end += wcsnlen(end, SIZE_T_MAX) + 1)
    {
    }

    return { beg, end };
}

static void serializeUint32(std::vector<uint8_t>& out, uint32_t value)
{
    out.insert(out.end(), reinterpret_cast<const uint8_t*>(&value), reinterpret_cast<const uint8_t*>(&value) + sizeof(value));
}

static const uint8_t* deserializeUint32(const uint8_t* it, const uint8_t* end, uint32_t& val)
{
    if (static_cast<size_t>(end - it) < sizeof(uint32_t))
    {
        throw std::out_of_range("Not enough data for uint32_t");
    }
    val = *reinterpret_cast<const uint32_t*>(it);
    return it + sizeof(uint32_t);
}

// Writes an uint32_t length prefix, followed by the string data, to the output vector.
static void serializeString(std::vector<uint8_t>& out, std::wstring_view str)
{
    const auto ptr = reinterpret_cast<const uint8_t*>(str.data());
    const auto len = gsl::narrow<uint32_t>(str.size());
    serializeUint32(out, len);
    out.insert(out.end(), ptr, ptr + len * sizeof(wchar_t));
}

// Parses the next string from the input iterator.
// Returns an iterator that points past it.
// Performs bounds-checks and throws std::out_of_range if the input is malformed.
static const uint8_t* deserializeString(const uint8_t* it, const uint8_t* end, std::wstring_view& str)
{
    uint32_t len;
    it = deserializeUint32(it, end, len);

    if (static_cast<size_t>(end - it) < len * sizeof(wchar_t))
    {
        throw std::out_of_range("Not enough data for string content");
    }

    str = { reinterpret_cast<const wchar_t*>(it), len };
    return it + len * sizeof(wchar_t);
}

struct Handoff
{
    std::wstring_view args;
    std::wstring_view env;
    std::wstring_view cwd;
    uint32_t show;
};

static std::vector<uint8_t> serializeHandoffPayload(int nCmdShow)
{
    const auto args = GetCommandLineW();
    const wil::unique_environstrings_ptr envMem{ GetEnvironmentStringsW() };
    const auto env = stringFromDoubleNullTerminated(envMem.get());
    const auto cwd = wil::GetCurrentDirectoryW<std::wstring>();

    std::vector<uint8_t> out;
    serializeString(out, args);
    serializeString(out, env);
    serializeString(out, cwd);
    serializeUint32(out, static_cast<uint32_t>(nCmdShow));
    return out;
}

static Handoff deserializeHandoffPayload(const uint8_t* beg, const uint8_t* end)
{
    Handoff result{};
    auto it = beg;
    it = deserializeString(it, end, result.args);
    it = deserializeString(it, end, result.env);
    it = deserializeString(it, end, result.cwd);
    it = deserializeUint32(it, end, result.show);
    return result;
}

static wil::unique_mutex acquireMutexOrAttemptHandoff(int nCmdShow)
{
    // If the process that owns the mutex has not finished creating the window yet,
    // FindWindowW will return nullptr. We'll retry a few times just in case.
    // At the 1.5x growth rate, this will retry up to ~30s in total.
    for (DWORD sleep = 50; sleep < 10000; sleep += sleep / 2)
    {
        wil::unique_mutex mutex{ CreateMutexW(nullptr, TRUE, TERMINAL_MESSAGE_CLASS_NAME) };
        if (GetLastError() == ERROR_ALREADY_EXISTS)
        {
            mutex.reset();
        }
        if (mutex)
        {
            return mutex;
        }

        // I found that FindWindow() with no other filters is substantially faster than
        // using FindWindowEx() and restricting the search to just HWND_MESSAGE windows.
        // In both cases it's quite fast though at ~1us/op vs ~3us/op (good old Win32 <3).
        if (const auto hwnd = FindWindowW(TERMINAL_MESSAGE_CLASS_NAME, nullptr))
        {
            auto payload = serializeHandoffPayload(nCmdShow);
            const COPYDATASTRUCT cds{
                .dwData = TERMINAL_HANDOFF_MAGIC,
                .cbData = gsl::narrow<DWORD>(payload.size()),
                .lpData = payload.data(),
            };
            if (SendMessageTimeoutW(hwnd, WM_COPYDATA, 0, reinterpret_cast<LPARAM>(&cds), SMTO_ABORTIFHUNG | SMTO_ERRORONEXIT, 10000, nullptr))
            {
                return {};
            }
        }

        Sleep(sleep);
    }

    return {};
}

WindowEmperor::WindowEmperor() noexcept
{
    _dispatcher = winrt::Windows::System::DispatcherQueue::GetForCurrentThread();
}

void WindowEmperor::HandleCommandlineArgs(int nCmdShow)
{
    _app.Logic().ReloadSettings();

    wil::unique_mutex mutex;
    // The "isolatedMode" setting was introduced as an "escape hatch" during the initial multi-window architecture.
    // It's not really a feature that any user asked for, so if it becomes an issue it could be removed.
    // We could consider removing it to dramatically reduce the cost of the `wt` command.
    if (!_app.Logic().IsolatedMode())
    {
        mutex = acquireMutexOrAttemptHandoff(nCmdShow);
        if (!mutex)
        {
            return;
        }
    }

    _createMessageWindow();
    _setupGlobalHotkeys();

    // When the settings change, we'll want to update our global hotkeys and our
    // notification icon based on the new settings.
    _app.Logic().SettingsChanged([this](auto&&, const TerminalApp::SettingsLoadEventArgs& args) {
        if (SUCCEEDED(args.Result()))
        {
            _setupGlobalHotkeys();
            _checkWindowsForNotificationIcon();
        }
    });

    // On startup, immediately check if we need to show the notification icon.
    _checkWindowsForNotificationIcon();

    _currentSystemThemeIsDark = Theme::IsSystemInDarkTheme();

    // If a previous session of Windows Terminal stored buffer_*.txt files, then we need to clean all those up on exit
    // that aren't needed anymore, even if the user disabled the ShouldUsePersistedLayout() setting in the meantime.
    {
        const auto state = ApplicationState::SharedInstance();
        const auto layouts = state.PersistedWindowLayouts();
        _requiresPersistenceCleanupOnExit = layouts && layouts.Size() > 0;
    }

    {
        const auto args = buildArgsFromCommandline(GetCommandLineW());
        const wil::unique_environstrings_ptr envMem{ GetEnvironmentStringsW() };
        const auto env = stringFromDoubleNullTerminated(envMem.get());
        const auto cwd = wil::GetCurrentDirectoryW<std::wstring>();
        const winrt::TerminalApp::CommandlineArgs eventArgs{ args, cwd, gsl::narrow_cast<uint32_t>(nCmdShow), std::move(env) };
        _createNewWindowThread(winrt::TerminalApp::WindowRequestedArgs{ eventArgs });
    }

    // ALWAYS change the _real_ CWD of the Terminal to system32, so that we
    // don't lock the directory we were spawned in.
    if (std::wstring system32; SUCCEEDED_LOG(wil::GetSystemDirectoryW(system32)))
    {
        LOG_IF_WIN32_BOOL_FALSE(SetCurrentDirectoryW(system32.c_str()));
    }

    MSG message{};
    while (GetMessageW(&message, nullptr, 0, 0))
    {
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }

    _finalizeSessionPersistence();

    // There's a mysterious crash in XAML on Windows 10 if you just let _app get destroyed (GH#15410).
    // We also need to ensure that all UI threads exit before WindowEmperor leaves the scope on the main thread (MSFT:46744208).
    // Both problems can be solved and the shutdown accelerated by using TerminateProcess.
    // std::exit(), etc., cannot be used here, because those use ExitProcess for unpackaged applications.
    TerminateProcess(GetCurrentProcess(), gsl::narrow_cast<UINT>(0));
    __assume(false);
}

void WindowEmperor::_createNewWindowThread(const winrt::TerminalApp::WindowRequestedArgs& args)
{
    std::shared_ptr<WindowThread> window;

    // FIRST: Attempt to reheat an existing window that we refrigerated for
    // later. If we have an existing unused window, then we don't need to create
    // a new WindowThread & HWND for this request.
    // Add a scope to minimize lock duration.
    {
        const auto fridge = _oldWindows.lock();
        if (!fridge->empty())
        {
            // Look at that, a refrigerated thread ready to be used. Let's use that!
            window = std::move(fridge->back());
            fridge->pop_back();
        }
    }
    if (window)
    {
        window->Microwave(args);
        // This will unblock the event we're waiting on in KeepWarm, and the
        // window thread (started below) will continue through it's loop
        return;
    }

    std::thread t([window = std::make_shared<WindowThread>(_app.Logic(), args, weak_from_this()), weakThis = weak_from_this()]() {
        try
        {
            window->CreateHost();

            if (auto self{ weakThis.lock() })
            {
                window->Logic().IsQuakeWindowChanged({ &*self, &WindowEmperor::_windowIsQuakeWindowChanged });
            }

            while (window->KeepWarm())
            {
                // Now that the window is ready to go, we can add it to our list of windows,
                // because we know it will be well behaved.
                //
                // Be sure to only modify the list of windows under lock.

                if (auto self{ weakThis.lock() })
                {
                    auto lockedWindows{ self->_windows.lock() };
                    lockedWindows->push_back(window);
                }
                auto removeWindow = wil::scope_exit([&]() {
                    if (auto self{ weakThis.lock() })
                    {
                        self->_removeWindow(window->PeasantID());
                    }
                });

                auto decrementWindowCount = wil::scope_exit([&]() {
                    if (auto self{ weakThis.lock() })
                    {
                        self->_decrementWindowCount();
                    }
                });

                window->RunMessagePump();

                // Manually trigger the cleanup callback. This will ensure that we
                // remove the window from our list of windows, before we release the
                // AppHost (and subsequently, the host's Logic() member that we use
                // elsewhere).
                removeWindow.reset();

                // On Windows 11, we DONT want to refrigerate the window. There,
                // we can just close it like normal. Break out of the loop, so
                // we don't try to put this window in the fridge.
                if (Utils::IsWindows11())
                {
                    decrementWindowCount.reset();
                    break;
                }
                else
                {
                    window->Refrigerate();
                    decrementWindowCount.reset();

                    if (auto self{ weakThis.lock() })
                    {
                        auto fridge{ self->_oldWindows.lock() };
                        fridge->push_back(window);
                    }
                }
            }

            // Now that we no longer care about this thread's window, let it
            // release it's app host and flush the rest of the XAML queue.
            window->RundownForExit();
        }
        CATCH_LOG()
    });
    LOG_IF_FAILED(SetThreadDescription(t.native_handle(), L"Window Thread"));
    t.detach();
}

void WindowEmperor::_removeWindow(uint64_t senderID)
{
    auto lockedWindows{ _windows.lock() };

    // find the window in _windows who's peasant's Id matches the peasant's Id
    // and remove it
    std::erase_if(*lockedWindows, [&](const auto& w) {
        return w->PeasantID() == senderID;
    });
}

void WindowEmperor::_decrementWindowCount()
{
    // When we run out of windows, exit our process if and only if:
    // * We're not allowed to run headless OR
    // * we've explicitly been told to "quit", which should fully exit the Terminal.
    const bool quitWhenLastWindowExits{ !_app.Logic().AllowHeadless() };
    const bool noMoreWindows{ _windowThreadInstances.fetch_sub(1, std::memory_order_relaxed) == 1 };
    if (noMoreWindows &&
        (_quitting || quitWhenLastWindowExits))
    {
        _close();
    }
}

// sender and args are always nullptr
void WindowEmperor::_numberOfWindowsChanged(const winrt::Windows::Foundation::IInspectable&,
                                            const winrt::Windows::Foundation::IInspectable&)
{
    // If we closed out the quake window, and don't otherwise need the tray
    // icon, let's get rid of it.
    _checkWindowsForNotificationIcon();
}

#pragma region WindowProc

static WindowEmperor* GetThisFromHandle(HWND const window) noexcept
{
    const auto data = GetWindowLongPtrW(window, GWLP_USERDATA);
    return reinterpret_cast<WindowEmperor*>(data);
}

[[nodiscard]] LRESULT __stdcall WindowEmperor::_wndProc(HWND const window, UINT const message, WPARAM const wparam, LPARAM const lparam) noexcept
{
    WINRT_ASSERT(window);

    if (WM_NCCREATE == message)
    {
        auto cs = reinterpret_cast<CREATESTRUCT*>(lparam);
        WindowEmperor* that = static_cast<WindowEmperor*>(cs->lpCreateParams);
        WINRT_ASSERT(that);
        WINRT_ASSERT(!that->_window);
        that->_window = wil::unique_hwnd(window);
        SetWindowLongPtr(that->_window.get(), GWLP_USERDATA, reinterpret_cast<LONG_PTR>(that));
    }
    else if (WindowEmperor* that = GetThisFromHandle(window))
    {
        return that->_messageHandler(window, message, wparam, lparam);
    }

    return DefWindowProcW(window, message, wparam, lparam);
}
void WindowEmperor::_createMessageWindow()
{
    const auto instance = reinterpret_cast<HINSTANCE>(&__ImageBase);

    const WNDCLASS wc{
        .lpfnWndProc = &_wndProc,
        .hInstance = instance,
        .hIcon = LoadIconW(instance, MAKEINTRESOURCEW(IDI_APPICON)),
        .lpszClassName = TERMINAL_MESSAGE_CLASS_NAME,
    };
    RegisterClassW(&wc);

    WINRT_VERIFY(CreateWindowExW(
        /* dwExStyle    */ 0,
        /* lpClassName  */ TERMINAL_MESSAGE_CLASS_NAME,
        /* lpWindowName */ L"Windows Terminal",
        /* dwStyle      */ 0,
        /* X            */ 0,
        /* Y            */ 0,
        /* nWidth       */ 0,
        /* nHeight      */ 0,
        /* hWndParent   */ HWND_MESSAGE,
        /* hMenu        */ nullptr,
        /* hInstance    */ instance,
        /* lpParam      */ this));

    _notificationIcon.cbSize = sizeof(NOTIFYICONDATA);
    _notificationIcon.hWnd = _window.get();
    _notificationIcon.uID = 1;
    _notificationIcon.uFlags = NIF_MESSAGE | NIF_TIP | NIF_SHOWTIP | NIF_ICON;
    _notificationIcon.uCallbackMessage = CM_NOTIFY_FROM_NOTIFICATION_AREA;
    _notificationIcon.hIcon = static_cast<HICON>(GetActiveAppIconHandle(true));
    _notificationIcon.uVersion = NOTIFYICON_VERSION_4;

    // AppName happens to be in the ContextMenu's Resources, see GH#12264
    ScopedResourceLoader loader{ L"TerminalApp/ContextMenu" };
    const auto appNameLoc = loader.GetLocalizedString(L"AppName");
    StringCchCopy(_notificationIcon.szTip, ARRAYSIZE(_notificationIcon.szTip), appNameLoc.c_str());
}

LRESULT WindowEmperor::_messageHandler(HWND window, UINT const message, WPARAM const wParam, LPARAM const lParam) noexcept
{
    // use C++11 magic statics to make sure we only do this once.
    // This won't change over the lifetime of the application
    static const UINT WM_TASKBARCREATED = []() { return RegisterWindowMessageW(L"TaskbarCreated"); }();

    try
    {
        switch (message)
        {
        case WM_SETTINGCHANGE:
        {
            // Currently, we only support checking when the OS theme changes. In
            // that case, wParam is 0. Re-evaluate when we decide to reload env vars
            // (GH#1125)
            if (wParam == 0 && lParam != 0)
            {
                // ImmersiveColorSet seems to be the notification that the OS theme
                // changed. If that happens, let the app know, so it can hot-reload
                // themes, color schemes that might depend on the OS theme
                if (wcscmp(reinterpret_cast<const wchar_t*>(lParam), L"ImmersiveColorSet") == 0)
                {
                    try
                    {
                        // GH#15732: Don't update the settings, unless the theme
                        // _actually_ changed. ImmersiveColorSet gets sent more often
                        // than just on a theme change. It notably gets sent when the PC
                        // is locked, or the UAC prompt opens.
                        const auto isCurrentlyDark = Theme::IsSystemInDarkTheme();
                        if (isCurrentlyDark != _currentSystemThemeIsDark)
                        {
                            _currentSystemThemeIsDark = isCurrentlyDark;
                            _app.Logic().ReloadSettings();
                        }
                    }
                    CATCH_LOG();
                }
            }
            return 0;
        }
        case WM_COPYDATA:
        {
            const auto cds = reinterpret_cast<COPYDATASTRUCT*>(lParam);
            if (cds->dwData == TERMINAL_HANDOFF_MAGIC)
            {
                try
                {
                    const auto handoff = deserializeHandoffPayload(static_cast<const uint8_t*>(cds->lpData), static_cast<const uint8_t*>(cds->lpData) + cds->cbData);
                    const winrt::hstring args{ handoff.args };
                    const winrt::hstring env{ handoff.env };
                    const winrt::hstring cwd{ handoff.cwd };
                    const auto argv = buildArgsFromCommandline(args.c_str());
                    const winrt::TerminalApp::CommandlineArgs eventArgs{ argv, cwd, gsl::narrow_cast<uint32_t>(handoff.show), env };
                    _createNewWindowThread(winrt::TerminalApp::WindowRequestedArgs{ eventArgs });
                }
                CATCH_LOG();
            }
            return 0;
        }
        case WM_HOTKEY:
        {
            _hotkeyPressed(static_cast<long>(wParam));
            return 0;
        }
        case CM_NOTIFY_FROM_NOTIFICATION_AREA:
        {
            switch (LOWORD(lParam))
            {
            case NIN_SELECT:
            case NIN_KEYSELECT:
            {
                SummonWindowSelectionArgs args{};
                args.SummonBehavior.MoveToCurrentDesktop(false);
                args.SummonBehavior.ToMonitor(winrt::TerminalApp::MonitorBehavior::InPlace);
                args.SummonBehavior.ToggleVisibility(false);
                //_manager.SummonWindow(args);
                return 0;
            }
            case WM_CONTEXTMENU:
            {
                if (const auto hMenu = CreatePopupMenu())
                {
                    MENUINFO mi{};
                    mi.cbSize = sizeof(MENUINFO);
                    mi.fMask = MIM_STYLE | MIM_APPLYTOSUBMENUS | MIM_MENUDATA;
                    mi.dwStyle = MNS_NOTIFYBYPOS;
                    mi.dwMenuData = NULL;
                    SetMenuInfo(hMenu, &mi);

                    // Focus Current Terminal Window
                    AppendMenu(hMenu, MF_STRING, gsl::narrow<UINT_PTR>(NotificationIconMenuItemAction::FocusTerminal), RS_(L"NotificationIconFocusTerminal").c_str());
                    AppendMenu(hMenu, MF_SEPARATOR, 0, L"");

                    // Submenu for Windows
                    if (const auto submenu = CreatePopupMenu())
                    {
                        /*for (const auto& p : peasants)
                        {
                            std::wstringstream displayText;
                            displayText << L"#" << p.Id;

                            if (!p.TabTitle.empty())
                            {
                                displayText << L": " << std::wstring_view{ p.TabTitle };
                            }

                            if (!p.Name.empty())
                            {
                                displayText << L" [" << std::wstring_view{ p.Name } << L"]";
                            }

                            AppendMenu(submenu, MF_STRING, gsl::narrow<UINT_PTR>(p.Id), displayText.str().c_str());
                        }*/

                        MENUINFO submenuInfo{};
                        submenuInfo.cbSize = sizeof(MENUINFO);
                        submenuInfo.fMask = MIM_MENUDATA;
                        submenuInfo.dwStyle = MNS_NOTIFYBYPOS;
                        submenuInfo.dwMenuData = (UINT_PTR)NotificationIconMenuItemAction::SummonWindow;
                        SetMenuInfo(submenu, &submenuInfo);

                        AppendMenu(hMenu, MF_POPUP, (UINT_PTR)submenu, RS_(L"NotificationIconWindowSubmenu").c_str());
                    }

                    // We'll need to set our window to the foreground before calling
                    // TrackPopupMenuEx or else the menu won't dismiss when clicking away.
                    SetForegroundWindow(window);

                    // User can select menu items with the left and right buttons.
                    const auto rightAlign = GetSystemMetrics(SM_MENUDROPALIGNMENT) != 0;
                    const UINT uFlags = TPM_RIGHTBUTTON | (rightAlign ? TPM_RIGHTALIGN : TPM_LEFTALIGN);
                    TrackPopupMenuEx(hMenu, uFlags, GET_X_LPARAM(wParam), GET_Y_LPARAM(wParam), window, nullptr);
                }

                return 0;
            }
            default:
                break;
            }
            break;
        }
        case WM_MENUCOMMAND:
        {
            const auto menu = (HMENU)lParam;
            const auto menuItemIndex = LOWORD(wParam);

            MENUINFO mi{};
            mi.cbSize = sizeof(MENUINFO);
            mi.fMask = MIM_MENUDATA;
            GetMenuInfo(menu, &mi);

            if (mi.dwMenuData)
            {
                if (gsl::narrow<NotificationIconMenuItemAction>(mi.dwMenuData) == NotificationIconMenuItemAction::SummonWindow)
                {
                    SummonWindowSelectionArgs args{};
                    args.WindowID = GetMenuItemID(menu, menuItemIndex);
                    args.SummonBehavior.ToggleVisibility(false);
                    args.SummonBehavior.MoveToCurrentDesktop(false);
                    args.SummonBehavior.ToMonitor(winrt::TerminalApp::MonitorBehavior::InPlace);
                    //_manager.SummonWindow(args);
                    return 0;
                }
            }

            // Now check the menu item itself for an action.
            const auto action = gsl::narrow<NotificationIconMenuItemAction>(GetMenuItemID(menu, menuItemIndex));
            switch (action)
            {
            case NotificationIconMenuItemAction::FocusTerminal:
            {
                SummonWindowSelectionArgs args{};
                args.SummonBehavior.ToggleVisibility(false);
                args.SummonBehavior.MoveToCurrentDesktop(false);
                args.SummonBehavior.ToMonitor(winrt::TerminalApp::MonitorBehavior::InPlace);
                //_manager.SummonWindow(args);
                break;
            }
            default:
                break;
            }
            return 0;
        }
        default:
        {
            // We'll want to receive this message when explorer.exe restarts
            // so that we can re-add our icon to the notification area.
            // This unfortunately isn't a switch case because we register the
            // message at runtime.
            if (message == WM_TASKBARCREATED)
            {
                _checkWindowsForNotificationIcon();
                return 0;
            }
        }
        }
    }
    catch (...)
    {
        LOG_CAUGHT_EXCEPTION();
    }

    return DefWindowProcW(window, message, wParam, lParam);
}

// Close the Terminal application. This will exit the main thread for the
// emperor itself. We should probably only ever be called when we have no
// windows left, and we don't want to keep running anymore. This will discard
// all our refrigerated windows. If we try to use XAML on Windows 10 after this,
// we'll undoubtedly crash.
safe_void_coroutine WindowEmperor::_close()
{
    // Important! Switch back to the main thread for the emperor. That way, the
    // quit will go to the emperor's message pump.
    co_await wil::resume_foreground(_dispatcher);
    PostQuitMessage(0);
}

void WindowEmperor::_finalizeSessionPersistence() const
{
    const auto state = ApplicationState::SharedInstance();

    // Ensure to write the state.json before we TerminateProcess()
    state.Flush();

    if (!_requiresPersistenceCleanupOnExit)
    {
        return;
    }

    // Get the "buffer_{guid}.txt" files that we expect to be there
    std::unordered_set<winrt::guid> sessionIds;
    if (const auto layouts = state.PersistedWindowLayouts())
    {
        for (const auto& windowLayout : layouts)
        {
            for (const auto& actionAndArgs : windowLayout.TabLayout())
            {
                const auto args = actionAndArgs.Args();
                NewTerminalArgs terminalArgs{ nullptr };

                if (const auto tabArgs = args.try_as<NewTabArgs>())
                {
                    terminalArgs = tabArgs.ContentArgs().try_as<NewTerminalArgs>();
                }
                else if (const auto paneArgs = args.try_as<SplitPaneArgs>())
                {
                    terminalArgs = paneArgs.ContentArgs().try_as<NewTerminalArgs>();
                }

                if (terminalArgs)
                {
                    sessionIds.emplace(terminalArgs.SessionId());
                }
            }
        }
    }

    // Remove the "buffer_{guid}.txt" files that shouldn't be there
    // e.g. "buffer_FD40D746-163E-444C-B9B2-6A3EA2B26722.txt"
    {
        const std::filesystem::path settingsDirectory{ std::wstring_view{ CascadiaSettings::SettingsDirectory() } };
        const auto filter = settingsDirectory / L"buffer_*";
        WIN32_FIND_DATAW ffd;

        // This could also use std::filesystem::directory_iterator.
        // I was just slightly bothered by how it doesn't have a O(1) .filename()
        // function, even though the underlying Win32 APIs provide it for free.
        // Both work fine.
        const wil::unique_hfind handle{ FindFirstFileExW(filter.c_str(), FindExInfoBasic, &ffd, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH) };
        if (!handle)
        {
            return;
        }

        do
        {
            const auto nameLen = wcsnlen_s(&ffd.cFileName[0], ARRAYSIZE(ffd.cFileName));
            const std::wstring_view name{ &ffd.cFileName[0], nameLen };

            if (nameLen != 47)
            {
                continue;
            }

            wchar_t guidStr[39];
            guidStr[0] = L'{';
            memcpy(&guidStr[1], name.data() + 7, 36 * sizeof(wchar_t));
            guidStr[37] = L'}';
            guidStr[38] = L'\0';

            const auto id = Utils::GuidFromString(&guidStr[0]);
            if (!sessionIds.contains(id))
            {
                std::filesystem::remove(settingsDirectory / name);
            }
        } while (FindNextFileW(handle.get(), &ffd));
    }
}

#pragma endregion
#pragma region GlobalHotkeys

void WindowEmperor::_hotkeyPressed(const long hotkeyIndex)
{
    if (hotkeyIndex < 0 || static_cast<size_t>(hotkeyIndex) > _hotkeys.size())
    {
        return;
    }

    const auto& summonArgs = til::at(_hotkeys, hotkeyIndex);
    SummonWindowSelectionArgs args{ summonArgs.Name() };

    // desktop:any - MoveToCurrentDesktop=false, OnCurrentDesktop=false
    // desktop:toCurrent - MoveToCurrentDesktop=true, OnCurrentDesktop=false
    // desktop:onCurrent - MoveToCurrentDesktop=false, OnCurrentDesktop=true
    args.OnCurrentDesktop = summonArgs.Desktop() == DesktopBehavior::OnCurrent;
    args.SummonBehavior.MoveToCurrentDesktop(summonArgs.Desktop() == DesktopBehavior::ToCurrent);
    args.SummonBehavior.ToggleVisibility(summonArgs.ToggleVisibility());
    args.SummonBehavior.DropdownDuration(summonArgs.DropdownDuration());

    switch (summonArgs.Monitor())
    {
    case MonitorBehavior::Any:
        args.SummonBehavior.ToMonitor(TerminalApp::MonitorBehavior::InPlace);
        break;
    case MonitorBehavior::ToCurrent:
        args.SummonBehavior.ToMonitor(TerminalApp::MonitorBehavior::ToCurrent);
        break;
    case MonitorBehavior::ToMouse:
        args.SummonBehavior.ToMonitor(TerminalApp::MonitorBehavior::ToMouse);
        break;
    }

    ////_manager.SummonWindow(args);
    //if (args.FoundMatch())
    {
        return;
    }

    hstring argv[2];
    argv[0] = L"wt";
    argv[1] = summonArgs.Name();
    if (argv[1].empty())
    {
        argv[1] = L"new";
    }

    const wil::unique_environstrings_ptr envMem{ GetEnvironmentStringsW() };
    const auto env = stringFromDoubleNullTerminated(envMem.get());
    const auto cwd = wil::GetCurrentDirectoryW<std::wstring>();
    const TerminalApp::CommandlineArgs eventArgs{ argv, cwd, SW_SHOWDEFAULT, std::move(env) };
    _createNewWindowThread(TerminalApp::WindowRequestedArgs{ eventArgs });
}

bool WindowEmperor::_registerHotKey(const int index, const winrt::Microsoft::Terminal::Control::KeyChord& hotkey) noexcept
{
    const auto vkey = hotkey.Vkey();
    auto hotkeyFlags = MOD_NOREPEAT;
    {
        const auto modifiers = hotkey.Modifiers();
        WI_SetFlagIf(hotkeyFlags, MOD_WIN, WI_IsFlagSet(modifiers, VirtualKeyModifiers::Windows));
        WI_SetFlagIf(hotkeyFlags, MOD_ALT, WI_IsFlagSet(modifiers, VirtualKeyModifiers::Menu));
        WI_SetFlagIf(hotkeyFlags, MOD_CONTROL, WI_IsFlagSet(modifiers, VirtualKeyModifiers::Control));
        WI_SetFlagIf(hotkeyFlags, MOD_SHIFT, WI_IsFlagSet(modifiers, VirtualKeyModifiers::Shift));
    }

    // TODO GH#8888: We should display a warning of some kind if this fails.
    // This can fail if something else already bound this hotkey.
    const auto result = ::RegisterHotKey(_window.get(), index, hotkeyFlags, vkey);
    LOG_LAST_ERROR_IF(!result);
    TraceLoggingWrite(g_hWindowsTerminalProvider,
                      "RegisterHotKey",
                      TraceLoggingDescription("Emitted when setting hotkeys"),
                      TraceLoggingInt64(index, "index", "the index of the hotkey to add"),
                      TraceLoggingUInt64(vkey, "vkey", "the key"),
                      TraceLoggingUInt64(WI_IsFlagSet(hotkeyFlags, MOD_WIN), "win", "is WIN in the modifiers"),
                      TraceLoggingUInt64(WI_IsFlagSet(hotkeyFlags, MOD_ALT), "alt", "is ALT in the modifiers"),
                      TraceLoggingUInt64(WI_IsFlagSet(hotkeyFlags, MOD_CONTROL), "control", "is CONTROL in the modifiers"),
                      TraceLoggingUInt64(WI_IsFlagSet(hotkeyFlags, MOD_SHIFT), "shift", "is SHIFT in the modifiers"),
                      TraceLoggingBool(result, "succeeded", "true if we succeeded"),
                      TraceLoggingLevel(WINEVENT_LEVEL_VERBOSE),
                      TraceLoggingKeyword(TIL_KEYWORD_TRACE));

    return result;
}

// Method Description:
// - Call UnregisterHotKey once for each previously registered hotkey.
// Return Value:
// - <none>
void WindowEmperor::_unregisterHotKey(const int index) noexcept
{
    TraceLoggingWrite(
        g_hWindowsTerminalProvider,
        "UnregisterHotKey",
        TraceLoggingDescription("Emitted when clearing previously set hotkeys"),
        TraceLoggingInt64(index, "index", "the index of the hotkey to remove"),
        TraceLoggingLevel(WINEVENT_LEVEL_VERBOSE),
        TraceLoggingKeyword(TIL_KEYWORD_TRACE));

    LOG_IF_WIN32_BOOL_FALSE(::UnregisterHotKey(_window.get(), index));
}

safe_void_coroutine WindowEmperor::_setupGlobalHotkeys()
{
    // The hotkey MUST be registered on the main thread. It will fail otherwise!
    co_await wil::resume_foreground(_dispatcher);

    if (!_window)
    {
        // MSFT:36797001 There's a surprising number of hits of this callback
        // getting triggered during teardown. As a best practice, we really
        // should make sure _window exists before accessing it on any coroutine.
        // We might be getting called back after the app already began getting
        // cleaned up.
        co_return;
    }
    // Unregister all previously registered hotkeys.
    //
    // RegisterHotKey(), will not unregister hotkeys automatically.
    // If a hotkey with a given HWND and ID combination already exists
    // then a duplicate one will be added, which we don't want.
    // (Additionally we want to remove hotkeys that were removed from the settings.)
    for (auto i = 0, count = gsl::narrow_cast<int>(_hotkeys.size()); i < count; ++i)
    {
        _unregisterHotKey(i);
    }

    _hotkeys.clear();

    // Re-register all current hotkeys.
    for (const auto& [keyChord, cmd] : _app.Logic().GlobalHotkeys())
    {
        if (auto summonArgs = cmd.ActionAndArgs().Args().try_as<Settings::Model::GlobalSummonArgs>())
        {
            auto index = gsl::narrow_cast<int>(_hotkeys.size());
            const auto succeeded = _registerHotKey(index, keyChord);

            TraceLoggingWrite(g_hWindowsTerminalProvider,
                              "AppHost_setupGlobalHotkey",
                              TraceLoggingDescription("Emitted when setting a single hotkey"),
                              TraceLoggingInt64(index, "index", "the index of the hotkey to add"),
                              TraceLoggingWideString(cmd.Name().c_str(), "name", "the name of the command"),
                              TraceLoggingBoolean(succeeded, "succeeded", "true if we succeeded"),
                              TraceLoggingLevel(WINEVENT_LEVEL_VERBOSE),
                              TraceLoggingKeyword(TIL_KEYWORD_TRACE));
            _hotkeys.emplace_back(summonArgs);
        }
    }
}

#pragma endregion

#pragma region NotificationIcon

void WindowEmperor::_checkWindowsForNotificationIcon()
{
    // We need to check some conditions to show the notification icon.
    //
    // * If there's a Quake window somewhere, we'll want to keep the
    //   notification icon.
    // * There's two settings - MinimizeToNotificationArea and
    //   AlwaysShowNotificationIcon. If either one of them are true, we want to
    //   make sure there's a notification icon.
    //
    // If both are false, we want to remove our icon from the notification area.
    // When we remove our icon from the notification area, we'll also want to
    // re-summon any hidden windows, but right now we're not keeping track of
    // who's hidden, so just summon them all. Tracking the work to do a "summon
    // all minimized" in GH#10448
    //
    // To avoid races between us thinking the settings updated, and the windows
    // themselves getting the new settings, only ask the app logic for the
    // RequestsTrayIcon setting value, and combine that with the result of each
    // window (which won't change during a settings reload).
    bool needsIcon = _app.Logic().RequestsTrayIcon();
    {
        auto windows{ _windows.lock_shared() };
        for (const auto& _windowThread : *windows)
        {
            needsIcon |= _windowThread->Logic().IsQuakeWindow();
        }
    }

    if (_notificationIconShown == needsIcon)
    {
        return;
    }

    if (needsIcon)
    {
        Shell_NotifyIconW(NIM_ADD, &_notificationIcon);
        Shell_NotifyIconW(NIM_SETVERSION, &_notificationIcon);
    }
    else
    {
        Shell_NotifyIconW(NIM_DELETE, &_notificationIcon);
        // If we no longer want the tray icon, but we did have one, then quick
        // re-summon all our windows, so they don't get lost when the icon
        // disappears forever.
        //_manager.SummonAllWindows();
    }

    _notificationIconShown = needsIcon;
}

#pragma endregion

// A callback to the window's logic to let us know when the window's
// quake mode state changes. We'll use this to check if we need to add
// or remove the notification icon.
safe_void_coroutine WindowEmperor::_windowIsQuakeWindowChanged(winrt::Windows::Foundation::IInspectable sender,
                                                               winrt::Windows::Foundation::IInspectable args)
{
    co_await wil::resume_foreground(this->_dispatcher);
    _checkWindowsForNotificationIcon();
}
