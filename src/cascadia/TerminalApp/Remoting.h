// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

#pragma once

#include "AttachRequest.g.h"
#include "CommandlineArgs.g.h"
#include "FindTargetWindowArgs.g.h"
#include "RequestReceiveContentArgs.g.h"
#include "SummonWindowBehavior.g.h"
#include "WindowActivatedArgs.g.h"
#include "WindowRequestedArgs.g.h"

namespace winrt::TerminalApp::implementation
{
    struct FindTargetWindowArgs : public FindTargetWindowArgsT<FindTargetWindowArgs>
    {
        WINRT_PROPERTY(winrt::TerminalApp::CommandlineArgs, Args, nullptr);
        WINRT_PROPERTY(int, ResultTargetWindow, -1);
        WINRT_PROPERTY(winrt::hstring, ResultTargetWindowName);

    public:
        FindTargetWindowArgs(winrt::TerminalApp::CommandlineArgs args) :
            _Args{ args } {};
    };

    struct AttachRequest : public AttachRequestT<AttachRequest>
    {
        WINRT_PROPERTY(winrt::hstring, Content);
        WINRT_PROPERTY(uint32_t, TabIndex);

    public:
        AttachRequest(winrt::hstring content,
                      uint32_t tabIndex) :
            _Content{ content },
            _TabIndex{ tabIndex } {};
    };

    struct RequestReceiveContentArgs : RequestReceiveContentArgsT<RequestReceiveContentArgs>
    {
        WINRT_PROPERTY(uint64_t, SourceWindow);
        WINRT_PROPERTY(uint64_t, TargetWindow);
        WINRT_PROPERTY(uint32_t, TabIndex);

    public:
        RequestReceiveContentArgs(const uint64_t src, const uint64_t tgt, const uint32_t tabIndex) :
            _SourceWindow{ src },
            _TargetWindow{ tgt },
            _TabIndex{ tabIndex } {};
    };

    struct CommandlineArgs : public CommandlineArgsT<CommandlineArgs>
    {
    public:
        CommandlineArgs() :
            _args{},
            _cwd{ L"" }
        {
        }

        CommandlineArgs(const winrt::array_view<const winrt::hstring>& args,
                        winrt::hstring currentDirectory,
                        const uint32_t showWindowCommand,
                        winrt::hstring envString) :
            _args{ args.begin(), args.end() },
            _cwd{ currentDirectory },
            _ShowWindowCommand{ showWindowCommand },
            CurrentEnvironment{ envString }
        {
        }

        winrt::hstring CurrentDirectory() { return _cwd; };

        void Commandline(const winrt::array_view<const winrt::hstring>& value);
        winrt::com_array<winrt::hstring> Commandline();

        til::property<winrt::hstring> CurrentEnvironment;

        WINRT_PROPERTY(uint32_t, ShowWindowCommand, SW_NORMAL); // SW_NORMAL is 1, 0 is SW_HIDE

    private:
        winrt::com_array<winrt::hstring> _args;
        winrt::hstring _cwd;
    };

    struct WindowRequestedArgs : public WindowRequestedArgsT<WindowRequestedArgs>
    {
    public:
        WindowRequestedArgs(const winrt::TerminalApp::CommandlineArgs& command) :
            _args{ command.Commandline() },
            _CurrentDirectory{ command.CurrentDirectory() },
            _ShowWindowCommand{ command.ShowWindowCommand() },
            _CurrentEnvironment{ command.CurrentEnvironment() } {};

        WindowRequestedArgs(const winrt::hstring& window, const winrt::hstring& content, const Windows::Foundation::IReference<Windows::Foundation::Rect>& bounds) :
            _Id{ 0u },
            _WindowName{ window },
            _args{},
            _CurrentDirectory{},
            _Content{ content },
            _InitialBounds{ bounds } {};

        void Commandline(const winrt::array_view<const winrt::hstring>& value) { _args = { value.begin(), value.end() }; };
        winrt::com_array<winrt::hstring> Commandline() { return winrt::com_array<winrt::hstring>{ _args.begin(), _args.end() }; }

        WINRT_PROPERTY(uint64_t, Id);
        WINRT_PROPERTY(winrt::hstring, WindowName);
        WINRT_PROPERTY(winrt::hstring, CurrentDirectory);
        WINRT_PROPERTY(winrt::hstring, Content);
        WINRT_PROPERTY(uint32_t, ShowWindowCommand, SW_NORMAL);
        WINRT_PROPERTY(winrt::hstring, CurrentEnvironment);
        WINRT_PROPERTY(Windows::Foundation::IReference<Windows::Foundation::Rect>, InitialBounds);

    private:
        winrt::com_array<winrt::hstring> _args;
    };

    struct SummonWindowBehavior : public SummonWindowBehaviorT<SummonWindowBehavior>
    {
    public:
        SummonWindowBehavior() = default;
        WINRT_PROPERTY(bool, MoveToCurrentDesktop, true);
        WINRT_PROPERTY(bool, ToggleVisibility, true);
        WINRT_PROPERTY(uint32_t, DropdownDuration, 0);
        WINRT_PROPERTY(MonitorBehavior, ToMonitor, MonitorBehavior::ToCurrent);

    public:
        SummonWindowBehavior(const SummonWindowBehavior& other) :
            _MoveToCurrentDesktop{ other.MoveToCurrentDesktop() },
            _ToMonitor{ other.ToMonitor() },
            _DropdownDuration{ other.DropdownDuration() },
            _ToggleVisibility{ other.ToggleVisibility() } {};
    };

    struct CompareWindowActivatedArgs
    {
        bool operator()(const WindowActivatedArgs& lhs, const WindowActivatedArgs& rhs) const
        {
            return lhs.ActivatedTime() > rhs.ActivatedTime();
        }
    };
    struct WindowActivatedArgs : public WindowActivatedArgsT<WindowActivatedArgs>
    {
        WINRT_PROPERTY(uint64_t, PeasantID, 0);
        WINRT_PROPERTY(winrt::guid, DesktopID);
        WINRT_PROPERTY(winrt::Windows::Foundation::DateTime, ActivatedTime, {});
        WINRT_PROPERTY(uint64_t, Hwnd, 0);

    public:
        WindowActivatedArgs(uint64_t peasantID,
                            uint64_t hwnd,
                            winrt::guid desktopID,
                            winrt::Windows::Foundation::DateTime timestamp) :
            _PeasantID{ peasantID },
            _Hwnd{ hwnd },
            _DesktopID{ desktopID },
            _ActivatedTime{ timestamp } {};

        WindowActivatedArgs(uint64_t peasantID,
                            winrt::guid desktopID,
                            winrt::Windows::Foundation::DateTime timestamp) :
            WindowActivatedArgs(peasantID, 0, desktopID, timestamp){};

        WindowActivatedArgs(const WindowActivatedArgs& other) :
            _PeasantID{ other.PeasantID() },
            _Hwnd{ other.Hwnd() },
            _DesktopID{ other.DesktopID() },
            _ActivatedTime{ other.ActivatedTime() } {};
    };
}

namespace winrt::TerminalApp::factory_implementation
{
    BASIC_FACTORY(SummonWindowBehavior);
    BASIC_FACTORY(WindowActivatedArgs);
    BASIC_FACTORY(CommandlineArgs);
    BASIC_FACTORY(RequestReceiveContentArgs);
    BASIC_FACTORY(WindowRequestedArgs);
}
