// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
#pragma once
#include "pch.h"
#include "AppHost.h"

class WindowThread : public std::enable_shared_from_this<WindowThread>
{
public:
    WindowThread(winrt::TerminalApp::AppLogic logic, winrt::TerminalApp::WindowRequestedArgs args, std::weak_ptr<WindowEmperor> manager);

    winrt::TerminalApp::TerminalWindow Logic();
    void CreateHost();
    int RunMessagePump();
    void RundownForExit();

    bool KeepWarm();
    void Refrigerate();
    void Microwave(winrt::TerminalApp::WindowRequestedArgs args);

    uint64_t PeasantID();

private:
    winrt::TerminalApp::AppLogic _appLogic{ nullptr };
    winrt::TerminalApp::WindowRequestedArgs _args{ nullptr };
    std::weak_ptr<WindowEmperor> _manager;

    // This is a "shared_ptr", but it should be treated as a unique, owning ptr.
    // It's shared, because there are edge cases in refrigeration where internal
    // co_awaits inside AppHost might resume after we've dtor'd it, and there's
    // no other way for us to let the AppHost know it has passed on.
    std::shared_ptr<::AppHost> _host{ nullptr };

    std::unique_ptr<::IslandWindow> _warmWindow{ nullptr };
    static bool _loggedInteraction;

    int _messagePump();
    void _pumpRemainingXamlMessages();
};
