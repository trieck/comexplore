#pragma once

class COMExplorer
{
public:
    virtual ~COMExplorer();

    BOOL Init();
    int Run(HINSTANCE hInstance, LPTSTR lpCmdLine, int nCmdShow);
private:
    HMODULE m_hInstRich = nullptr;
};
