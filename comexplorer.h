#pragma once

class COMExplorer
{
public:
    virtual ~COMExplorer();
    void FinalRelease();

    BOOL Init();
    int Run(HINSTANCE hInstance, LPTSTR lpCmdLine, int nCmdShow);
private:
    HMODULE m_hInstRich = nullptr;
};
