#pragma once

#include "comtreeview.h"
#include "aboutdlg.h"
#include "comdetailview.h"
#include "ribbon.h"

class ObjectPane : public CPaneContainerImpl<ObjectPane>
{
};

class CMainFrame : public CRibbonFrameWindowImpl<CMainFrame>,
                   public CMessageFilter,
                   public CIdleHandler
{
public:
    DECLARE_FRAME_WND_CLASS(nullptr, IDR_MAINFRAME)

    BOOL PreTranslateMessage(MSG* pMsg) override;
    BOOL OnIdle() override;
    BOOL DefCreate();

    BEGIN_RIBBON_CONTROL_MAP(CMainFrame)
    END_RIBBON_CONTROL_MAP()

    BEGIN_UPDATE_UI_MAP(CMainFrame)
        UPDATE_ELEMENT(ID_VIEW_STATUS_BAR, UPDUI_MENUPOPUP)
        UPDATE_ELEMENT(ID_RELEASE_OBJECT, UPDUI_MENUPOPUP)
        UPDATE_ELEMENT(ID_COPY_GUID, UPDUI_MENUPOPUP)
        UPDATE_ELEMENT(ID_REGEDIT_HERE, UPDUI_MENUPOPUP)
    END_UPDATE_UI_MAP()

BEGIN_MSG_MAP(CMainFrame)
        MSG_WM_CREATE(OnCreate)
        COMMAND_ID_HANDLER2(ID_APP_EXIT, OnFileExit)
        COMMAND_ID_HANDLER2(ID_FILE_OPEN, OnFileOpen)
        COMMAND_ID_HANDLER2(ID_FILE_TYPELIB, OnFileTypeLib)
        COMMAND_ID_HANDLER2(ID_VIEW_STATUS_BAR, OnViewStatusBar)
        COMMAND_ID_HANDLER2(ID_APP_ABOUT, OnAppAbout)
        COMMAND_ID_HANDLER2(ID_RELEASE_OBJECT, OnReleaseObject)
        COMMAND_ID_HANDLER2(ID_COPY_GUID, OnCopyGUID)
        COMMAND_ID_HANDLER2(ID_REGEDIT_HERE, OnRegEditHere)
        REFLECT_NOTIFY_CODE(TVN_ITEMEXPANDING)
        REFLECT_NOTIFY_CODE(TVN_DELETEITEM)
        NOTIFY_CODE_HANDLER_EX(TVN_SELCHANGED, OnTVSelChanged)
        NOTIFY_CODE_HANDLER_EX(NM_RCLICK, OnRClick)
        CHAIN_MSG_MAP(CRibbonFrameWindowImpl<CMainFrame>)
    END_MSG_MAP()

    LRESULT OnTVSelChanged(LPNMHDR pnmhdr);
    LRESULT OnRClick(LPNMHDR pnmh);
    LRESULT OnCreate(LPCREATESTRUCT pcs);
    LRESULT OnFileExit();
    LRESULT OnFileOpen();
    LRESULT OnFileTypeLib();
    LRESULT OnViewStatusBar();
    LRESULT OnAppAbout();
    LRESULT OnReleaseObject();
    LRESULT OnCopyGUID();
    LRESULT OnRegEditHere();

private:
    BOOL IsSelectedInstance() const;
    BOOL IsGUIDSelected() const;

    void AddFileMoniker(LPCTSTR pFilename, LPUNKNOWN pUnk, REFCLSID clsid);
    void AddFileTypeLib(LPCTSTR pFilename, LPTYPELIB pTypeLib);

    ObjectPane m_objectPane;
    ComTreeView m_treeView;
    ComDetailView m_detailView;
    CCommandBarCtrl m_cmdBar;
    CSplitterWindow m_splitter;
};
