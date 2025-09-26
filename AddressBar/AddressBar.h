#pragma once

#include "stdafx.h"
#include "framework.h"
#include "resource.h"
#include "ClassicExplorer_i.h"
#include "dllmain.h"
#include "util/util.h"

#include <shlobj.h>
#include <shlwapi.h>
#include <shellapi.h>
#include <vector>
#include <array>
#include <memory>

class CAddressBar;

class ExplorerTabDropTarget : public IDropTarget
{
public:
        explicit ExplorerTabDropTarget(CAddressBar *owner);

        // IUnknown
        IFACEMETHODIMP QueryInterface(REFIID riid, void **ppvObject) override;
        ULONG STDMETHODCALLTYPE AddRef() override;
        ULONG STDMETHODCALLTYPE Release() override;

        // IDropTarget
        IFACEMETHODIMP DragEnter(IDataObject *pDataObject, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect) override;
        IFACEMETHODIMP DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect) override;
        IFACEMETHODIMP DragLeave() override;
        IFACEMETHODIMP Drop(IDataObject *pDataObject, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect) override;

private:
        ~ExplorerTabDropTarget() = default;

        LONG m_refCount = 1;
        CAddressBar *m_owner = nullptr;
        CComPtr<IDataObject> m_currentDataObject;
};

class CAddressBar : public CWindowImpl<CAddressBar>
{
private:
        struct UniquePidl
        {
                PIDLIST_ABSOLUTE pidl = nullptr;

                UniquePidl() = default;
                explicit UniquePidl(PIDLIST_ABSOLUTE source)
                {
                        reset(source);
                }

                UniquePidl(const UniquePidl &other)
                {
                        if (other.pidl)
                                pidl = ILCloneFull(other.pidl);
                }

                UniquePidl(UniquePidl &&other) noexcept
                {
                        pidl = other.pidl;
                        other.pidl = nullptr;
                }

                UniquePidl &operator=(const UniquePidl &other)
                {
                        if (this != &other)
                        {
                                reset(other.pidl);
                        }
                        return *this;
                }

                UniquePidl &operator=(UniquePidl &&other) noexcept
                {
                        if (this != &other)
                        {
                                reset();
                                pidl = other.pidl;
                                other.pidl = nullptr;
                        }
                        return *this;
                }

                ~UniquePidl()
                {
                        reset();
                }

                void reset(PIDLIST_ABSOLUTE source = nullptr)
                {
                        if (pidl)
                        {
                                CoTaskMemFree(pidl);
                                pidl = nullptr;
                        }
                        if (source)
                        {
                                pidl = ILCloneFull(source);
                        }
                }
        };

        struct Tab
        {
                UniquePidl pidl;
                std::wstring title;
                RECT bounds = {0};
                bool active = false;
        };

        struct TabGroup
        {
                std::wstring name;
                COLORREF color = RGB(180, 200, 235);
                std::vector<Tab> tabs;
                RECT bounds = {0};
        };

        struct HitTestResult
        {
                bool valid = false;
                bool groupHandle = false;
                int groupIndex = -1;
                int tabIndex = -1;
        };

public:
        DECLARE_WND_CLASS(L"ClassicExplorer.TabBar")

        BEGIN_MSG_MAP(CAddressBar)
                MESSAGE_HANDLER(WM_CREATE, OnCreate)
                MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
                MESSAGE_HANDLER(WM_PAINT, OnPaint)
                MESSAGE_HANDLER(WM_ERASEBKGND, OnEraseBackground)
                MESSAGE_HANDLER(WM_SIZE, OnSize)
                MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
                MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
                MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
                MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
                MESSAGE_HANDLER(WM_CAPTURECHANGED, OnCaptureChanged)
        END_MSG_MAP()

        HWND GetToolbar() const { return m_hWnd; }

        void SetBrowsers(CComPtr<IShellBrowser> pShellBrowser, CComPtr<IWebBrowser2> pWebBrowser);
        HRESULT InitializeTabs();
        void OnExplorerNavigate();
        SIZE GetDesiredSize() const;

        void HandleExternalDragEnter(DWORD keyState, POINTL pt, IDataObject *pDataObject);
        void HandleExternalDragOver(DWORD keyState, POINTL pt);
        void HandleExternalDragLeave();
        void HandleExternalDrop(IDataObject *dataObject, DWORD keyState, POINTL pt);
        bool CanAcceptDataObject(IDataObject *dataObject) const;

private:
        // message handlers
        LRESULT OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);
        LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);
        LRESULT OnPaint(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);
        LRESULT OnEraseBackground(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);
        LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);
        LRESULT OnLButtonDown(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);
        LRESULT OnLButtonUp(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);
        LRESULT OnMouseMove(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);
        LRESULT OnContextMenu(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);
        LRESULT OnCaptureChanged(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled);

        // layout helpers
        void LoadSettings();
        void EnsureDefaultGroup();
        void LayoutTabs();
        void LayoutTabsIfNeeded();
        int CalculateTabWidth(HDC hdc, const std::wstring &text) const;
        void RefreshActiveState();

        // painting helpers
        void DrawBackground(HDC hdc, const RECT &clientRect) const;
        void DrawGroup(HDC hdc, const TabGroup &group) const;
        void DrawTab(HDC hdc, const Tab &tab, COLORREF groupColor) const;
        void DrawGroupHandle(HDC hdc, const TabGroup &group) const;
        void DrawGhost(HDC hdc) const;
        void DrawDropHover(HDC hdc) const;
        static COLORREF AdjustColor(COLORREF color, double factor);

        // tab management
        HRESULT AddTabForLocation(PIDLIST_ABSOLUTE pidl, bool makeActive, bool navigate, COLORREF colorOverride = RGB(180, 200, 235));
        void RemoveEmptyGroups();
        void ActivateTab(int groupIndex, int tabIndex, bool navigate);
        void ActivateTabByPidl(PIDLIST_ABSOLUTE pidl);
        void UpdateActiveTabFromExplorer();
        void CreateNewWindowForTab(const Tab &tab);
        std::wstring GetTabFilesystemPath(const Tab &tab) const;
        void SetGroupColor(int groupIndex, COLORREF color);
        void ShowGroupColorMenu(int groupIndex, POINT screenPoint);
        void ShowContextMenuForTab(int groupIndex, int tabIndex, POINT screenPoint);
        void EnsureGhostRect(const POINT &pt);
        void UpdatePendingDropTarget(const POINT &pt);
        void UpdateDropHover(const POINT &pt);

        // drag helpers
        HitTestResult HitTest(const POINT &pt) const;
        void StartTabDrag(int groupIndex, int tabIndex, const POINT &pt);
        void StartGroupDrag(int groupIndex, const POINT &pt);
        void UpdateDrag(const POINT &pt);
        void CommitDrag(const POINT &pt);
        void CancelDrag();
        void DetachDraggedTab();
        void ReorderTab(int targetGroup, int targetIndex);
        void ReorderGroup(int targetIndex);

        // drop helpers
        std::vector<std::wstring> ExtractFilePathsFromDataObject(IDataObject *dataObject) const;
        HRESULT PerformFileOperation(const std::vector<std::wstring> &sourcePaths, const std::wstring &targetFolder, bool move) const;

private:
        CComPtr<IShellBrowser> m_pShellBrowser = nullptr;
        CComPtr<IWebBrowser2> m_pWebBrowser = nullptr;
        CComPtr<ExplorerTabDropTarget> m_dropTarget;

        std::vector<TabGroup> m_groups;
        bool m_layoutDirty = true;
        bool m_autoSizeTabs = true;
        SIZE m_fixedTabSize = {180, 32};
        int m_tabPaddingX = 14;
        int m_tabPaddingY = 6;
        int m_tabSpacing = 6;
        int m_groupSpacing = 14;
        int m_groupHandleWidth = 8;
        int m_tabMargin = 6;
        int m_rowSpacing = 6;
        int m_minTabWidth = 120;
        int m_maxTabWidth = 280;
        int m_maxRows = 10;
        int m_activeGroup = 0;
        int m_activeTab = 0;
        int m_totalHeight = 0;

        bool m_draggingTab = false;
        bool m_draggingGroup = false;
        bool m_dragClickCandidate = false;
        bool m_detachPending = false;
        POINT m_dragStart = {0};
        POINT m_dragPoint = {0};
        RECT m_dragGhostRect = {0};
        bool m_showGhost = false;
        int m_draggedGroupIndex = -1;
        int m_draggedTabIndex = -1;
        int m_pendingDropGroup = -1;
        int m_pendingDropTab = -1;

        int m_dropHoverGroup = -1;
        int m_dropHoverTab = -1;

        COLORREF m_backgroundColor = RGB(245, 246, 247);
        COLORREF m_borderColor = RGB(160, 160, 160);
};

