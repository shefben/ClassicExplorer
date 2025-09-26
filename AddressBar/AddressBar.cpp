#include "stdafx.h"
#include "framework.h"
#include "resource.h"
#include "ClassicExplorer_i.h"
#include "dllmain.h"

#include "AddressBar.h"

#include "util/shell_helpers.h"

#include <shobjidl.h>

#include <algorithm>
#include <array>
#include <numeric>

namespace
{
        const std::array<COLORREF, 8> kDefaultGroupPalette = {
                RGB(180, 200, 235),
                RGB(190, 220, 180),
                RGB(230, 205, 175),
                RGB(210, 185, 230),
                RGB(200, 200, 200),
                RGB(180, 215, 215),
                RGB(235, 190, 190),
                RGB(200, 210, 165)
        };

        int GetSystemDragThresholdX()
        {
                return GetSystemMetrics(SM_CXDRAG);
        }

        int GetSystemDragThresholdY()
        {
                return GetSystemMetrics(SM_CYDRAG);
        }
}

// ============================================================================
// ExplorerTabDropTarget
// ============================================================================

ExplorerTabDropTarget::ExplorerTabDropTarget(CAddressBar *owner) : m_owner(owner)
{
}

IFACEMETHODIMP ExplorerTabDropTarget::QueryInterface(REFIID riid, void **ppvObject)
{
        if (!ppvObject)
                return E_POINTER;

        if (riid == IID_IUnknown || riid == IID_IDropTarget)
        {
                *ppvObject = static_cast<IDropTarget *>(this);
                AddRef();
                return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
}

ULONG ExplorerTabDropTarget::AddRef()
{
        return InterlockedIncrement(&m_refCount);
}

ULONG ExplorerTabDropTarget::Release()
{
        LONG ref = InterlockedDecrement(&m_refCount);
        if (ref == 0)
        {
                delete this;
        }
        return static_cast<ULONG>(ref);
}

IFACEMETHODIMP ExplorerTabDropTarget::DragEnter(IDataObject *pDataObject, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
        if (!pdwEffect)
                return E_INVALIDARG;

        m_currentDataObject = pDataObject;
        if (!m_owner || !m_owner->CanAcceptDataObject(pDataObject))
        {
                *pdwEffect = DROPEFFECT_NONE;
                return S_OK;
        }

        *pdwEffect = (grfKeyState & MK_SHIFT) ? DROPEFFECT_MOVE : DROPEFFECT_COPY;
        if (m_owner)
        {
                m_owner->HandleExternalDragEnter(grfKeyState, pt, pDataObject);
        }
        return S_OK;
}

IFACEMETHODIMP ExplorerTabDropTarget::DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
        if (!pdwEffect)
                return E_INVALIDARG;

        if (!m_owner || !m_owner->CanAcceptDataObject(m_currentDataObject))
        {
                *pdwEffect = DROPEFFECT_NONE;
                return S_OK;
        }

        *pdwEffect = (grfKeyState & MK_SHIFT) ? DROPEFFECT_MOVE : DROPEFFECT_COPY;
        m_owner->HandleExternalDragOver(grfKeyState, pt);
        return S_OK;
}

IFACEMETHODIMP ExplorerTabDropTarget::DragLeave()
{
        if (m_owner)
        {
                m_owner->HandleExternalDragLeave();
        }
        m_currentDataObject.Release();
        return S_OK;
}

IFACEMETHODIMP ExplorerTabDropTarget::Drop(IDataObject *pDataObject, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
        if (!pdwEffect)
                return E_INVALIDARG;

        if (!m_owner || !m_owner->CanAcceptDataObject(pDataObject))
        {
                *pdwEffect = DROPEFFECT_NONE;
                return S_OK;
        }

        *pdwEffect = (grfKeyState & MK_SHIFT) ? DROPEFFECT_MOVE : DROPEFFECT_COPY;
        m_owner->HandleExternalDrop(pDataObject, grfKeyState, pt);
        m_currentDataObject.Release();
        return S_OK;
}

// ============================================================================
// CAddressBar implementation
// ============================================================================

void CAddressBar::SetBrowsers(CComPtr<IShellBrowser> pShellBrowser, CComPtr<IWebBrowser2> pWebBrowser)
{
        m_pShellBrowser = pShellBrowser;
        m_pWebBrowser = pWebBrowser;
}

HRESULT CAddressBar::InitializeTabs()
{
        LoadSettings();
        EnsureDefaultGroup();
        UpdateActiveTabFromExplorer();
        LayoutTabs();
        InvalidateRect(nullptr, TRUE);
        return S_OK;
}

void CAddressBar::OnExplorerNavigate()
{
        UpdateActiveTabFromExplorer();
}

SIZE CAddressBar::GetDesiredSize() const
{
        SIZE size = { 600, 0 };
        int desiredHeight = m_totalHeight;
        if (desiredHeight <= 0)
        {
                desiredHeight = m_fixedTabSize.cy + (m_tabMargin * 2);
        }
        desiredHeight = std::max(desiredHeight, m_fixedTabSize.cy + (m_tabMargin * 2));
        size.cy = desiredHeight;
        return size;
}

void CAddressBar::HandleExternalDragEnter(DWORD keyState, POINTL pt, IDataObject *dataObject)
{
        UNREFERENCED_PARAMETER(keyState);
        UNREFERENCED_PARAMETER(dataObject);
        POINT client = { static_cast<LONG>(pt.x), static_cast<LONG>(pt.y) };
        ScreenToClient(&client);
        UpdateDropHover(client);
}

void CAddressBar::HandleExternalDragOver(DWORD keyState, POINTL pt)
{
        UNREFERENCED_PARAMETER(keyState);
        POINT client = { static_cast<LONG>(pt.x), static_cast<LONG>(pt.y) };
        ScreenToClient(&client);
        UpdateDropHover(client);
}

void CAddressBar::HandleExternalDragLeave()
{
        m_dropHoverGroup = -1;
        m_dropHoverTab = -1;
        InvalidateRect(nullptr, FALSE);
}

void CAddressBar::HandleExternalDrop(IDataObject *dataObject, DWORD keyState, POINTL pt)
{
        POINT client = { static_cast<LONG>(pt.x), static_cast<LONG>(pt.y) };
        ScreenToClient(&client);

        std::vector<std::wstring> paths = ExtractFilePathsFromDataObject(dataObject);
        if (paths.empty())
        {
                HandleExternalDragLeave();
                return;
        }

        HitTestResult result = HitTest(client);
        if (!result.valid)
        {
                result.groupIndex = m_activeGroup;
                result.tabIndex = m_activeTab;
        }

        if (result.groupIndex < 0 || result.groupIndex >= static_cast<int>(m_groups.size()))
        {
                HandleExternalDragLeave();
                return;
        }

        if (result.tabIndex < 0 && !m_groups[result.groupIndex].tabs.empty())
        {
                result.tabIndex = std::min(m_activeTab, static_cast<int>(m_groups[result.groupIndex].tabs.size() - 1));
        }

        if (result.tabIndex < 0 || result.tabIndex >= static_cast<int>(m_groups[result.groupIndex].tabs.size()))
        {
                HandleExternalDragLeave();
                return;
        }

        const Tab &targetTab = m_groups[result.groupIndex].tabs[result.tabIndex];
        std::wstring targetPath = GetTabFilesystemPath(targetTab);
        if (!targetPath.empty())
        {
                bool move = (keyState & MK_SHIFT) != 0 || (GetKeyState(VK_SHIFT) < 0);
                PerformFileOperation(paths, targetPath, move);
        }

        HandleExternalDragLeave();
}

bool CAddressBar::CanAcceptDataObject(IDataObject *dataObject) const
{
        if (!dataObject)
                return false;

        FORMATETC format = { CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        if (SUCCEEDED(dataObject->QueryGetData(&format)))
                return true;

        FORMATETC shellFormat = { RegisterClipboardFormat(CFSTR_SHELLIDLIST), nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        return SUCCEEDED(dataObject->QueryGetData(&shellFormat));
}

// ============================================================================
// Message handlers
// ============================================================================

LRESULT CAddressBar::OnCreate(UINT, WPARAM, LPARAM, BOOL &)
{
        LoadSettings();
        EnsureDefaultGroup();

        m_dropTarget.Attach(new ExplorerTabDropTarget(this));
        RegisterDragDrop(m_hWnd, m_dropTarget);

        UpdateActiveTabFromExplorer();
        LayoutTabs();
        return 0;
}

LRESULT CAddressBar::OnDestroy(UINT, WPARAM, LPARAM, BOOL &)
{
        RevokeDragDrop(m_hWnd);
        m_dropTarget.Release();
        CancelDrag();
        m_groups.clear();
        return 0;
}

LRESULT CAddressBar::OnPaint(UINT, WPARAM, LPARAM, BOOL &)
{
        LayoutTabsIfNeeded();

        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(&ps);
        RECT clientRect;
        GetClientRect(&clientRect);

        DrawBackground(hdc, clientRect);

        for (const TabGroup &group : m_groups)
        {
                if (!group.tabs.empty())
                {
                        DrawGroup(hdc, group);
                }
        }

        if (m_dropHoverGroup >= 0)
        {
                DrawDropHover(hdc);
        }

        if (m_showGhost)
        {
                DrawGhost(hdc);
        }

        EndPaint(&ps);
        return 0;
}

LRESULT CAddressBar::OnEraseBackground(UINT, WPARAM, LPARAM, BOOL &)
{
        return 1;
}

LRESULT CAddressBar::OnSize(UINT, WPARAM, LPARAM, BOOL &)
{
        m_layoutDirty = true;
        LayoutTabs();
        InvalidateRect(nullptr, TRUE);
        return 0;
}

LRESULT CAddressBar::OnLButtonDown(UINT, WPARAM, LPARAM lParam, BOOL &)
{
        POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        HitTestResult result = HitTest(pt);

        if (result.groupHandle)
        {
                StartGroupDrag(result.groupIndex, pt);
        }
        else if (result.valid)
        {
                StartTabDrag(result.groupIndex, result.tabIndex, pt);
        }

        return 0;
}

LRESULT CAddressBar::OnLButtonUp(UINT, WPARAM, LPARAM lParam, BOOL &)
{
        POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        CommitDrag(pt);
        return 0;
}

LRESULT CAddressBar::OnMouseMove(UINT, WPARAM, LPARAM lParam, BOOL &)
{
        POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        UpdateDrag(pt);
        return 0;
}

LRESULT CAddressBar::OnContextMenu(UINT, WPARAM, LPARAM lParam, BOOL &)
{
        POINT screenPt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        POINT clientPt = screenPt;
        ScreenToClient(&clientPt);

        HitTestResult result = HitTest(clientPt);
        if (result.groupHandle && result.groupIndex >= 0)
        {
                ShowGroupColorMenu(result.groupIndex, screenPt);
        }
        else if (result.valid)
        {
                ShowContextMenuForTab(result.groupIndex, result.tabIndex, screenPt);
        }
        return 0;
}

LRESULT CAddressBar::OnCaptureChanged(UINT, WPARAM, LPARAM, BOOL &)
{
        CancelDrag();
        return 0;
}

// ============================================================================
// Layout helpers
// ============================================================================

void CAddressBar::LoadSettings()
{
        CEUtil::CESettings settings = CEUtil::GetCESettings();
        m_autoSizeTabs = settings.tabAutoSize != 0;
        if (settings.tabFixedWidth > 0)
                m_fixedTabSize.cx = static_cast<int>(settings.tabFixedWidth);
        if (settings.tabFixedHeight > 0)
                m_fixedTabSize.cy = static_cast<int>(settings.tabFixedHeight);
}

void CAddressBar::EnsureDefaultGroup()
{
        if (!m_groups.empty())
                return;

        TabGroup group;
        group.name = L"Group 1";
        group.color = kDefaultGroupPalette.front();
        m_groups.push_back(group);
        m_activeGroup = 0;
        m_activeTab = 0;
}

void CAddressBar::LayoutTabsIfNeeded()
{
        if (m_layoutDirty)
        {
                LayoutTabs();
        }
}

void CAddressBar::LayoutTabs()
{
        RECT clientRect;
        GetClientRect(&clientRect);
        int rowHeight = m_fixedTabSize.cy;

        int x = clientRect.left + m_tabMargin;
        int y = clientRect.top + m_tabMargin;
        int currentRow = 0;

        HDC hdc = GetDC();
        for (auto &group : m_groups)
        {
                if (group.tabs.empty())
                        continue;

                std::vector<int> tabWidths;
                tabWidths.reserve(group.tabs.size());
                for (const Tab &tab : group.tabs)
                {
                        int width = m_autoSizeTabs ? CalculateTabWidth(hdc, tab.title) : m_fixedTabSize.cx;
                        width = std::min(std::max(width, m_minTabWidth), m_maxTabWidth);
                        tabWidths.push_back(width);
                }

                int groupWidth = m_groupHandleWidth;
                if (!tabWidths.empty())
                {
                        groupWidth += std::accumulate(tabWidths.begin(), tabWidths.end(), 0);
                        groupWidth += static_cast<int>((tabWidths.size() - 1)) * m_tabSpacing;
                }

                if (x + groupWidth > clientRect.right - m_tabMargin && currentRow < (m_maxRows - 1))
                {
                        currentRow++;
                        x = clientRect.left + m_tabMargin;
                        y += rowHeight + m_rowSpacing;
                }

                group.bounds.left = x;
                group.bounds.top = y;
                group.bounds.right = x + groupWidth;
                group.bounds.bottom = y + rowHeight;

                int tabX = x + m_groupHandleWidth;
                for (size_t idx = 0; idx < group.tabs.size(); ++idx)
                {
                        RECT tabRect = { tabX, y, tabX + tabWidths[idx], y + rowHeight };
                        group.tabs[idx].bounds = tabRect;
                        tabX += tabWidths[idx] + m_tabSpacing;
                }

                x = group.bounds.right + m_groupSpacing;
        }
        ReleaseDC(hdc);

        m_totalHeight = (currentRow + 1) * rowHeight + (m_tabMargin * 2) + (currentRow * m_rowSpacing);
        m_layoutDirty = false;
        RefreshActiveState();
}

int CAddressBar::CalculateTabWidth(HDC hdc, const std::wstring &text) const
{
        RECT rc = { 0,0,0,0 };
        DrawTextW(hdc, text.c_str(), static_cast<int>(text.length()), &rc, DT_CALCRECT | DT_SINGLELINE);
        int width = rc.right - rc.left + (m_tabPaddingX * 2);
        return width;
}

void CAddressBar::RefreshActiveState()
{
        for (size_t groupIndex = 0; groupIndex < m_groups.size(); ++groupIndex)
        {
                auto &group = m_groups[groupIndex];
                for (size_t tabIndex = 0; tabIndex < group.tabs.size(); ++tabIndex)
                {
                        group.tabs[tabIndex].active = (static_cast<int>(groupIndex) == m_activeGroup && static_cast<int>(tabIndex) == m_activeTab);
                }
        }
}

// ============================================================================
// Painting helpers
// ============================================================================

void CAddressBar::DrawBackground(HDC hdc, const RECT &clientRect) const
{
        HBRUSH background = CreateSolidBrush(m_backgroundColor);
        FillRect(hdc, &clientRect, background);
        DeleteObject(background);
}

void CAddressBar::DrawGroup(HDC hdc, const TabGroup &group) const
{
        DrawGroupHandle(hdc, group);
        for (const Tab &tab : group.tabs)
        {
                        DrawTab(hdc, tab, group.color);
        }
}

void CAddressBar::DrawTab(HDC hdc, const Tab &tab, COLORREF groupColor) const
{
        COLORREF baseColor = tab.active ? AdjustColor(groupColor, 1.2) : groupColor;
        HBRUSH brush = CreateSolidBrush(baseColor);
        FillRect(hdc, &tab.bounds, brush);
        DeleteObject(brush);

        HPEN pen = CreatePen(PS_SOLID, 1, m_borderColor);
        HPEN oldPen = (HPEN)SelectObject(hdc, pen);
        HBRUSH oldBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
        Rectangle(hdc, tab.bounds.left, tab.bounds.top, tab.bounds.right, tab.bounds.bottom);
        SelectObject(hdc, oldBrush);
        SelectObject(hdc, oldPen);
        DeleteObject(pen);

        RECT textRect = tab.bounds;
        InflateRect(&textRect, -m_tabPaddingX, -m_tabPaddingY);
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, RGB(40, 40, 40));
        DrawTextW(hdc, tab.title.c_str(), static_cast<int>(tab.title.length()), &textRect, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_END_ELLIPSIS);
}

void CAddressBar::DrawGroupHandle(HDC hdc, const TabGroup &group) const
{
        RECT handleRect = group.bounds;
        handleRect.right = handleRect.left + m_groupHandleWidth;
        HBRUSH brush = CreateSolidBrush(AdjustColor(group.color, 0.8));
        FillRect(hdc, &handleRect, brush);
        DeleteObject(brush);
}

void CAddressBar::DrawGhost(HDC hdc) const
{
        if (!m_showGhost)
                return;

        RECT rect = m_dragGhostRect;
        int width = rect.right - rect.left;
        int height = rect.bottom - rect.top;
        if (width <= 0 || height <= 0)
                return;

        BITMAPINFO bmi = {};
        bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
        bmi.bmiHeader.biWidth = width;
        bmi.bmiHeader.biHeight = -height;
        bmi.bmiHeader.biPlanes = 1;
        bmi.bmiHeader.biBitCount = 32;
        bmi.bmiHeader.biCompression = BI_RGB;

        void *bits = nullptr;
        HBITMAP bitmap = CreateDIBSection(hdc, &bmi, DIB_RGB_COLORS, &bits, nullptr, 0);
        if (!bitmap)
                return;

        UINT32 *pixelData = static_cast<UINT32 *>(bits);
        COLORREF ghostColor = RGB(120, 120, 120);
        UINT32 argb = (120 << 24) | (GetBValue(ghostColor) << 16) | (GetGValue(ghostColor) << 8) | GetRValue(ghostColor);
        std::fill(pixelData, pixelData + (width * height), argb);

        HDC memDc = CreateCompatibleDC(hdc);
        HGDIOBJ old = SelectObject(memDc, bitmap);
        BLENDFUNCTION blend = { AC_SRC_OVER, 0, 150, 0 };
        AlphaBlend(hdc, rect.left, rect.top, width, height, memDc, 0, 0, width, height, blend);
        SelectObject(memDc, old);
        DeleteDC(memDc);
        DeleteObject(bitmap);
}

void CAddressBar::DrawDropHover(HDC hdc) const
{
        if (m_dropHoverGroup < 0 || m_dropHoverGroup >= static_cast<int>(m_groups.size()))
                return;

        RECT highlightRect = {0};
        if (m_dropHoverTab >= 0 && m_dropHoverTab < static_cast<int>(m_groups[m_dropHoverGroup].tabs.size()))
        {
                highlightRect = m_groups[m_dropHoverGroup].tabs[m_dropHoverTab].bounds;
        }
        else
        {
                highlightRect = m_groups[m_dropHoverGroup].bounds;
        }

        HPEN pen = CreatePen(PS_DOT, 2, RGB(30, 120, 215));
        HPEN oldPen = (HPEN)SelectObject(hdc, pen);
        HBRUSH oldBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
        Rectangle(hdc, highlightRect.left, highlightRect.top, highlightRect.right, highlightRect.bottom);
        SelectObject(hdc, oldBrush);
        SelectObject(hdc, oldPen);
        DeleteObject(pen);
}

COLORREF CAddressBar::AdjustColor(COLORREF color, double factor)
{
        int r = std::clamp(static_cast<int>(GetRValue(color) * factor), 0, 255);
        int g = std::clamp(static_cast<int>(GetGValue(color) * factor), 0, 255);
        int b = std::clamp(static_cast<int>(GetBValue(color) * factor), 0, 255);
        return RGB(r, g, b);
}

// ============================================================================
// Tab management
// ============================================================================

HRESULT CAddressBar::AddTabForLocation(PIDLIST_ABSOLUTE pidl, bool makeActive, bool navigate, COLORREF colorOverride)
{
        if (!pidl)
                return E_INVALIDARG;

        EnsureDefaultGroup();

        Tab newTab;
        newTab.pidl.reset(pidl);
        newTab.active = false;

        CComHeapPtr<wchar_t> name;
        if (SUCCEEDED(SHGetNameFromIDList(pidl, SIGDN_NORMALDISPLAY, &name)))
        {
                newTab.title = name.m_pData;
        }
        else
        {
                newTab.title = L"Tab";
        }

        TabGroup &targetGroup = m_groups[m_activeGroup];
        if (targetGroup.tabs.empty())
        {
                targetGroup.color = colorOverride;
        }
        targetGroup.tabs.push_back(std::move(newTab));
        m_activeTab = static_cast<int>(targetGroup.tabs.size() - 1);

        if (makeActive)
        {
                ActivateTab(m_activeGroup, m_activeTab, navigate);
        }

        m_layoutDirty = true;
        LayoutTabs();
        InvalidateRect(nullptr, TRUE);
        return S_OK;
}

void CAddressBar::RemoveEmptyGroups()
{
        if (m_groups.empty())
                return;

        bool removed = false;
        for (auto it = m_groups.begin(); it != m_groups.end();)
        {
                if (it->tabs.empty() && m_groups.size() > 1)
                {
                        int index = static_cast<int>(std::distance(m_groups.begin(), it));
                        it = m_groups.erase(it);
                        if (m_activeGroup >= index && m_activeGroup > 0)
                                --m_activeGroup;
                        removed = true;
                }
                else
                {
                        ++it;
                }
        }

        if (m_groups.empty())
        {
                EnsureDefaultGroup();
        }

        if (removed)
        {
                m_activeGroup = std::clamp(m_activeGroup, 0, static_cast<int>(m_groups.size() - 1));
                if (!m_groups[m_activeGroup].tabs.empty())
                        m_activeTab = std::clamp(m_activeTab, 0, static_cast<int>(m_groups[m_activeGroup].tabs.size() - 1));
        }
}

void CAddressBar::ActivateTab(int groupIndex, int tabIndex, bool navigate)
{
        if (groupIndex < 0 || groupIndex >= static_cast<int>(m_groups.size()))
                return;
        if (tabIndex < 0 || tabIndex >= static_cast<int>(m_groups[groupIndex].tabs.size()))
                return;

        m_activeGroup = groupIndex;
        m_activeTab = tabIndex;
        RefreshActiveState();

        if (navigate && m_pShellBrowser)
        {
                Tab &tab = m_groups[groupIndex].tabs[tabIndex];
                if (tab.pidl.pidl)
                {
                        m_pShellBrowser->BrowseObject(tab.pidl.pidl, SBSP_SAMEBROWSER | SBSP_ABSOLUTE);
                }
        }

        m_layoutDirty = true;
        LayoutTabs();
        InvalidateRect(nullptr, FALSE);
}

void CAddressBar::ActivateTabByPidl(PIDLIST_ABSOLUTE pidl)
{
        if (!pidl)
                return;

        for (size_t groupIndex = 0; groupIndex < m_groups.size(); ++groupIndex)
        {
                auto &group = m_groups[groupIndex];
                for (size_t tabIndex = 0; tabIndex < group.tabs.size(); ++tabIndex)
                {
                        if (group.tabs[tabIndex].pidl.pidl && ILIsEqual(group.tabs[tabIndex].pidl.pidl, pidl))
                        {
                                ActivateTab(static_cast<int>(groupIndex), static_cast<int>(tabIndex), false);
                                return;
                        }
                }
        }
}

void CAddressBar::UpdateActiveTabFromExplorer()
{
        if (!m_pShellBrowser)
                return;

        PIDLIST_ABSOLUTE pidl = nullptr;
        if (FAILED(CEUtil::GetCurrentFolderPidl(m_pShellBrowser, &pidl)))
                return;

        ActivateTabByPidl(pidl);

        bool alreadyPresent = false;
        for (const auto &group : m_groups)
        {
                for (const auto &tab : group.tabs)
                {
                        if (tab.pidl.pidl && ILIsEqual(tab.pidl.pidl, pidl))
                        {
                                alreadyPresent = true;
                                break;
                        }
                }
                if (alreadyPresent)
                        break;
        }

        if (!alreadyPresent)
        {
                AddTabForLocation(pidl, true, false, m_groups[m_activeGroup].color);
        }

        CoTaskMemFree(pidl);
        m_layoutDirty = true;
        LayoutTabs();
        InvalidateRect(nullptr, FALSE);
}

void CAddressBar::CreateNewWindowForTab(const Tab &tab)
{
        std::wstring path = GetTabFilesystemPath(tab);
        if (path.empty())
                return;

        ShellExecuteW(nullptr, L"open", L"explorer.exe", path.c_str(), nullptr, SW_SHOWNORMAL);
}

std::wstring CAddressBar::GetTabFilesystemPath(const Tab &tab) const
{
        if (!tab.pidl.pidl)
                return L"";

        CComHeapPtr<wchar_t> buffer;
        if (SUCCEEDED(SHGetNameFromIDList(tab.pidl.pidl, SIGDN_FILESYSPATH, &buffer)))
        {
                        return std::wstring(buffer.m_pData);
        }

        if (SUCCEEDED(SHGetNameFromIDList(tab.pidl.pidl, SIGDN_DESKTOPABSOLUTEPARSING, &buffer)))
        {
                return std::wstring(buffer.m_pData);
        }

        return L"";
}

void CAddressBar::SetGroupColor(int groupIndex, COLORREF color)
{
        if (groupIndex < 0 || groupIndex >= static_cast<int>(m_groups.size()))
                return;
        m_groups[groupIndex].color = color;
        InvalidateRect(nullptr, FALSE);
}

void CAddressBar::ShowGroupColorMenu(int groupIndex, POINT screenPoint)
{
        if (groupIndex < 0 || groupIndex >= static_cast<int>(m_groups.size()))
                return;

        HMENU menu = CreatePopupMenu();
        for (size_t idx = 0; idx < kDefaultGroupPalette.size(); ++idx)
        {
                UINT flags = MF_STRING;
                if (m_groups[groupIndex].color == kDefaultGroupPalette[idx])
                        flags |= MF_CHECKED;
                wchar_t label[32];
                swprintf_s(label, L"Color %zu", idx + 1);
                AppendMenuW(menu, flags, 7200 + static_cast<UINT>(idx), label);
        }

        AppendMenuW(menu, MF_STRING, 7300, L"Reset color");

        UINT command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_LEFTALIGN | TPM_TOPALIGN, screenPoint.x, screenPoint.y, 0, m_hWnd, nullptr);
        DestroyMenu(menu);

        if (command >= 7200 && command < 7200 + kDefaultGroupPalette.size())
        {
                size_t paletteIndex = command - 7200;
                SetGroupColor(groupIndex, kDefaultGroupPalette[paletteIndex]);
        }
        else if (command == 7300)
        {
                SetGroupColor(groupIndex, kDefaultGroupPalette.front());
        }
}

void CAddressBar::ShowContextMenuForTab(int groupIndex, int tabIndex, POINT screenPoint)
{
        if (groupIndex < 0 || groupIndex >= static_cast<int>(m_groups.size()))
                return;
        if (tabIndex < 0 || tabIndex >= static_cast<int>(m_groups[groupIndex].tabs.size()))
                return;

        HMENU menu = CreatePopupMenu();
        AppendMenuW(menu, MF_STRING, 7400, L"Close tab");
        AppendMenuW(menu, MF_STRING, 7401, L"Move to new group");
        AppendMenuW(menu, MF_STRING, 7402, L"Open in new window");

        UINT command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_LEFTALIGN | TPM_TOPALIGN, screenPoint.x, screenPoint.y, 0, m_hWnd, nullptr);
        DestroyMenu(menu);

        switch (command)
        {
        case 7400:
        {
                auto &group = m_groups[groupIndex];
                if (tabIndex >= 0 && tabIndex < static_cast<int>(group.tabs.size()))
                {
                        group.tabs.erase(group.tabs.begin() + tabIndex);
                        if (m_activeGroup == groupIndex)
                        {
                                m_activeTab = std::clamp(m_activeTab, 0, static_cast<int>(group.tabs.size()) - 1);
                        }
                        RemoveEmptyGroups();
                        m_layoutDirty = true;
                        LayoutTabs();
                        InvalidateRect(nullptr, TRUE);
                }
                break;
        }
        case 7401:
        {
                Tab tab = std::move(m_groups[groupIndex].tabs[tabIndex]);
                m_groups[groupIndex].tabs.erase(m_groups[groupIndex].tabs.begin() + tabIndex);
                RemoveEmptyGroups();

                TabGroup newGroup;
                wchar_t label[32];
                swprintf_s(label, L"Group %zu", m_groups.size() + 1ULL);
                newGroup.name = label;
                size_t paletteIndex = (m_groups.size()) % kDefaultGroupPalette.size();
                newGroup.color = kDefaultGroupPalette[paletteIndex];
                newGroup.tabs.push_back(std::move(tab));
                m_groups.insert(m_groups.begin() + std::min(static_cast<size_t>(groupIndex + 1), m_groups.size()), std::move(newGroup));
                m_activeGroup = static_cast<int>(std::min(static_cast<size_t>(groupIndex + 1), m_groups.size() - 1));
                m_activeTab = 0;
                RefreshActiveState();
                m_layoutDirty = true;
                LayoutTabs();
                InvalidateRect(nullptr, TRUE);
                break;
        }
        case 7402:
                CreateNewWindowForTab(m_groups[groupIndex].tabs[tabIndex]);
                break;
        default:
                break;
        }
}

void CAddressBar::EnsureGhostRect(const POINT &pt)
{
        if (m_draggingTab && m_draggedGroupIndex >= 0 && m_draggedTabIndex >= 0)
        {
                const RECT &original = m_groups[m_draggedGroupIndex].tabs[m_draggedTabIndex].bounds;
                int dx = pt.x - m_dragStart.x;
                int dy = pt.y - m_dragStart.y;
                m_dragGhostRect = { original.left + dx, original.top + dy, original.right + dx, original.bottom + dy };
        }
        else if (m_draggingGroup && m_draggedGroupIndex >= 0)
        {
                const RECT &original = m_groups[m_draggedGroupIndex].bounds;
                int dx = pt.x - m_dragStart.x;
                int dy = pt.y - m_dragStart.y;
                m_dragGhostRect = { original.left + dx, original.top + dy, original.right + dx, original.bottom + dy };
        }
}

void CAddressBar::UpdatePendingDropTarget(const POINT &pt)
{
        m_pendingDropGroup = -1;
        m_pendingDropTab = -1;

        for (size_t groupIndex = 0; groupIndex < m_groups.size(); ++groupIndex)
        {
                const TabGroup &group = m_groups[groupIndex];
                if (!PtInRect(&group.bounds, pt))
                        continue;

                if (m_draggingGroup)
                {
                        int midpoint = group.bounds.left + ((group.bounds.right - group.bounds.left) / 2);
                        if (pt.x < midpoint)
                                m_pendingDropGroup = static_cast<int>(groupIndex);
                        else
                                m_pendingDropGroup = static_cast<int>(groupIndex + 1);
                        return;
                }

                for (size_t tabIndex = 0; tabIndex < group.tabs.size(); ++tabIndex)
                {
                        const RECT &bounds = group.tabs[tabIndex].bounds;
                        if (pt.x <= bounds.left + ((bounds.right - bounds.left) / 2))
                        {
                                m_pendingDropGroup = static_cast<int>(groupIndex);
                                m_pendingDropTab = static_cast<int>(tabIndex);
                                return;
                        }
                }

                m_pendingDropGroup = static_cast<int>(groupIndex);
                m_pendingDropTab = static_cast<int>(group.tabs.size());
                return;
        }
}

void CAddressBar::UpdateDropHover(const POINT &pt)
{
        m_dropHoverGroup = -1;
        m_dropHoverTab = -1;

        for (size_t groupIndex = 0; groupIndex < m_groups.size(); ++groupIndex)
        {
                const TabGroup &group = m_groups[groupIndex];
                if (!PtInRect(&group.bounds, pt))
                        continue;

                m_dropHoverGroup = static_cast<int>(groupIndex);
                for (size_t tabIndex = 0; tabIndex < group.tabs.size(); ++tabIndex)
                {
                        if (PtInRect(&group.tabs[tabIndex].bounds, pt))
                        {
                                m_dropHoverTab = static_cast<int>(tabIndex);
                                break;
                        }
                }
                break;
        }
        InvalidateRect(nullptr, FALSE);
}

// ============================================================================
// Drag helpers
// ============================================================================

CAddressBar::HitTestResult CAddressBar::HitTest(const POINT &pt) const
{
        HitTestResult result;
        for (size_t groupIndex = 0; groupIndex < m_groups.size(); ++groupIndex)
        {
                const TabGroup &group = m_groups[groupIndex];
                if (group.tabs.empty())
                        continue;

                RECT handleRect = group.bounds;
                handleRect.right = handleRect.left + m_groupHandleWidth;
                if (PtInRect(&handleRect, pt))
                {
                        result.valid = true;
                        result.groupHandle = true;
                        result.groupIndex = static_cast<int>(groupIndex);
                        return result;
                }

                for (size_t tabIndex = 0; tabIndex < group.tabs.size(); ++tabIndex)
                {
                        if (PtInRect(&group.tabs[tabIndex].bounds, pt))
                        {
                                result.valid = true;
                                result.groupIndex = static_cast<int>(groupIndex);
                                result.tabIndex = static_cast<int>(tabIndex);
                                return result;
                        }
                }
        }
        return result;
}

void CAddressBar::StartTabDrag(int groupIndex, int tabIndex, const POINT &pt)
{
        if (groupIndex < 0 || groupIndex >= static_cast<int>(m_groups.size()))
                return;
        if (tabIndex < 0 || tabIndex >= static_cast<int>(m_groups[groupIndex].tabs.size()))
                return;

        m_draggingTab = true;
        m_draggingGroup = false;
        m_dragClickCandidate = true;
        m_detachPending = false;
        m_draggedGroupIndex = groupIndex;
        m_draggedTabIndex = tabIndex;
        m_dragStart = pt;
        m_dragPoint = pt;
        m_showGhost = false;
        SetCapture();
}

void CAddressBar::StartGroupDrag(int groupIndex, const POINT &pt)
{
        if (groupIndex < 0 || groupIndex >= static_cast<int>(m_groups.size()))
                return;

        m_draggingGroup = true;
        m_draggingTab = false;
        m_dragClickCandidate = true;
        m_detachPending = false;
        m_draggedGroupIndex = groupIndex;
        m_draggedTabIndex = -1;
        m_dragStart = pt;
        m_dragPoint = pt;
        m_showGhost = false;
        SetCapture();
}

void CAddressBar::UpdateDrag(const POINT &pt)
{
        if (!m_draggingTab && !m_draggingGroup)
                return;

        m_dragPoint = pt;
        int dx = std::abs(pt.x - m_dragStart.x);
        int dy = std::abs(pt.y - m_dragStart.y);
        if (m_dragClickCandidate && (dx > GetSystemDragThresholdX() / 2 || dy > GetSystemDragThresholdY() / 2))
        {
                m_dragClickCandidate = false;
                m_showGhost = true;
        }

        if (m_showGhost)
        {
                EnsureGhostRect(pt);
                UpdatePendingDropTarget(pt);
                InvalidateRect(nullptr, FALSE);
        }

        RECT clientRect;
        GetClientRect(&clientRect);
        int detachThreshold = 40;
        if (pt.y < clientRect.top - detachThreshold || pt.y > clientRect.bottom + detachThreshold)
        {
                m_detachPending = true;
        }
        else
        {
                m_detachPending = false;
        }
}

void CAddressBar::CommitDrag(const POINT &pt)
{
        if (!m_draggingTab && !m_draggingGroup)
                return;

        ReleaseCapture();

        if (m_dragClickCandidate)
        {
                if (m_draggingTab && m_draggedGroupIndex >= 0 && m_draggedTabIndex >= 0)
                {
                        ActivateTab(m_draggedGroupIndex, m_draggedTabIndex, true);
                }
                CancelDrag();
                return;
        }

        if (m_draggingTab && m_detachPending)
        {
                DetachDraggedTab();
                CancelDrag();
                return;
        }

        if (m_draggingTab && m_pendingDropGroup >= 0)
        {
                int targetIndex = (m_pendingDropTab >= 0) ? m_pendingDropTab : m_groups[m_pendingDropGroup].tabs.size();
                ReorderTab(m_pendingDropGroup, targetIndex);
        }
        else if (m_draggingGroup && m_pendingDropGroup >= 0)
        {
                ReorderGroup(m_pendingDropGroup);
        }

        CancelDrag();
        m_layoutDirty = true;
        LayoutTabs();
        InvalidateRect(nullptr, TRUE);
}

void CAddressBar::CancelDrag()
{
        m_draggingTab = false;
        m_draggingGroup = false;
        m_dragClickCandidate = false;
        m_detachPending = false;
        m_showGhost = false;
        m_pendingDropGroup = -1;
        m_pendingDropTab = -1;
        m_draggedGroupIndex = -1;
        m_draggedTabIndex = -1;
}

void CAddressBar::DetachDraggedTab()
{
        if (!m_draggingTab || m_draggedGroupIndex < 0 || m_draggedGroupIndex >= static_cast<int>(m_groups.size()))
                return;
        if (m_draggedTabIndex < 0 || m_draggedTabIndex >= static_cast<int>(m_groups[m_draggedGroupIndex].tabs.size()))
                return;

        Tab tab = std::move(m_groups[m_draggedGroupIndex].tabs[m_draggedTabIndex]);
        m_groups[m_draggedGroupIndex].tabs.erase(m_groups[m_draggedGroupIndex].tabs.begin() + m_draggedTabIndex);
        RemoveEmptyGroups();
        CreateNewWindowForTab(tab);
}

void CAddressBar::ReorderTab(int targetGroup, int targetIndex)
{
        if (m_draggedGroupIndex < 0 || m_draggedGroupIndex >= static_cast<int>(m_groups.size()))
                return;
        if (m_draggedTabIndex < 0 || m_draggedTabIndex >= static_cast<int>(m_groups[m_draggedGroupIndex].tabs.size()))
                return;

        Tab movingTab = std::move(m_groups[m_draggedGroupIndex].tabs[m_draggedTabIndex]);
        m_groups[m_draggedGroupIndex].tabs.erase(m_groups[m_draggedGroupIndex].tabs.begin() + m_draggedTabIndex);

        if (targetGroup == m_draggedGroupIndex && targetIndex > m_draggedTabIndex)
                --targetIndex;

        if (m_groups[m_draggedGroupIndex].tabs.empty() && m_groups.size() > 1)
        {
                if (targetGroup > m_draggedGroupIndex)
                        --targetGroup;
                m_groups.erase(m_groups.begin() + m_draggedGroupIndex);
                if (m_activeGroup > m_draggedGroupIndex && m_activeGroup > 0)
                        --m_activeGroup;
        }

        targetGroup = std::clamp(targetGroup, 0, static_cast<int>(m_groups.size() - 1));
        targetIndex = std::clamp(targetIndex, 0, static_cast<int>(m_groups[targetGroup].tabs.size()));

        m_groups[targetGroup].tabs.insert(m_groups[targetGroup].tabs.begin() + targetIndex, std::move(movingTab));
        m_activeGroup = targetGroup;
        m_activeTab = targetIndex;
        RefreshActiveState();
}

void CAddressBar::ReorderGroup(int targetIndex)
{
        if (m_draggedGroupIndex < 0 || m_draggedGroupIndex >= static_cast<int>(m_groups.size()))
                return;

        TabGroup movingGroup = std::move(m_groups[m_draggedGroupIndex]);
        m_groups.erase(m_groups.begin() + m_draggedGroupIndex);

        if (targetIndex > m_draggedGroupIndex)
                --targetIndex;
        targetIndex = std::clamp(targetIndex, 0, static_cast<int>(m_groups.size()));

        m_groups.insert(m_groups.begin() + targetIndex, std::move(movingGroup));
        m_activeGroup = targetIndex;
        m_activeTab = std::clamp(m_activeTab, 0, static_cast<int>(m_groups[m_activeGroup].tabs.size()) - 1);
        RefreshActiveState();
}

// ============================================================================
// Drop helpers
// ============================================================================

std::vector<std::wstring> CAddressBar::ExtractFilePathsFromDataObject(IDataObject *dataObject) const
{
        std::vector<std::wstring> paths;
        if (!dataObject)
                return paths;

        FORMATETC fmt = { CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        STGMEDIUM medium = {};
        if (SUCCEEDED(dataObject->GetData(&fmt, &medium)))
        {
                HDROP drop = static_cast<HDROP>(GlobalLock(medium.hGlobal));
                if (drop)
                {
                        UINT count = DragQueryFileW(drop, 0xFFFFFFFF, nullptr, 0);
                        for (UINT i = 0; i < count; ++i)
                        {
                        UINT length = DragQueryFileW(drop, i, nullptr, 0);
                        std::wstring buffer(length + 1, L'\0');
                        DragQueryFileW(drop, i, buffer.data(), length + 1);
                        paths.push_back(buffer.c_str());
                        }
                        GlobalUnlock(medium.hGlobal);
                }
                ReleaseStgMedium(&medium);
                if (!paths.empty())
                        return paths;
        }

        FORMATETC fmtShell = { RegisterClipboardFormat(CFSTR_SHELLIDLIST), nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        if (SUCCEEDED(dataObject->GetData(&fmtShell, &medium)))
        {
                CIDA *cIda = static_cast<CIDA *>(GlobalLock(medium.hGlobal));
                if (cIda)
                {
                        for (UINT i = 0; i < cIda->cidl; ++i)
                        {
                                LPCITEMIDLIST folder = reinterpret_cast<LPCITEMIDLIST>(reinterpret_cast<const BYTE *>(cIda) + cIda->aoffset[0]);
                                LPCITEMIDLIST relative = reinterpret_cast<LPCITEMIDLIST>(reinterpret_cast<const BYTE *>(cIda) + cIda->aoffset[i + 1]);
                                PIDLIST_ABSOLUTE absolute = ILCombine(folder, relative);
                                if (absolute)
                                {
                                        CComHeapPtr<wchar_t> buffer;
                                        if (SUCCEEDED(SHGetNameFromIDList(absolute, SIGDN_FILESYSPATH, &buffer)))
                                        {
                                                paths.push_back(buffer.m_pData);
                                        }
                                        CoTaskMemFree(absolute);
                                }
                        }
                        GlobalUnlock(medium.hGlobal);
                }
                ReleaseStgMedium(&medium);
        }

        return paths;
}

HRESULT CAddressBar::PerformFileOperation(const std::vector<std::wstring> &sourcePaths, const std::wstring &targetFolder, bool move) const
{
        if (sourcePaths.empty() || targetFolder.empty())
                return E_INVALIDARG;

        CComPtr<IFileOperation> fileOperation;
        HRESULT hr = CoCreateInstance(CLSID_FileOperation, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&fileOperation));
        if (FAILED(hr))
                return hr;

        DWORD flags = FOFX_SHOWELEVATIONPROMPT | FOF_NOCONFIRMATION | FOF_NOCONFIRMMKDIR | FOF_SILENT;
        fileOperation->SetOperationFlags(flags);

        CComPtr<IShellItem> destination;
        hr = SHCreateItemFromParsingName(targetFolder.c_str(), nullptr, IID_PPV_ARGS(&destination));
        if (FAILED(hr))
                return hr;

        for (const std::wstring &path : sourcePaths)
        {
                CComPtr<IShellItem> source;
                if (FAILED(SHCreateItemFromParsingName(path.c_str(), nullptr, IID_PPV_ARGS(&source))))
                        continue;

                if (move)
                        fileOperation->MoveItem(source, destination, nullptr, nullptr);
                else
                        fileOperation->CopyItem(source, destination, nullptr, nullptr);
        }

        return fileOperation->PerformOperations();
}

