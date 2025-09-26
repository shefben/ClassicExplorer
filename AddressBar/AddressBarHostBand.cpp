/*
 * AddressBarHostBand.cpp: Implements the host rebar band for the address bar toolbar.
 * 
 * For the address bar implementation proper, see AddressBar.cpp.
 * 
 * As usual, the window itself is created in the SetSite method.
 * 
 * ================================================================================================
 * ---- IMPORTANT FUNCTIONS ----
 * 
 *  - SetSite: Installs the toolbar.
 *  - GetBandInfo: Provides important metadata about the toolbar.
 */

#include "stdafx.h"
#include "framework.h"
#include "resource.h"
#include "ClassicExplorer_i.h"
#include "dllmain.h"
#include <commoncontrols.h>

#include "AddressBarHostBand.h"

//================================================================================================================
// implement IDeskBand:
//

/*
 * GetBandInfo: This is queried by the Shell and must return relevant information about
 *              the Desk Band.
 */
STDMETHODIMP CAddressBarHostBand::GetBandInfo(DWORD dwBandId, DWORD dwViewMode, DESKBANDINFO *pDbi)
{
        UNREFERENCED_PARAMETER(dwBandId);
        UNREFERENCED_PARAMETER(dwViewMode);

        if (pDbi)
        {
                SIZE desiredSize = m_addressBar.GetDesiredSize();
                if (pDbi->dwMask & DBIM_MINSIZE)
                {
                        pDbi->ptMinSize.x = 200;
                        pDbi->ptMinSize.y = desiredSize.cy;
                }
                if (pDbi->dwMask & DBIM_MAXSIZE)
                {
                        pDbi->ptMaxSize.x = 0; // 0 = ignored
                        pDbi->ptMaxSize.y = -1; // -1 = unlimited
                }
                if (pDbi->dwMask & DBIM_INTEGRAL)
                {
                        pDbi->ptIntegral.x = 0;
                        pDbi->ptIntegral.y = 1;
                }
                if (pDbi->dwMask & DBIM_ACTUAL)
                {
                        pDbi->ptActual.x = desiredSize.cx;
                        pDbi->ptActual.y = desiredSize.cy;
                }
                if (pDbi->dwMask & DBIM_TITLE)
                {
                        wcscpy_s(pDbi->wszTitle, L"Explorer Tabs");
                }
                if (pDbi->dwMask & DBIM_BKCOLOR)
                {
                        pDbi->dwMask &= ~DBIM_BKCOLOR;
                }
        }

        return S_OK;
}

//================================================================================================================
// implement IOleWindow:
//
STDMETHODIMP CAddressBarHostBand::GetWindow(HWND *hWnd)
{
	if (!hWnd)
	{
		return E_INVALIDARG;
	}

	*hWnd = m_addressBar.GetToolbar();
	return S_OK;
}

STDMETHODIMP CAddressBarHostBand::ContextSensitiveHelp(BOOL fEnterMode)
{
	return S_OK;
}

//================================================================================================================
// implement IDockingWindow:
//
STDMETHODIMP CAddressBarHostBand::CloseDW(unsigned long dwReserved)
{
	return ShowDW(FALSE);
}

STDMETHODIMP CAddressBarHostBand::ResizeBorderDW(const RECT *pRcBorder, IUnknown *pUnkToolbarSite, BOOL fReserved)
{
	return E_NOTIMPL;
}

STDMETHODIMP CAddressBarHostBand::ShowDW(BOOL fShow)
{
	ShowWindow(m_addressBar.GetToolbar(), fShow ? SW_SHOW : SW_HIDE);
	return S_OK;
}

//================================================================================================================
// implement IObjectWithSite:
//

/*
 * SetSite: Responsible for installation or removal of the toolbar band from a location provided
 *          by the Shell.
 * 
 * This function is additionally responsible for obtaining the shell control APIs and creating
 * the actual toolbar control window.
 */
STDMETHODIMP CAddressBarHostBand::SetSite(IUnknown *pUnkSite)
{
	IObjectWithSiteImpl<CAddressBarHostBand>::SetSite(pUnkSite);

	if (m_addressBar.IsWindow())
	{
		m_addressBar.DestroyWindow();
	}

	// If pUnkSite is not NULL, then the site is being set.
	// Otherwise, the site is being removed.
	if (pUnkSite) // hook:
	{
		if (FAILED(pUnkSite->QueryInterface(IID_IInputObjectSite, (void **)&m_pSite)))
		{
			return E_FAIL;
		}

		HWND hWndParent = NULL;

		CComQIPtr<IOleWindow> pOleWindow = pUnkSite;
		if (pOleWindow)
			pOleWindow->GetWindow(&hWndParent);

		if (!IsWindow(hWndParent))
		{
			return E_FAIL;
		}

		m_parentWindow = GetAncestor(hWndParent, GA_ROOT);

		// Create the toolbar window proper:
		m_addressBar.Create(hWndParent, NULL, NULL, WS_CHILD);
		//m_addressBar.CreateBand(hWndParent);

		if (!m_addressBar.IsWindow())
		{
			return E_FAIL;
		}

		CComQIPtr<IServiceProvider> pProvider = pUnkSite;

		if (pProvider)
		{
			CComPtr<IShellBrowser> pShellBrowser;
			pProvider->QueryService(SID_SShellBrowser, IID_IShellBrowser, (void **)&pShellBrowser);
			pProvider->QueryService(SID_SWebBrowserApp, IID_IWebBrowser2, (void **)&m_pWebBrowser);

			if (m_pWebBrowser)
			{
				if (m_dwEventCookie == 0xFEFEFEFE)
				{
					DispEventAdvise(m_pWebBrowser, &DIID_DWebBrowserEvents2);
				}
			}

                        m_addressBar.SetBrowsers(pShellBrowser, m_pWebBrowser);
                        m_addressBar.InitializeTabs();
                }
        }
        else // unhook:
	{
		m_pSite = NULL;
		m_parentWindow = NULL;
	}

	return S_OK;
}

//================================================================================================================
// handle DWebBrowserEvents2:
//

STDMETHODIMP CAddressBarHostBand::OnNavigateComplete(IDispatch *pDisp, VARIANT *url)
{
        UNREFERENCED_PARAMETER(pDisp);
        UNREFERENCED_PARAMETER(url);
        m_addressBar.OnExplorerNavigate();

        return S_OK;
}

/**
 * OnQuit: Called when the user attempts to quit the Shell browser.
 * 
 * This detaches the event listener we installed in order to listen for navigation
 * events.
 * 
 * Copied from Open-Shell implementation here:
 * https://github.com/Open-Shell/Open-Shell-Menu/blob/master/Src/ClassicExplorer/ExplorerBand.cpp#L2280-L2285
 */
STDMETHODIMP CAddressBarHostBand::OnQuit()
{
	if (m_pWebBrowser && m_dwEventCookie != 0xFEFEFEFE)
	{
		return DispEventUnadvise(m_pWebBrowser, &DIID_DWebBrowserEvents2);
	}

	return S_OK;
}

//================================================================================================================
// implement IInputObject:
//

STDMETHODIMP CAddressBarHostBand::HasFocusIO()
{
        return GetFocus() == m_addressBar.GetToolbar() ? S_OK : S_FALSE;
}

STDMETHODIMP CAddressBarHostBand::TranslateAcceleratorIO(MSG *pMsg)
{
        UNREFERENCED_PARAMETER(pMsg);
        return S_FALSE;
}

STDMETHODIMP CAddressBarHostBand::UIActivateIO(BOOL fActivate, MSG *pMsg)
{
        UNREFERENCED_PARAMETER(pMsg);
        if (fActivate)
        {
                m_pSite->OnFocusChangeIS((IDeskBand *)this, fActivate);
                ::SetFocus(m_addressBar.GetToolbar());
        }

        return S_OK;
}

//================================================================================================================
// implement IInputObjectSite:
//

STDMETHODIMP CAddressBarHostBand::OnFocusChangeIS(IUnknown *pUnkObj, BOOL fSetFocus)
{
	return E_NOTIMPL;
}