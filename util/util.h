#pragma once
#ifndef _UTIL_H
#define _UTIL_H

#include "stdafx.h"
#include "framework.h"
#include "resource.h"
//#include "ClassicExplorer_i.h"
#include "dllmain.h"

enum ClassicExplorerTheme
{
	CLASSIC_EXPLORER_NONE = -1,
	CLASSIC_EXPLORER_2K = 0,
	CLASSIC_EXPLORER_XP = 1,
	CLASSIC_EXPLORER_10 = 2,
	CLASSIC_EXPLORER_MEMPHIS = 3
};

namespace CEUtil
{
        struct CESettings
        {
                ClassicExplorerTheme theme = CLASSIC_EXPLORER_NONE;
                LONG showGoButton = -1;
                LONG showAddressLabel = -1;
                LONG showFullAddress = -1;
                LONG tabAutoSize = -1;
                LONG tabFixedWidth = -1;
                LONG tabFixedHeight = -1;

                CESettings() = default;

                CESettings(
                        ClassicExplorerTheme t,
                        LONG showGo,
                        LONG showLabel,
                        LONG showFull,
                        LONG autoSize = -1,
                        LONG fixedWidth = -1,
                        LONG fixedHeight = -1)
                {
                        theme = t;
                        showGoButton = showGo;
                        showAddressLabel = showLabel;
                        showFullAddress = showFull;
                        tabAutoSize = autoSize;
                        tabFixedWidth = fixedWidth;
                        tabFixedHeight = fixedHeight;
                }
        };
	CESettings GetCESettings();
	void WriteCESettings(CESettings& toWrite);
	HRESULT GetCurrentFolderPidl(CComPtr<IShellBrowser> pShellBrowser, PIDLIST_ABSOLUTE *pidlOut);
	HRESULT FixExplorerSizes(HWND explorerChild);
	HRESULT FixExplorerSizesIfNecessary(HWND explorerChild);
}

#endif