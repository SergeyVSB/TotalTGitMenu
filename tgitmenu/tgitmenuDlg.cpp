
// tgitmenuDlg.cpp : implementation file
//

#include "stdafx.h"
#include "tgitmenu.h"
#include "tgitmenuDlg.h"
#include "afxdialogex.h"

#include <Shobjidl.h>
#include <comdef.h>
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CtgitmenuDlg dialog



CtgitmenuDlg::CtgitmenuDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_TGITMENU_DIALOG, pParent)
{

	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CtgitmenuDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CtgitmenuDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_CONTEXTMENU()
	ON_WM_WINDOWPOSCHANGING()
END_MESSAGE_MAP()


// CtgitmenuDlg message handlers

BOOL CtgitmenuDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();


	//ModifyStyle(WS_VISIBLE, 0);
	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here

	this->ShowWindow(SW_HIDE);

	PostMessageW(WM_CONTEXTMENU, (WPARAM)GetSafeHwnd(), MAKELPARAM(0, 0));

	return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

#pragma warning(suppress: 26434)
void CtgitmenuDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
#pragma warning(suppress: 26434)
HCURSOR CtgitmenuDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// {10A0FDD2-B0C0-4CD4-A7AE-E594CE3B91C8}
static const GUID gTgit =
{ 0x10A0FDD2, 0xB0C0, 0x4CD4,{ 0xA7, 0xAE, 0xE5, 0x94, 0xCE, 0x3B, 0x91, 0xC8 } };


#define SCRATCH_QCM_FIRST 1
#define SCRATCH_QCM_LAST  0x7FFF

bool MergeMenu(CMenu * pMenuDestination, const CMenu * pMenuAdd, bool bTopLevel = false)
{
	// Abstract:
	//      Merges two menus.
	//
	// Parameters:
	//      pMenuDestination    - [in, retval] destination menu handle
	//      pMenuAdd            - [in] menu to merge
	//      bTopLevel           - [in] indicator for special top level behavior
	//
	// Return value:
	//      <false> in case of error.
	//
	// Comments:
	//      This function calles itself recursivley. If bTopLevel is set to true,
	//      we append popups at top level or we insert before <Window> or <Help>.

	// get the number menu items in the menus
	const int iMenuAddItemCount = pMenuAdd->GetMenuItemCount();
	int iMenuDestItemCount = pMenuDestination->GetMenuItemCount();

	// if there are no items return
	if (iMenuAddItemCount == 0)
		return true;

	// if we are not at top level and the destination menu is not empty
	// -> we append a seperator
	if (!bTopLevel && iMenuDestItemCount > 0)
		pMenuDestination->AppendMenu(MF_SEPARATOR);

	// iterate through the top level of 
	for (int iLoop = 0; iLoop < iMenuAddItemCount; iLoop++)
	{
		// get the menu string from the add menu
		CString sMenuAddString;
		pMenuAdd->GetMenuString(iLoop, sMenuAddString, MF_BYPOSITION);

		// try to get the submenu of the current menu item
		CMenu* pSubMenu = pMenuAdd->GetSubMenu(iLoop);

		// check if we have a sub menu
		if (!pSubMenu)
		{
			// normal menu item
			// read the source and append at the destination
			const auto nState = pMenuAdd->GetMenuState(iLoop, MF_BYPOSITION);
			const auto nItemID = pMenuAdd->GetMenuItemID(iLoop);

			if (pMenuDestination->AppendMenu(nState, nItemID, sMenuAddString))
			{
				// menu item added, don't forget to correct the item count
				iMenuDestItemCount++;
				// copy menu bitmap
				MENUITEMINFO mif;
				mif.cbSize = sizeof(MENUITEMINFO);
				mif.fMask = MIIM_BITMAP;
				if (const_cast<CMenu*>(pMenuAdd)->GetMenuItemInfo(iLoop, &mif, TRUE))
				{
					pMenuDestination->SetMenuItemInfo(iMenuDestItemCount, &mif, TRUE);
				}
			}
			else
			{
				TRACE("MergeMenu: AppendMenu failed!\n");
				return false;
			}
		}
		else
		{
			// create or insert a new popup menu item

			// default insert pos is like ap
			int iInsertPosDefault = -1;

			// if we are at top level merge into existing popups rather than
			// creating new ones
			if (bTopLevel)
			{
				ASSERT(sMenuAddString != "&?" && sMenuAddString != "?");
				CString sAdd(sMenuAddString);
				sAdd.Remove('&');  // for comparison of menu items supress '&'
				bool bAdded = false;

				// try to find existing popup
				for (int iLoop1 = 0; iLoop1 < iMenuDestItemCount; iLoop1++)
				{
					// get the menu string from the destination menu
					CString sDest;
					pMenuDestination->GetMenuString(iLoop1, sDest, MF_BYPOSITION);
					sDest.Remove('&'); // for a better compare (s.a.)

					if (sAdd == sDest)
					{
						// we got a hit -> merge the two popups
						// try to get the submenu of the desired destination menu item
						CMenu* pSubMenuDest = pMenuDestination->GetSubMenu(iLoop1);

						if (pSubMenuDest)
						{
							// merge the popup recursivly and continue with outer for loop
							if (!MergeMenu(pSubMenuDest, pSubMenu))
								return false;

							bAdded = true;
							break;
						}
					}

					// alternativ insert before <Window> or <Help>
					if (iInsertPosDefault == -1 && (sDest == "Window" || sDest == "?" || sDest == "Help"))
						iInsertPosDefault = iLoop1;

				}

				if (bAdded)
				{
					// menu added, so go on with loop over pMenuAdd's top level
					continue;
				}
			}

			// if the top level search did not find a position append the menu
			if (iInsertPosDefault == -1)
				iInsertPosDefault = pMenuDestination->GetMenuItemCount();

			// create a new popup and insert before <Window> or <Help>
			CMenu NewPopupMenu;
			if (!NewPopupMenu.CreatePopupMenu())
			{
				TRACE("MergeMenu: CreatePopupMenu failed!\n");
				return false;
			}

			// merge the new popup recursivly
			if (!MergeMenu(&NewPopupMenu, pSubMenu))
				return false;

			// insert the new popup menu into the destination menu
			HMENU hNewMenu = NewPopupMenu.GetSafeHmenu();

			if (pMenuDestination->InsertMenu(iInsertPosDefault,
				MF_BYPOSITION | MF_POPUP | MF_ENABLED,
				(UINT)hNewMenu, sMenuAddString))
			{
				// menu item added, don't forget to correct the item count
				iMenuDestItemCount++;
				// copy menu bitmap
				MENUITEMINFO mif;
				mif.cbSize = sizeof(MENUITEMINFO);
				mif.fMask = MIIM_BITMAP;
				if (const_cast<CMenu*>(pMenuAdd)->GetMenuItemInfo(iLoop, &mif, TRUE))
				{
					pMenuDestination->SetMenuItemInfo(iMenuDestItemCount, &mif, TRUE);
				}
			}
			else
			{
				TRACE("MergeMenu: InsertMenu failed!\n");
				return false;
			}

			// don't destroy the new menu       
			NewPopupMenu.Detach();
		}
	}

	return true;
}


#define PRINT_FAILURE_HR(expr, hr) wprintf(L"%S failed with 0x%x\n", expr, hr)

#define ABORT_IF_FAILED(expr) \
do { \
    HRESULT _hr = (expr);\
    if (!SUCCEEDED(_hr)) { \
        PRINT_FAILURE_HR(#expr, _hr); \
        return; \
    } \
} while (0);

#define ABORT_IF_FAILED_HR(expr) \
do { \
    HRESULT _hr = (expr);\
    if (!SUCCEEDED(_hr)) { \
        PRINT_FAILURE_HR(#expr, _hr); \
        return _hr; \
    } \
} while (0);


static HRESULT CreateIDataObject(CComHeapPtr<ITEMIDLIST>& pidl, CComPtr<IDataObject>& dataObject, PCWSTR FileName)
{
	CComPtr<IShellFolder> desktopFolder;
	ABORT_IF_FAILED_HR(SHGetDesktopFolder(&desktopFolder));
	ABORT_IF_FAILED_HR(desktopFolder->ParseDisplayName(NULL, NULL, const_cast<LPWSTR>(FileName), NULL, &pidl, NULL));

	CComPtr<IShellFolder> shellFolder;
	PCUITEMID_CHILD childs;
	ABORT_IF_FAILED_HR(SHBindToParent(pidl, IID_PPV_ARGS(&shellFolder), &childs));
	ABORT_IF_FAILED_HR(shellFolder->GetUIObjectOf(NULL, 1, &childs, IID_IDataObject, NULL, reinterpret_cast<PVOID*>(&dataObject)));
	return S_OK;
}

static HRESULT GetUIObjectOfFile(HWND hwnd, LPCWSTR pszPath, REFIID riid, void **ppv)
{
	*ppv = nullptr;
	HRESULT hr;
	CComHeapPtr<ITEMIDLIST> pidl;
	SFGAOF sfgao;
	if (SUCCEEDED(hr = SHParseDisplayName(pszPath, nullptr, &pidl, 0, &sfgao)))
	{
		CComPtr<IShellFolder> psf;
		PCUITEMID_CHILD  pidlChild;
		//IID_PPV_ARGS(&psf)
		if (SUCCEEDED(hr = SHBindToParent(pidl, IID_PPV_ARGS(&psf), &pidlChild)))
		{
			hr = psf->GetUIObjectOf(hwnd, 1, &pidlChild, riid, nullptr, ppv);
		}
	}
	return hr;
}


void CtgitmenuDlg::OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/)
{
	CComPtr<IContextMenu> pMenu;
	if (SUCCEEDED(pMenu.CoCreateInstance(gTgit)))
	{
		CString sFile;

		if (__argc < 2)
		{
			AfxMessageBox(L"No file specified!");
			return;
		}

		if (__argc > 2)
		{
			AfxMessageBox(L"Multiple files not supported!!");
			return;
		}

		if (__argc == 2)
		{
			sFile = __wargv[1];
			if (_tcscmp(PathFindExtension(sFile), _T(".lnk")) == 0)
			{
				AfxMessageBox(L"ShellLinks not supported!");
				return;
			}
		}

		CComPtr<IDataObject> pdo;
		HRESULT hr;
		//CComHeapPtr<ITEMIDLIST> pidl;

		if (SUCCEEDED(hr = GetUIObjectOfFile(GetSafeHwnd(), sFile, IID_PPV_ARGS(&pdo))))
			//if (SUCCEEDED(hr = CreateIDataObject(pidl, pdo, sFile)))
		{
			CMenu mMenu;
			mMenu.CreatePopupMenu();
			CComPtr<IShellExtInit> pInit;
			if (SUCCEEDED(hr = pMenu->QueryInterface(&pInit)))
			{
				if (SUCCEEDED(pInit->Initialize(nullptr, pdo, NULL)))
				{
					UINT uFlags = (GetKeyState(VK_SHIFT) < 0) ? CMF_EXTENDEDVERBS : CMF_NORMAL;
					if (_tcscmp(PathFindExtension(sFile), _T(".lnk")) == 0)
					{
						uFlags |= CMF_VERBSONLY;
					}
					if (SUCCEEDED(pMenu->QueryContextMenu(mMenu.GetSafeHmenu(), 0, SCRATCH_QCM_FIRST, SCRATCH_QCM_LAST, uFlags)))
					{

						mMenu.DeleteMenu(0, MF_BYPOSITION);
						mMenu.DeleteMenu(mMenu.GetMenuItemCount() - 1, MF_BYPOSITION);

						const auto nSub = mMenu.GetMenuItemCount() - 1;
						if (const auto * const pSub = mMenu.GetSubMenu(nSub))
						{
							MergeMenu(&mMenu, pSub);
							mMenu.DeleteMenu(nSub, MF_BYPOSITION);
						}
						// delete first separator
						if (nSub == 0)
						{
							mMenu.DeleteMenu(0, MF_BYPOSITION);
						}

						CPoint mPointCurrent;
						::GetCursorPos(&mPointCurrent);

						const auto iCmd = mMenu.TrackPopupMenu(TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RETURNCMD | TPM_NONOTIFY, mPointCurrent.x, mPointCurrent.y, this);
						if (iCmd > 0)
						{
							CMINVOKECOMMANDINFOEX info = { 0 };
							info.cbSize = sizeof(info);
							info.fMask = CMIC_MASK_UNICODE | CMIC_MASK_PTINVOKE;
							if (GetKeyState(VK_CONTROL) < 0) {
								info.fMask |= CMIC_MASK_CONTROL_DOWN;
							}
							if (GetKeyState(VK_SHIFT) < 0) {
								info.fMask |= CMIC_MASK_SHIFT_DOWN;
							}
							info.hwnd = GetSafeHwnd();
							info.lpVerb = MAKEINTRESOURCEA(iCmd - SCRATCH_QCM_FIRST);
							info.lpVerbW = MAKEINTRESOURCEW(iCmd - SCRATCH_QCM_FIRST);
							info.nShow = SW_SHOWNORMAL;
							info.ptInvoke = mPointCurrent;
							pMenu->InvokeCommand(reinterpret_cast<LPCMINVOKECOMMANDINFO>(&info));
						}
					}
				}
			}
		}
	}
	PostMessage(WM_CLOSE);
}


void CtgitmenuDlg::OnWindowPosChanging(WINDOWPOS* lpwndpos)
{
	lpwndpos->flags &= ~SWP_SHOWWINDOW;
	CDialogEx::OnWindowPosChanging(lpwndpos);
}
