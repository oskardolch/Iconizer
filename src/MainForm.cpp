/*
    Copyright (C) 2009  Pavel Krejcir

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

#include <stdio.h>
#include <tchar.h>
#include "MainForm.hpp"
#include "Iconizer.rh"
#include "../Common/COMUtils.hpp"
#include "../Common/XMLUtils.hpp"
#include "Render.hpp"

static int spwidth = 4;

CMainForm::CMainForm(HINSTANCE hInstance, DLGPROC lpDialogFunc, HWND *phWnd,
    HACCEL *phAccTable)
{
    m_hInstance = hInstance;
    m_sAppName = GetCommandLine();
    m_bHasChanged = FALSE;
    m_sCurFileName[0] = 0;
    m_uiLBDragMsg = 0;
    m_iDragItem = -1;

    WNDCLASSEX WC;
    WC.cbSize = sizeof(WNDCLASSEX);
    WC.style = CS_SAVEBITS | CS_DBLCLKS;
    WC.hInstance = m_hInstance;
    WC.lpfnWndProc = DefDlgProc;
    WC.cbClsExtra = 0;
    WC.cbWndExtra = DLGWINDOWEXTRA;
    WC.hCursor = LoadCursor(0, IDC_ARROW);
    WC.hbrBackground = (HBRUSH__*)COLOR_BTNSHADOW;
    WC.hIcon = LoadIcon(m_hInstance, _T("MAINICON"));
    //WC.hIcon = LoadIcon(0, IDI_APPLICATION);
    WC.lpszMenuName = _T("MAINMENU");
    WC.lpszClassName = _T("ICONIZERWND");
    WC.hIconSm = 0;
    RegisterClassEx(&WC);

    *phWnd = CreateDialogParam(m_hInstance, _T("MAINFORM"), 0, lpDialogFunc,
        (LPARAM)this);
    *phAccTable = LoadAccelerators(m_hInstance, MAKEINTRESOURCE(1));

    m_AddImgDlg = new CAddImageDlg(m_hInstance);
    m_EditImgDlg = new CEditImageDlg(m_hInstance);

    return;
}

CMainForm::~CMainForm()
{
    delete m_EditImgDlg;
    delete m_AddImgDlg;
    UnregisterClass(_T("ICONIZERWND"), m_hInstance);
    //free(m_sAppName);
    return;
}

void CMainForm::RestoreSettings(HWND hWnd)
{
#ifdef UNICODE
    char* xname = WideCharToStr(&m_sAppName[1]);
#else
    char* xname = (char*)malloc(strlen(m_sAppName));
    strcpy(xname, &m_sAppName[1]);
#endif
    int n = strlen(xname);
    xname[n - 2] = 0;
    xname[n - 3] = 'l';
    xname[n - 4] = 'm';
    xname[n - 5] = 'x';

    int x = 100;
    int y = 100;

    CXMLReader* fRdr = new CXMLReader(xname);
    free(xname);
    IXMLDOMElement* Elem = fRdr->OpenSection("MainForm");
    if(Elem)
    {
        int i;
        if(fRdr->GetIntValue(Elem, "Left", &i)) x = i;
        if(fRdr->GetIntValue(Elem, "Top", &i)) y = i;

        LPTSTR buf;
        if(fRdr->GetStringValue(Elem, "LastPath", &buf))
        {
            _tcscpy(m_sLastOpenPath, buf);
            FreeTString(&buf);
        }

        Elem->Release();
    }
    SetWindowPos(hWnd, 0, x, y, 0, 0, SWP_NOSIZE);
    delete fRdr;
}

INT_PTR CMainForm::WMInitDialog(HWND hWnd, HWND hwndFocus, LPARAM lInitParam)
{
    SetWindowLongPtr(hWnd, DWLP_USER, lInitParam);

    RECT R;
    HWND wnd = GetDlgItem(hWnd, CTRL_ICONSLISTBOX);
    GetWindowRect(wnd, &R);
    MapWindowPoints(NULL, hWnd, (LPPOINT)&R, 2);
    int dx = R.left;
    int dy = R.top;
    GetClientRect(wnd, &R);
    R.bottom = 256;
    DWORD style = GetWindowLongPtr(wnd, GWL_STYLE);
    DWORD exstyle = GetWindowLongPtr(wnd, GWL_EXSTYLE);
    AdjustWindowRectEx(&R, style, FALSE, exstyle);
    SetWindowPos(wnd, NULL, 0, 0, R.right - R.left, R.bottom - R.top, SWP_NOMOVE);
    MakeDragList(wnd);
    m_uiLBDragMsg = RegisterWindowMessage(DRAGLISTMSGSTRING);

    int height = R.bottom - R.top + 2*dy;

    wnd = GetDlgItem(hWnd, CTRL_IMAGE);
    GetWindowRect(wnd, &R);
    MapWindowPoints(NULL, hWnd, (LPPOINT)&R, 2);
    int width = R.left;

    GetClientRect(wnd, &R);
    R.right = 256;
    R.bottom = 256;
    style = GetWindowLongPtr(wnd, GWL_STYLE);
    exstyle = GetWindowLongPtr(wnd, GWL_EXSTYLE);
    AdjustWindowRectEx(&R, style, FALSE, exstyle);
    SetWindowPos(wnd, NULL, 0, 0, R.right - R.left, R.bottom - R.top, SWP_NOMOVE);
/*LONG bsunits = GetDialogBaseUnits();
wchar_t buf[64];
swprintf(buf, L"%d, %d", ((R.right - R.left)*4)/LOWORD(bsunits),
((R.bottom - R.top)*8)/HIWORD(bsunits));
MessageBox(0, buf, L"dimen", MB_OK);*/

    width += R.right - R.left + dx;
    R.left = 0;
    R.right = width;
    R.top = 0;
    R.bottom = height;
    style = GetWindowLongPtr(hWnd, GWL_STYLE);
    exstyle = GetWindowLongPtr(hWnd, GWL_EXSTYLE);
    AdjustWindowRectEx(&R, style, TRUE, exstyle);

    SetWindowPos(hWnd, NULL, 0, 0, R.right - R.left, R.bottom - R.top, SWP_NOMOVE);

    RestoreSettings(hWnd);
    EnableMenuItems(hWnd);
    return(1);
}

void CMainForm::SaveSettings(HWND hWnd)
{
#ifdef UNICODE
    char* xname = WideCharToStr(&m_sAppName[1]);
#else
    char* xname = (char*)malloc(strlen(m_sAppName));
    strcpy(xname, &m_sAppName[1]);
#endif
    int n = strlen(xname);
    xname[n - 2] = 0;
    xname[n - 3] = 'l';
    xname[n - 4] = 'm';
    xname[n - 5] = 'x';
    CXMLWritter* fWrit = new CXMLWritter(xname);
    fWrit->WriteComment("Iconizer Workspace Settings");
    fWrit->CreateRoot("Settings");

    IXMLDOMElement* Elem = NULL;
    RECT R;
    GetWindowRect(hWnd, &R);

    Elem = fWrit->CreateSection("MainForm");
    fWrit->AddIntValue(Elem, "Left", R.left);
    fWrit->AddIntValue(Elem, "Top", R.top);

    fWrit->AddStringValue(Elem, "LastPath", m_sLastOpenPath);

    Elem->Release();

    fWrit->Save();
    delete fWrit;
    return;
}

INT_PTR CMainForm::WMCommand(HWND hWnd, WORD wNotifyCode, WORD wID, HWND hwndCtl)
{
    switch(wID)
    {
    case IDM_FILENEW:
        return(NewFile(hWnd));
    case IDM_FILEOPEN:
        return(OpenFile(hWnd));
    case IDM_FILESAVE:
        return(SaveFile(hWnd));
    case IDM_FILESAVEAS:
        return(SaveFileAs(hWnd));
    case IDM_FILEEXIT:
        PostMessage(hWnd, WM_CLOSE, 0, 0);
        return(1);
    case IDM_IMAGEADD:
        return(AddImage(hWnd));
    case IDM_IMAGELOAD:
        MessageBox(hWnd, _T("Load image"), _T("Not implemented yet"), MB_OK);
        return(1);
    case IDM_IMAGESAVE:
        //MessageBox(hWnd, _T("Save image"), _T("Not implemented yet"), MB_OK);
        return(ImageSave(hWnd));
    case IDM_IMAGEDEL:
        return(ImageDelete(hWnd));
    case IDM_IMAGEEDIT:
        return(ImageEdit(hWnd));
    case IDM_HELPCONTENT:
        MessageBox(hWnd, _T("Help content"), _T("Not implemented yet"), MB_OK);
        return(1);
    case IDM_HELPONHELP:
        MessageBox(hWnd, _T("Help on help"), _T("Not implemented yet"), MB_OK);
        return(1);
    case IDM_HELPABOUT:
        MessageBox(hWnd, _T("Help about"), _T("Not implemented yet"), MB_OK);
        return(1);
    case CTRL_ICONSLISTBOX:
        switch(wNotifyCode)
        {
        case LBN_SELCHANGE:
            return(ListBoxChange(hWnd, hwndCtl));
        default:
            return(0);
        }
    default:
        return(0);
    }
}

INT_PTR CMainForm::WMSysCommand(HWND hWnd, WPARAM uCmdType, WORD xPos, WORD yPos)
{
    switch(uCmdType)
    {
    case SC_CLOSE:
        PostMessage(hWnd, WM_CLOSE, 0, 0);
        return(1);
    default:
        return(0);
    }
}

INT_PTR CMainForm::WMDestroy(HWND hWnd)
{
    PostMessage(hWnd, WM_CLOSE, 0, 0);
    return(0);
}

INT_PTR CMainForm::WMClose(HWND hWnd)
{
    if(ExitApp(hWnd)) PostQuitMessage(0);
    return(0);
}

INT_PTR CMainForm::WMUnknownMessage(HWND hWnd, UINT uMsg, WPARAM wParam,
    LPARAM lParam)
{
    if(uMsg == m_uiLBDragMsg)
        return(ListBoxDragMsg(hWnd, (int)wParam, (LPDRAGLISTINFO)lParam));
    return(0);
}

INT_PTR CMainForm::OpenFile(HWND hWnd)
{
    TCHAR namebuf[512];
    TCHAR filterbuf[128];
    namebuf[0] = 0;
    LoadString(m_hInstance, IDS_ICO_FILTER, filterbuf, 128);
    for(int i = 0; i < 128; i++)
        if(filterbuf[i] == 1) filterbuf[i] = 0;

    OPENFILENAME ofn;
    ofn.lStructSize = sizeof(OPENFILENAME);
    ofn.hwndOwner = hWnd;
    ofn.hInstance = m_hInstance;
    ofn.lpstrFilter = filterbuf;
    ofn.lpstrCustomFilter = NULL;
    ofn.nMaxCustFilter = 0;
    ofn.nFilterIndex = 0;
    ofn.lpstrFile = namebuf;
    ofn.nMaxFile = 512;
    ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.lpstrInitialDir = m_sLastOpenPath;
    ofn.lpstrTitle = NULL;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_NONETWORKBUTTON |
        OFN_EXPLORER | OFN_ENABLESIZING;
    ofn.nFileOffset = 0;
    ofn.nFileExtension = 0;
    ofn.lpstrDefExt = NULL;
    ofn.lCustData = 0;
    ofn.lpfnHook = NULL;
    ofn.lpTemplateName = NULL;
#ifndef _MSCOMP_
    ofn.pvReserved = NULL;
    ofn.dwReserved = 0;
    ofn.FlagsEx = 0;
#endif
    if(GetOpenFileName(&ofn)) OpenIconFile(hWnd, namebuf);
    return(TRUE);
}

BOOL CMainForm::ExitApp(HWND hWnd)
{
    if(!CheckChangedFile(hWnd)) return(0);

    ClearListBox(hWnd);
    SaveSettings(hWnd);
    return(1);
}

void CMainForm::EnableMenuItems(HWND hWnd)
{
    HMENU hmnu = GetMenu(hWnd);
    EnableMenuItem(hmnu, IDM_IMAGEEDIT, MF_BYCOMMAND | MF_GRAYED);
    BOOL bIsBmp = FALSE;
    if(m_bHasChanged || _tcslen(m_sCurFileName) > 0)
    {
        EnableMenuItem(hmnu, IDM_FILESAVE, MF_BYCOMMAND | MF_ENABLED);
        EnableMenuItem(hmnu, IDM_FILESAVEAS, MF_BYCOMMAND | MF_ENABLED);
        EnableMenuItem(hmnu, IDM_IMAGEADD, MF_BYCOMMAND | MF_ENABLED);
        int i = SendDlgItemMessage(hWnd, CTRL_ICONSLISTBOX, LB_GETCURSEL, 0, 0);
        if(i < 0)
        {
            EnableMenuItem(hmnu, IDM_IMAGELOAD, MF_BYCOMMAND | MF_GRAYED);
            EnableMenuItem(hmnu, IDM_IMAGESAVE, MF_BYCOMMAND | MF_GRAYED);
            EnableMenuItem(hmnu, IDM_IMAGEDEL, MF_BYCOMMAND | MF_GRAYED);
        }
        else
        {
            EnableMenuItem(hmnu, IDM_IMAGELOAD, MF_BYCOMMAND | MF_ENABLED);
            EnableMenuItem(hmnu, IDM_IMAGESAVE, MF_BYCOMMAND | MF_ENABLED);
            EnableMenuItem(hmnu, IDM_IMAGEDEL, MF_BYCOMMAND | MF_ENABLED);
            LPICONDATA pid = (LPICONDATA)SendDlgItemMessage(hWnd,
                CTRL_ICONSLISTBOX, LB_GETITEMDATA, i, 0);

            // we don't allow editing of PNG images
            // even if it is a palette color image
            bIsBmp = pid->icoimg.icHeader.biSize == sizeof(BITMAPINFOHEADER);
            if(pid->icdi.wBitCount < 16 && bIsBmp)
                EnableMenuItem(hmnu, IDM_IMAGEEDIT, MF_BYCOMMAND | MF_ENABLED);
        }
    }
    else
    {
        EnableMenuItem(hmnu, IDM_FILESAVE, MF_BYCOMMAND | MF_GRAYED);
        EnableMenuItem(hmnu, IDM_FILESAVEAS, MF_BYCOMMAND | MF_GRAYED);
        EnableMenuItem(hmnu, IDM_IMAGEADD, MF_BYCOMMAND | MF_GRAYED);
        EnableMenuItem(hmnu, IDM_IMAGELOAD, MF_BYCOMMAND | MF_GRAYED);
        EnableMenuItem(hmnu, IDM_IMAGESAVE, MF_BYCOMMAND | MF_GRAYED);
        EnableMenuItem(hmnu, IDM_IMAGEDEL, MF_BYCOMMAND | MF_GRAYED);
    }
    return;
}

void CMainForm::OpenIconFile(HWND hWnd, LPTSTR sFileName)
{
    if(CheckChangedFile(hWnd))
    {
        if(ReadIconFile(hWnd, TRUE, sFileName))
        {
            _tcscpy(m_sCurFileName, sFileName);
            m_bHasChanged = FALSE;
        }
        EnableMenuItems(hWnd);
    }
    return;
}

BOOL CMainForm::CheckChangedFile(HWND hWnd)
{
    if(m_bHasChanged)
    {
        TCHAR cap[64];
        TCHAR msg[128];
        LoadString(m_hInstance, IDS_DLG_WARN, cap, 64);
        LoadString(m_hInstance, IDS_MSG_ICOCHANGED, msg, 128);
        int conf = MessageBox(hWnd, msg, cap, MB_YESNOCANCEL | MB_ICONQUESTION);
        if(conf == IDYES)
        {
            if(_tcslen(m_sCurFileName) > 0) SaveFile(hWnd);
            else return(SaveFileAs(hWnd));
        }
        else if(conf == IDCANCEL) return(FALSE);
    }
    return(TRUE);
}

INT_PTR CMainForm::SaveFile(HWND hWnd)
{
    if(m_bHasChanged)
    {
        if(_tcslen(m_sCurFileName) > 0)
        {
            SaveIconFile(hWnd);
            m_bHasChanged = FALSE;
        }
        else SaveFileAs(hWnd);
    }
    return(TRUE);
}

BOOL CMainForm::SaveFileAs(HWND hWnd)
{
    TCHAR namebuf[512];
    TCHAR filterbuf[128];
    namebuf[0] = 0;
    LoadString(m_hInstance, IDS_ICO_FILTER, filterbuf, 128);
    for(int i = 0; i < 128; i++)
        if(filterbuf[i] == 1) filterbuf[i] = 0;

    OPENFILENAME ofn;
    ofn.lStructSize = sizeof(OPENFILENAME);
    ofn.hwndOwner = hWnd;
    ofn.hInstance = m_hInstance;
    ofn.lpstrFilter = filterbuf;
    ofn.lpstrCustomFilter = NULL;
    ofn.nMaxCustFilter = 0;
    ofn.nFilterIndex = 0;
    ofn.lpstrFile = namebuf;
    ofn.nMaxFile = 512;
    ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.lpstrInitialDir = m_sLastOpenPath;
    ofn.lpstrTitle = NULL;
    ofn.Flags = OFN_OVERWRITEPROMPT | OFN_NONETWORKBUTTON | OFN_EXPLORER |
        OFN_ENABLESIZING;
    ofn.nFileOffset = 0;
    ofn.nFileExtension = 0;
    ofn.lpstrDefExt = _T("ico");
    ofn.lCustData = 0;
    ofn.lpfnHook = NULL;
    ofn.lpTemplateName = NULL;
#ifndef _MSCOMP_
    ofn.pvReserved = NULL;
    ofn.dwReserved = 0;
    ofn.FlagsEx = 0;
#endif
    if(GetSaveFileName(&ofn))
    {
        _tcscpy(m_sCurFileName, namebuf);
        SaveIconFile(hWnd);
        m_bHasChanged = FALSE;
        return(TRUE);
    }
    return(FALSE);
}

BOOL CMainForm::ReadIconFile(HWND hWnd, BOOL bClearLB, LPTSTR sFileName)
{
    BOOL bRes = FALSE;
    FILE *fp = _tfopen(sFileName, _T("rb"));
    if(fp)
    {
        TCHAR sIcoType[64];
        HWND lb = GetDlgItem(hWnd, CTRL_ICONSLISTBOX);
        if(bClearLB) ClearListBox(hWnd);

        ICONDIR icd;
        int idx, width, height;;

        fread(&icd, 1, 6, fp);
        if(icd.idReserved == 0 && icd.idType == 1)
        {
            bRes = icd.idCount > 0;
            int iListBase = SendMessage(lb, LB_GETCOUNT, 0, 0);

            ICONDIRENTRY icdi;
            LPICONDATA lpid = NULL;
            for(int i = 0; i < icd.idCount; i++)
            {
                fread(&icdi, 1, sizeof(icdi), fp);
                lpid = (LPICONDATA)malloc(sizeof(HBITMAP) + sizeof(icdi) +
                    icdi.dwBytesInRes);
                lpid->hPreview = NULL;
                memcpy(&lpid->icdi, &icdi, sizeof(icdi));
                if(icdi.bWidth > 0) width = icdi.bWidth;
                else width = 256;
                if(icdi.bHeight > 0) height = icdi.bHeight;
                else height = 256;
                _stprintf(sIcoType, _T("%dx%d - %d bpp"), width, height,
                    icdi.wBitCount);
                idx = SendMessage(lb, LB_ADDSTRING, 0, (LPARAM)sIcoType);
                SendMessage(lb, LB_SETITEMDATA, (WPARAM)idx, (LPARAM)lpid);
            }
            int n = SendMessage(lb, LB_GETCOUNT, 0, 0);
            for(int i = iListBase; i < n; i++)
            {
                lpid = (LPICONDATA)SendMessage(lb, LB_GETITEMDATA, i, 0);
                fseek(fp, lpid->icdi.dwImageOffset, SEEK_SET);
                fread(&lpid->icoimg, 1, lpid->icdi.dwBytesInRes, fp);
            }
        }
        fclose(fp);
    }
    return(bRes);
}

void CMainForm::SaveIconFile(HWND hWnd)
{
    FILE *fp = _tfopen(m_sCurFileName, _T("wb"));
    if(fp)
    {
        HWND lb = GetDlgItem(hWnd, CTRL_ICONSLISTBOX);

        ICONDIR icd;
        icd.idReserved = 0;
        icd.idType = 1;
        icd.idCount = SendMessage(lb, LB_GETCOUNT, 0, 0);
        fwrite(&icd, 1, 6, fp);

        LPICONDATA lpid;
        long imgoffs = 6 + icd.idCount*sizeof(ICONDIRENTRY);
        for(int i = 0; i < icd.idCount; i++)
        {
            lpid = (LPICONDATA)SendMessage(lb, LB_GETITEMDATA, i, 0);
            lpid->icdi.dwImageOffset = imgoffs;
            imgoffs += lpid->icdi.dwBytesInRes;
            fwrite(&lpid->icdi, 1, sizeof(ICONDIRENTRY), fp);
        }
        for(int i = 0; i < icd.idCount; i++)
        {
            lpid = (LPICONDATA)SendMessage(lb, LB_GETITEMDATA, i, 0);
            fwrite(&lpid->icoimg, 1, lpid->icdi.dwBytesInRes, fp);
        }
        fclose(fp);
    }
    return;
}

INT_PTR CMainForm::ListBoxChange(HWND hWnd, HWND hwndCtl)
{
    EnableMenuItems(hWnd);
    HWND wnd = GetDlgItem(hWnd, CTRL_IMAGE);
    int i = SendMessage(hwndCtl, LB_GETCURSEL, 0, 0);
    if(i > -1)
    {
        LPICONDATA lpid = (LPICONDATA)SendMessage(hwndCtl, LB_GETITEMDATA, i, 0);
        BOOL bIsBmp = lpid->icoimg.icHeader.biSize == sizeof(BITMAPINFOHEADER);
        int w, h;

        if(lpid->hPreview)
        {
            SendMessage(wnd, STM_SETIMAGE, IMAGE_BITMAP, (LPARAM)lpid->hPreview);
        }
        else if(bIsBmp)
        {
            w = lpid->icoimg.icHeader.biWidth;
            h = lpid->icoimg.icHeader.biHeight;
            if(h > lpid->icdi.bHeight) h /= 2;
            lpid->hPreview = RenderIcon(wnd, &lpid->icoimg, w, h);
        }
        else
        {
            //w = lpid->icdi.bWidth;
            //h = lpid->icdi.bHeight;
            lpid->hPreview = RenderPNG(wnd, &lpid->icoimg,
                lpid->icdi.dwBytesInRes, NULL, NULL);
        }
    }
    else SendMessage(wnd, STM_SETIMAGE, IMAGE_BITMAP, (LPARAM)NULL);
    return(1);
}

void CMainForm::ClearListBox(HWND hWnd)
{
    HWND lb = GetDlgItem(hWnd, CTRL_ICONSLISTBOX);
    int n = SendMessage(lb, LB_GETCOUNT, 0, 0);
    LPICONDATA lpid = NULL;
    for(int i = 0; i < n; i++)
    {
        lpid = (LPICONDATA)SendMessage(lb, LB_GETITEMDATA, i, 0);
        if(lpid->hPreview) DeleteObject(lpid->hPreview);
        free(lpid);
    }
    SendMessage(lb, LB_RESETCONTENT, 0, 0);
    SendDlgItemMessage(hWnd, CTRL_IMAGE, STM_SETIMAGE, IMAGE_BITMAP,
        (LPARAM)NULL);
    return;
}

INT_PTR CMainForm::ListBoxDragMsg(HWND hWnd, int idCtl, LPDRAGLISTINFO pDragInfo)
{
    switch(pDragInfo->uNotification)
    {
    case DL_BEGINDRAG:
        return(ListBoxBeginDrag(hWnd, idCtl, pDragInfo->hWnd,
            pDragInfo->ptCursor));
    case DL_CANCELDRAG:
        return(ListBoxCancelDrag(hWnd, idCtl, pDragInfo->hWnd,
            pDragInfo->ptCursor));
    case DL_DRAGGING:
        return(ListBoxDragging(hWnd, idCtl, pDragInfo->hWnd,
            pDragInfo->ptCursor));
    case DL_DROPPED:
        return(ListBoxDropped(hWnd, idCtl, pDragInfo->hWnd,
            pDragInfo->ptCursor));
    default:
        return(0);
    }
}

INT_PTR CMainForm::ListBoxBeginDrag(HWND hWnd, int idCtl, HWND hwndCtl,
    POINT ptCursor)
{
    m_iDragItem = LBItemFromPt(hwndCtl, ptCursor, TRUE);
    SetWindowLongPtr(hWnd, DWLP_MSGRESULT, TRUE);
    return(TRUE);
}

INT_PTR CMainForm::ListBoxCancelDrag(HWND hWnd, int idCtl, HWND hwndCtl,
    POINT ptCursor)
{
    m_iDragItem = -1;
    DrawInsert(hWnd, hwndCtl, -1);
    return(TRUE);
}

INT_PTR CMainForm::ListBoxDragging(HWND hWnd, int idCtl, HWND hwndCtl,
    POINT ptCursor)
{
    int ipos = LBItemFromPt(hwndCtl, ptCursor, TRUE);
    DrawInsert(hWnd, hwndCtl, ipos);
    SetWindowLongPtr(hWnd, DWLP_MSGRESULT, DL_MOVECURSOR);
    return(TRUE);
}

INT_PTR CMainForm::ListBoxDropped(HWND hWnd, int idCtl, HWND hwndCtl,
    POINT ptCursor)
{
    int ipos = LBItemFromPt(hwndCtl, ptCursor, FALSE);
    if(m_iDragItem > -1 && m_iDragItem != ipos)
    {
        TCHAR buf[64];
        SendMessage(hwndCtl, LB_GETTEXT, m_iDragItem, (LPARAM)buf);
        LONG_PTR pData = SendMessage(hwndCtl, LB_GETITEMDATA, m_iDragItem, 0);
        SendMessage(hwndCtl, LB_DELETESTRING, m_iDragItem, 0);
        if(m_iDragItem < ipos) ipos--;
        ipos = SendMessage(hwndCtl, LB_INSERTSTRING, ipos, (LPARAM)buf);
        SendMessage(hwndCtl, LB_SETITEMDATA, ipos, (LPARAM)pData);
        SendMessage(hwndCtl, LB_SETCURSEL, ipos, 0);
        m_bHasChanged = TRUE;
    }

    m_iDragItem = -1;
    DrawInsert(hWnd, hwndCtl, -1);
    return(TRUE);
}

INT_PTR CMainForm::AddImage(HWND hWnd)
{
    TCHAR namebuf[512];
    TCHAR filterbuf[128];
    namebuf[0] = 0;
    LoadString(m_hInstance, IDS_IMG_FILTER, filterbuf, 128);
    for(int i = 0; i < 128; i++)
        if(filterbuf[i] == 1) filterbuf[i] = 0;

    OPENFILENAME ofn;
    ofn.lStructSize = sizeof(OPENFILENAME);
    ofn.hwndOwner = hWnd;
    ofn.hInstance = m_hInstance;
    ofn.lpstrFilter = filterbuf;
    ofn.lpstrCustomFilter = NULL;
    ofn.nMaxCustFilter = 0;
    ofn.nFilterIndex = 0;
    ofn.lpstrFile = namebuf;
    ofn.nMaxFile = 512;
    ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.lpstrInitialDir = m_sLastOpenPath;
    ofn.lpstrTitle = NULL;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_NONETWORKBUTTON |
        OFN_EXPLORER | OFN_ENABLESIZING;
    ofn.nFileOffset = 0;
    ofn.nFileExtension = 0;
    ofn.lpstrDefExt = NULL;
    ofn.lCustData = 0;
    ofn.lpfnHook = NULL;
    ofn.lpTemplateName = NULL;
#ifndef _MSCOMP_
    ofn.pvReserved = NULL;
    ofn.dwReserved = 0;
    ofn.FlagsEx = 0;
#endif
    if(GetOpenFileName(&ofn))
    {
        if(IsIcoFile(namebuf))
        {
            if(ReadIconFile(hWnd, FALSE, namebuf)) m_bHasChanged = TRUE;
        }
        else
        {
            ADDIMGDATA aid;
            if(m_AddImgDlg->ShowModal(hWnd, namebuf, &aid))
            {
                InsertIconImages(hWnd, &aid);
                // insert images described by aid and/or (?) namebuf
            }
        }
    }
    return(TRUE);
}

INT_PTR CMainForm::NewFile(HWND hWnd)
{
    if(CheckChangedFile(hWnd))
    {
        ClearListBox(hWnd);
        m_bHasChanged = TRUE;
        m_sCurFileName[0] = 0;
        EnableMenuItems(hWnd);
    }
    return(TRUE);
}

BOOL CMainForm::IsIcoFile(LPTSTR sFileName)
{
    int slen = _tcslen(sFileName);
    if(slen < 4) return(FALSE);

    BOOL bRes = _tcsicmp(&sFileName[slen - 4], _T(".ico")) == 0;
    if(bRes)
    {
        FILE *fp = _tfopen(sFileName, _T("rb"));
        if(fp)
        {
            ICONDIR icd;

            bRes = fread(&icd, 1, 6, fp) > 5;
            bRes = bRes && icd.idReserved == 0 && icd.idType == 1;

            fclose(fp);
        }
        else bRes = FALSE;
    }
    return(bRes);
}

INT_PTR CMainForm::ImageDelete(HWND hWnd)
{
    HWND lb = GetDlgItem(hWnd, CTRL_ICONSLISTBOX);
    int i = SendMessage(lb, LB_GETCURSEL, 0, 0);
    if(i > -1)
    {
        LPICONDATA lpid = (LPICONDATA)SendMessage(lb, LB_GETITEMDATA, i, 0);
        free(lpid);
        SendMessage(lb, LB_DELETESTRING, i, 0);
        m_bHasChanged = TRUE;
        int n = SendMessage(lb, LB_GETCOUNT, 0, 0);
        if(n > 0)
        {
            if(i >= n) i = n - 1;
            SendMessage(lb, LB_SETCURSEL, i, 0);
        }
        ListBoxChange(hWnd, lb);
    }
    return(1);
}

INT_PTR CMainForm::ImageSave(HWND hWnd)
{
    HWND lb = GetDlgItem(hWnd, CTRL_ICONSLISTBOX);
    int i = SendMessage(lb, LB_GETCURSEL, 0, 0);
    if(i > -1)
    {
        LPICONDATA lpid = (LPICONDATA)SendMessage(lb, LB_GETITEMDATA, i, 0);
        BOOL bIsBmp = lpid->icoimg.icHeader.biSize == sizeof(BITMAPINFOHEADER);

        TCHAR namebuf[512];
        TCHAR filterbuf[128];
        namebuf[0] = 0;
        if(bIsBmp) LoadString(m_hInstance, IDS_BMP_FILTER, filterbuf, 128);
        else LoadString(m_hInstance, IDS_PNG_FILTER, filterbuf, 128);
#define IDS_BMP_FILTER 5
#define IDS_PNG_FILTER 6
        for(int i = 0; i < 128; i++)
            if(filterbuf[i] == 1) filterbuf[i] = 0;

        OPENFILENAME ofn;
        ofn.lStructSize = sizeof(OPENFILENAME);
        ofn.hwndOwner = hWnd;
        ofn.hInstance = m_hInstance;
        ofn.lpstrFilter = filterbuf;
        ofn.lpstrCustomFilter = NULL;
        ofn.nMaxCustFilter = 0;
        ofn.nFilterIndex = 0;
        ofn.lpstrFile = namebuf;
        ofn.nMaxFile = 512;
        ofn.lpstrFileTitle = NULL;
        ofn.nMaxFileTitle = 0;
        ofn.lpstrInitialDir = m_sLastOpenPath;
        ofn.lpstrTitle = NULL;
        ofn.Flags = OFN_OVERWRITEPROMPT | OFN_NONETWORKBUTTON | OFN_EXPLORER |
            OFN_ENABLESIZING;
        ofn.nFileOffset = 0;
        ofn.nFileExtension = 0;
        if(bIsBmp) ofn.lpstrDefExt = _T("bmp");
        else ofn.lpstrDefExt = _T("png");
        ofn.lCustData = 0;
        ofn.lpfnHook = NULL;
        ofn.lpTemplateName = NULL;
    #ifndef _MSCOMP_
        ofn.pvReserved = NULL;
        ofn.dwReserved = 0;
        ofn.FlagsEx = 0;
    #endif
        if(GetSaveFileName(&ofn))
        {
            FILE *fp = _tfopen(namebuf, _T("wb"));
            if(fp)
            {
                if(bIsBmp) SaveIconBitmap(lpid, fp);
                else fwrite(&lpid->icoimg, 1, lpid->icdi.dwBytesInRes, fp);
                fclose(fp);
            }
        }
    }
    return(1);
}

LPICONDATA CMainForm::IconDataFromAddImgData(PADDIMGDATA pAid, int iBitDepth,
    int iWidth, int iHeight)
{
    LPICONDATA icdRes = NULL;
    if(iBitDepth > 0)
    {
        switch(pAid->bImgType)
        {
        case 1:
            icdRes = PNGtoIconData((char*)pAid->pbRawData, pAid->lRawDataSize,
                iBitDepth, iWidth, iHeight, pAid->bDither, pAid->bOptimalPal);
            break;
        case 2:
            icdRes = BMPtoIconData((char*)pAid->pbRawData, pAid->lRawDataSize,
                iBitDepth, iWidth, iHeight, pAid->dwTranspColor, pAid->bDither,
                pAid->bOptimalPal);
            break;
        }
    }
    else icdRes = PNGtoIconCopy((char*)pAid->pbRawData, pAid->lRawDataSize);
    return(icdRes);
}

void CMainForm::InsertIconImages(HWND hWnd, PADDIMGDATA pAid)
{
    HWND lb = GetDlgItem(hWnd, CTRL_ICONSLISTBOX);
    LPICONDATA lpid = NULL;
    int idx, idepth, isize;
    int depths[] = {4, 8, 32};
    int sizes[] = {48, 32, 24, 16};
    TCHAR buf[32];

    for(int i = 0; i < 3; i++)
    {
        if(pAid->bColorDepth & (1 << i))
        {
            idepth = depths[i];
            for(int j = 0; j < 4; j++)
            {
                if(pAid->bSize & (1 << j))
                {
                    isize = sizes[j];
                    lpid = IconDataFromAddImgData(pAid, idepth, isize, isize);
                    if(lpid)
                    {
                        _stprintf(buf, _T("%dx%d - %d bpp"), isize, isize,
                            idepth);
                        idx = SendMessage(lb, LB_ADDSTRING, 0, (LPARAM)buf);
                        SendMessage(lb, LB_SETITEMDATA, (WPARAM)idx, (LPARAM)lpid);
                    }
                }
            }
            if(pAid->bSize & AID_SIZE_CUSTOM)
            {
                lpid = IconDataFromAddImgData(pAid, idepth, pAid->iCustWidth,
                    pAid->iCustHeight);
                if(lpid)
                {
                    _stprintf(buf, _T("%dx%d - %d bpp"), pAid->iCustWidth,
                        pAid->iCustHeight, idepth);
                    idx = SendMessage(lb, LB_ADDSTRING, 0, (LPARAM)buf);
                    SendMessage(lb, LB_SETITEMDATA, (WPARAM)idx, (LPARAM)lpid);
                }
            }
        }
    }

    if((pAid->bColorDepth & AID_DEPTH_PNG) && (pAid->bImgType == 1))
    {
        lpid = IconDataFromAddImgData(pAid, 0, 0, 0);
        if(lpid)
        {
            int width = 256;
            if(lpid->icdi.bWidth > 0) width = lpid->icdi.bWidth;
            int height = 256;
            if(lpid->icdi.bHeight > 0) height = lpid->icdi.bHeight;
            _stprintf(buf, _T("%dx%d - %d bpp"), width, height,
                lpid->icdi.wBitCount);
            idx = SendMessage(lb, LB_ADDSTRING, 0, (LPARAM)buf);
            SendMessage(lb, LB_SETITEMDATA, (WPARAM)idx, (LPARAM)lpid);
        }
    }
    return;
}

INT_PTR CMainForm::ImageEdit(HWND hWnd)
{
    HWND lb = GetDlgItem(hWnd, CTRL_ICONSLISTBOX);
    int i = SendMessage(lb, LB_GETCURSEL, 0, 0);
    if(i > -1)
    {
        LPICONDATA lpid = (LPICONDATA)SendMessage(lb, LB_GETITEMDATA, i, 0);
        if(m_EditImgDlg->ShowModal(hWnd, &lpid->icoimg, lpid->icdi.dwBytesInRes))
        {
            m_bHasChanged = TRUE;
            m_EditImgDlg->SaveImage();
            if(lpid->hPreview) DeleteObject(lpid->hPreview);
            lpid->hPreview = NULL;
            ListBoxChange(hWnd, lb);
        }
    }
    return(1);
}
