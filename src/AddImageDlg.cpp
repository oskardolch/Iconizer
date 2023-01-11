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

#include "AddImageDlg.hpp"
#include <tchar.h>
#include <stdio.h>
#include "../PNGUtils/PNGUtils.hpp"
#include "Iconizer.rh"
#include "Render.hpp"

INT_PTR CALLBACK AddImgDlgProc(HWND hwndDlg, UINT uMsg, WPARAM wParam,
    LPARAM lParam)
{
    CAddImageDlg *dlg = NULL;
    if(uMsg == WM_INITDIALOG) dlg = (CAddImageDlg*)lParam;
    else dlg = (CAddImageDlg*)GetWindowLongPtr(hwndDlg, DWLP_USER);

    switch(uMsg)
    {
    case WM_INITDIALOG:
        return(dlg->WMInitDialog(hwndDlg, (HWND)wParam, lParam));
    case WM_COMMAND:
        return(dlg->WMCommand(hwndDlg, HIWORD(wParam), LOWORD(wParam),
            (HWND)lParam));
    default:
        return(0);
    }
}

LONG_PTR CALLBACK ImageProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    WNDPROC oldProc = (WNDPROC)GetWindowLongPtr(hWnd, GWLP_USERDATA);
    CAddImageDlg *paid = (CAddImageDlg*)GetProp(hWnd, _T("AddImageDlg"));

    switch(uMsg)
    {
    case WM_SETCURSOR:
        return(paid->ImageWMSetCursor(hWnd, (HWND)wParam, LOWORD(lParam),
            HIWORD(lParam)));
    case WM_LBUTTONDOWN:
        return(paid->ImageWMLButtonDown(hWnd, wParam, LOWORD(lParam),
            HIWORD(lParam)));
    case WM_LBUTTONUP:
        return(paid->ImageWMLButtonUp(hWnd, wParam, LOWORD(lParam),
            HIWORD(lParam)));
    case WM_MOUSEMOVE:
        return(paid->ImageWMMouseMove(hWnd, wParam, LOWORD(lParam),
            HIWORD(lParam)));
    default:
        return(CallWindowProc(oldProc, hWnd, uMsg, wParam, lParam));
    }
}

CAddImageDlg::CAddImageDlg(HINSTANCE hInstance)
{
    m_hInstance = hInstance;
    m_bData = NULL;
    m_Bmp = NULL;
}

CAddImageDlg::~CAddImageDlg()
{
    if(m_Bmp) DeleteObject(m_Bmp);
    if(m_bData) free(m_bData);
}

BOOL CAddImageDlg::ShowModal(HWND hWnd, LPTSTR sFileName, PADDIMGDATA pAddImgData)
{
    m_sFileName = sFileName;
    m_AddImgData = pAddImgData;
    return(DialogBoxParam(m_hInstance, _T("ADDIMGDLG"), hWnd, AddImgDlgProc,
        (LPARAM)this) == IDOK);
}

INT_PTR CAddImageDlg::WMInitDialog(HWND hWnd, HWND hwndFocus, LPARAM lInitParam)
{
    m_bSelectCol = FALSE;
    m_ptImageStartClick.x = -1;
    m_ptImageStartClick.y = -1;
    m_AddImgData->dwTranspColor = 0xFF000000;

    SetWindowLongPtr(hWnd, DWLP_USER, lInitParam);

    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_EDTWIDTH, EM_SETLIMITTEXT, 8, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_EDTHEIGHT, EM_SETLIMITTEXT, 8, 0);
    EnableWindow(GetDlgItem(hWnd, ADDIMGDLG_CTLR_EDTWIDTH), FALSE);
    EnableWindow(GetDlgItem(hWnd, ADDIMGDLG_CTLR_EDTHEIGHT), FALSE);

    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHBDITHER, BM_SETCHECK,
        BST_CHECKED, 0);

    HWND hwndImg = GetDlgItem(hWnd, ADDIMGDLG_CTLR_IMAGE);
    SetProp(hwndImg, _T("AddImageDlg"), (HANDLE)this);
    SetWindowLongPtr(hwndImg, GWLP_USERDATA,
        SetWindowLongPtr(hwndImg, GWLP_WNDPROC, (LONG_PTR)ImageProc));

    LoadImage(hWnd, hwndImg);
    if(m_AddImgData->bImgType == 2) SetTransparentColor(hWnd, NULL);
    return(1);
}

INT_PTR CAddImageDlg::WMCommand(HWND hWnd, WORD wNotifyCode, WORD wID, HWND hwndCtl)
{
    switch(wID)
    {
    case IDOK:
        return(OKBtnClick(hWnd));
    case IDCANCEL:
        EndDialog(hWnd, 0);
        return(1);
    case ADDIMGDLG_CTLR_ACHBCUSTOM:
        return(CustomBtnClick(hWnd, hwndCtl));
    case ADDIMGDLG_CTLR_BTNVISTA:
        return(VistaBtnClick(hWnd));
    case ADDIMGDLG_CTLR_BTNXP:
        return(XPBtnClick(hWnd));
    case ADDIMGDLG_CTLR_BTNTRANSPSET:
        m_bSelectCol = TRUE;
        return(1);
    case ADDIMGDLG_CTLR_BTNTRANSPCLEAR:
        m_AddImgData->dwTranspColor = 0xFF000000;
        SetTransparentColor(hWnd, GetDlgItem(hWnd, ADDIMGDLG_CTLR_IMAGE));
        return(1);
    default:
        return(0);
    }
}

INT_PTR CAddImageDlg::CustomBtnClick(HWND hWnd, HWND hwndCtl)
{
    if(SendMessage(hwndCtl, BM_GETCHECK, 0, 0) == BST_CHECKED)
    {
        EnableWindow(GetDlgItem(hWnd, ADDIMGDLG_CTLR_EDTWIDTH), TRUE);
        EnableWindow(GetDlgItem(hWnd, ADDIMGDLG_CTLR_EDTHEIGHT), TRUE);
    }
    else
    {
        EnableWindow(GetDlgItem(hWnd, ADDIMGDLG_CTLR_EDTWIDTH), FALSE);
        EnableWindow(GetDlgItem(hWnd, ADDIMGDLG_CTLR_EDTHEIGHT), FALSE);
    }
    return(1);
}

INT_PTR CAddImageDlg::VistaBtnClick(HWND hWnd)
{
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB256, BM_SETCHECK, BST_CHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB48, BM_SETCHECK, BST_CHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB32, BM_SETCHECK, BST_CHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB24, BM_SETCHECK, BST_UNCHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB16, BM_SETCHECK, BST_CHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHBCUSTOM, BM_SETCHECK, BST_UNCHECKED, 0);
    EnableWindow(GetDlgItem(hWnd, ADDIMGDLG_CTLR_EDTWIDTH), FALSE);
    EnableWindow(GetDlgItem(hWnd, ADDIMGDLG_CTLR_EDTHEIGHT), FALSE);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHBPNG, BM_SETCHECK, BST_CHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHBRGBA, BM_SETCHECK, BST_CHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB8BIT, BM_SETCHECK, BST_CHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB4BIT, BM_SETCHECK, BST_CHECKED, 0);
    return(1);
}

INT_PTR CAddImageDlg::XPBtnClick(HWND hWnd)
{
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB256, BM_SETCHECK, BST_UNCHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB48, BM_SETCHECK, BST_CHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB32, BM_SETCHECK, BST_CHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB24, BM_SETCHECK, BST_UNCHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB16, BM_SETCHECK, BST_CHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHBCUSTOM, BM_SETCHECK, BST_UNCHECKED, 0);
    EnableWindow(GetDlgItem(hWnd, ADDIMGDLG_CTLR_EDTWIDTH), FALSE);
    EnableWindow(GetDlgItem(hWnd, ADDIMGDLG_CTLR_EDTHEIGHT), FALSE);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHBPNG, BM_SETCHECK, BST_UNCHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHBRGBA, BM_SETCHECK, BST_CHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB8BIT, BM_SETCHECK, BST_CHECKED, 0);
    SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB4BIT, BM_SETCHECK, BST_CHECKED, 0);
    return(1);
}

void CAddImageDlg::LoadImage(HWND hWnd, HWND hwndCtl)
{
    m_AddImgData->bImgType = 0;

    FILE *fp = _tfopen(m_sFileName, _T("rb"));
    if(fp)
    {
        if(m_Bmp) DeleteObject(m_Bmp);
        if(m_bData) free(m_bData);

        fseek(fp, 0, SEEK_END);
        m_AddImgData->lRawDataSize = ftell(fp);
        fseek(fp, 0, SEEK_SET);

        m_bData = (BYTE*)malloc(m_AddImgData->lRawDataSize);
        fread(m_bData, 1, m_AddImgData->lRawDataSize, fp);

        int iWidth = 0, iHeight = 0;

        if(m_bData[0] == 'B' && m_bData[1] == 'M')
        {
            // bitmap
            m_Bmp = RenderBitmap(hwndCtl, m_bData, m_AddImgData->dwTranspColor,
                &iWidth, &iHeight);
            m_AddImgData->bImgType = 2;
            ShowWindow(GetDlgItem(hWnd, ADDIMGDLG_CTLR_ACHBPNG), SW_HIDE);
        }
        else if(IsPNGFile((char*)m_bData, m_AddImgData->lRawDataSize))
        {
            // png
            m_Bmp = RenderPNG(hwndCtl, (LPICONIMAGE)m_bData,
                m_AddImgData->lRawDataSize, &iWidth, &iHeight);
            m_AddImgData->bImgType = 1;
            ShowWindow(GetDlgItem(hWnd, ADDIMGDLG_CTLR_GBTRANSP), SW_HIDE);
            ShowWindow(GetDlgItem(hWnd, ADDIMGDLG_CTLR_BTNTRANSPSET), SW_HIDE);
            ShowWindow(GetDlgItem(hWnd, ADDIMGDLG_CTLR_BTNTRANSPCLEAR), SW_HIDE);
            ShowWindow(GetDlgItem(hWnd, ADDIMGDLG_CTLR_TRANSPCOL), SW_HIDE);
        }
        fclose(fp);

        TCHAR buf[32];
        _stprintf(buf, _T("%d"), iWidth);
        SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_EDTWIDTH, WM_SETTEXT, 0,
            (LPARAM)buf);
        _stprintf(buf, _T("%d"), iHeight);
        SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_EDTHEIGHT, WM_SETTEXT, 0,
            (LPARAM)buf);
    }
    return;
}

LONG_PTR CAddImageDlg::ImageWMSetCursor(HWND hWnd, HWND hwnd, WORD nHittest,
    WORD wMouseMsg)
{
    if(m_bSelectCol) SetCursor(LoadCursor(0, IDC_CROSS));
    return(1);
}

LONG_PTR CAddImageDlg::ImageWMLButtonDown(HWND hWnd, WPARAM fwKeys, WORD xPos,
    WORD yPos)
{
    if(m_bSelectCol && fwKeys == MK_LBUTTON)
    {
        m_bHasMoved = FALSE;
        m_ptImageStartClick.x = xPos;
        m_ptImageStartClick.y = yPos;
        m_ptImageMove.x = xPos;
        m_ptImageMove.y = yPos;
        SetCapture(hWnd);
        m_ImageDC = GetDC(hWnd);
        m_iImageROP2 = SetROP2(m_ImageDC, R2_NOTXORPEN);
        return(0);
    }
    return(1);
}

LONG_PTR CAddImageDlg::ImageWMLButtonUp(HWND hWnd, WPARAM fwKeys, WORD xPos,
    WORD yPos)
{
    if(m_ptImageStartClick.x > -1)
    {
        Rectangle(m_ImageDC, m_ptImageStartClick.x, m_ptImageStartClick.y,
            m_ptImageMove.x, m_ptImageMove.y);
        SetROP2(m_ImageDC, m_iImageROP2);

        BOOL bDoSet = !m_bHasMoved;
        if(bDoSet) m_AddImgData->dwTranspColor = GetPixel(m_ImageDC, xPos, yPos);
        ReleaseDC(hWnd, m_ImageDC);
        ReleaseCapture();
        m_ptImageStartClick.x = -1;
        m_bSelectCol = FALSE;
        SetCursor(LoadCursor(0, IDC_ARROW));

        if(bDoSet) SetTransparentColor(GetParent(hWnd), hWnd);
        return(0);
    }
    return(1);
}

LONG_PTR CAddImageDlg::ImageWMMouseMove(HWND hWnd, WPARAM fwKeys, WORD xPos,
    WORD yPos)
{
    signed short dx = (signed short)xPos;
    signed short dy = (signed short)yPos;
    if(m_ptImageStartClick.x > -1)
    {
        Rectangle(m_ImageDC, m_ptImageStartClick.x, m_ptImageStartClick.y,
            m_ptImageMove.x, m_ptImageMove.y);
        m_bHasMoved |= abs(m_ptImageMove.x - dx) > 2 ||
            abs(m_ptImageMove.y - dy) > 2;
        m_ptImageMove.x = dx;
        m_ptImageMove.y = dy;
        Rectangle(m_ImageDC, m_ptImageStartClick.x, m_ptImageStartClick.y,
            m_ptImageMove.x, m_ptImageMove.y);
        return(0);
    }
    return(1);
}

void CAddImageDlg::SetTransparentColor(HWND hWnd, HWND hwndImg)
{
    const int bmpwidth = 24;

    HWND wnd = GetDlgItem(hWnd, ADDIMGDLG_CTLR_TRANSPCOL);
    HDC dc = GetDC(wnd);
    HDC ldc = CreateCompatibleDC(dc);
    HBITMAP bmp = CreateCompatibleBitmap(dc, bmpwidth, bmpwidth);
    SelectObject(ldc, bmp);

    DWORD bgr = m_AddImgData->dwTranspColor;
    if(m_AddImgData->dwTranspColor == 0xFF000000) bgr = RGB(224, 224, 224);

    HBRUSH hbr = CreateSolidBrush(bgr);
    HBRUSH hbrprev = (HBRUSH)SelectObject(ldc, hbr);
    Rectangle(ldc, 0, 0, bmpwidth, bmpwidth);

    if(m_AddImgData->dwTranspColor == 0xFF000000)
    {
        HPEN pnnul = CreatePen(PS_NULL, 0, 0);
        HPEN prevpen = (HPEN)SelectObject(ldc, pnnul);
        DeleteObject(hbr);
        hbr = CreateSolidBrush(RGB(192, 192, 192));
        SelectObject(ldc, hbr);

        HRGN rgn = CreateRectRgn(1, 1, bmpwidth - 1, bmpwidth - 1);
        SelectClipRgn(ldc, rgn);

        int dx, dy;
        for(int j = 0; j < bmpwidth / 4; j++)
        {
            dy = 4*j;
            for(int i = 0; i < bmpwidth / 8; i++)
            {
                dx = 4*(2*i + (j % 2));
                Rectangle(ldc, dx, dy, dx + 5, dy + 5);
            }
        }

        SelectClipRgn(ldc, NULL);
        DeleteObject(rgn);
        DeleteObject(SelectObject(ldc, prevpen));
    }

    DeleteObject(SelectObject(ldc, hbrprev));
    DeleteDC(ldc);
    ReleaseDC(wnd, dc);

    SendMessage(wnd, STM_SETIMAGE, IMAGE_BITMAP, (LPARAM)bmp);

    if(hwndImg)
    {
        if(m_Bmp) DeleteObject(m_Bmp);
        m_Bmp = RenderBitmap(hwndImg, m_bData, m_AddImgData->dwTranspColor,
            NULL, NULL);
    }
    return;
}

INT_PTR CAddImageDlg::OKBtnClick(HWND hWnd)
{
    m_AddImgData->bSize = 0;
    m_AddImgData->bColorDepth = 0;
    if(SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB256, BM_GETCHECK, 0, 0) ==
        BST_CHECKED) m_AddImgData->bSize |= AID_SIZE_256;
    if(SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB48, BM_GETCHECK, 0, 0) ==
        BST_CHECKED) m_AddImgData->bSize |= AID_SIZE_48;
    if(SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB32, BM_GETCHECK, 0, 0) ==
        BST_CHECKED) m_AddImgData->bSize |= AID_SIZE_32;
    if(SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB24, BM_GETCHECK, 0, 0) ==
        BST_CHECKED) m_AddImgData->bSize |= AID_SIZE_24;
    if(SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB16, BM_GETCHECK, 0, 0) ==
        BST_CHECKED) m_AddImgData->bSize |= AID_SIZE_16;
    if(SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHBCUSTOM, BM_GETCHECK, 0, 0) ==
        BST_CHECKED)
    {
        m_AddImgData->bSize |= AID_SIZE_CUSTOM;
        HWND wnd = GetDlgItem(hWnd, ADDIMGDLG_CTLR_EDTWIDTH);
        TCHAR buf[16];
        TCHAR errmsg[128];
        SendMessage(wnd, WM_GETTEXT, 8, (LPARAM)buf);
        if(_stscanf(buf, _T("%d"), &m_AddImgData->iCustWidth) != 1 ||
            m_AddImgData->iCustWidth < 0 || m_AddImgData->iCustWidth > 256)
        {
            LoadString(m_hInstance, IDS_DLG_ERR, buf, 16);
            LoadString(m_hInstance, IDS_MSG_INVALIDUSERINPUT, errmsg, 128);
            MessageBox(hWnd, errmsg, buf, MB_OK | MB_ICONERROR);
            SendMessage(wnd, EM_SETSEL, 0, (LPARAM)-1);
            SetFocus(wnd);
            return(1);
        }
        wnd = GetDlgItem(hWnd, ADDIMGDLG_CTLR_EDTHEIGHT);
        SendMessage(wnd, WM_GETTEXT, 8, (LPARAM)buf);
        if(_stscanf(buf, _T("%d"), &m_AddImgData->iCustHeight) != 1 ||
            m_AddImgData->iCustHeight < 0 || m_AddImgData->iCustHeight > 256)
        {
            LoadString(m_hInstance, IDS_DLG_ERR, buf, 16);
            LoadString(m_hInstance, IDS_MSG_INVALIDUSERINPUT, errmsg, 128);
            MessageBox(hWnd, errmsg, buf, MB_OK | MB_ICONERROR);
            SendMessage(wnd, EM_SETSEL, 0, (LPARAM)-1);
            SetFocus(wnd);
            return(1);
        }
    }

    if(SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHBPNG, BM_GETCHECK, 0, 0) ==
        BST_CHECKED) m_AddImgData->bColorDepth |= AID_DEPTH_PNG;
    if(SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHBRGBA, BM_GETCHECK, 0, 0) ==
        BST_CHECKED) m_AddImgData->bColorDepth |= AID_DEPTH_RGBA;
    if(SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB8BIT, BM_GETCHECK, 0, 0) ==
        BST_CHECKED) m_AddImgData->bColorDepth |= AID_DEPTH_8;
    if(SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHB4BIT, BM_GETCHECK, 0, 0) ==
        BST_CHECKED) m_AddImgData->bColorDepth |= AID_DEPTH_4;

    m_AddImgData->bDither = SendDlgItemMessage(hWnd, ADDIMGDLG_CTLR_ACHBDITHER,
        BM_GETCHECK, 0, 0) == BST_CHECKED;
    m_AddImgData->bOptimalPal = SendDlgItemMessage(hWnd,
        ADDIMGDLG_CTLR_ACHBOPTIMAL, BM_GETCHECK, 0, 0) == BST_CHECKED;

    m_AddImgData->pbRawData = m_bData;

    EndDialog(hWnd, 1);
    return(1);
}
