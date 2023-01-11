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

#include <tchar.h>
#include "Editor.hpp"
#include "Render.hpp"
#include "Iconizer.rh"

const int giPalSize = 12;
const int xpDx = 4;
const int giDragThreshold = 2;

INT_PTR CALLBACK EditImgDlgProc(HWND hwndDlg, UINT uMsg, WPARAM wParam,
    LPARAM lParam)
{
    CEditImageDlg *dlg = NULL;
    if(uMsg == WM_INITDIALOG) dlg = (CEditImageDlg*)lParam;
    else dlg = (CEditImageDlg*)GetWindowLongPtr(hwndDlg, DWLP_USER);

    switch(uMsg)
    {
    case WM_INITDIALOG:
        return(dlg->WMInitDialog(hwndDlg, (HWND)wParam, lParam));
    case WM_COMMAND:
        return(dlg->WMCommand(hwndDlg, HIWORD(wParam), LOWORD(wParam),
            (HWND)lParam));
    case WM_NOTIFY:
        return(dlg->WMNotify(hwndDlg, (int)wParam, (LPNMHDR)lParam));
    case WM_PAINT:
        return(dlg->WMPaint(hwndDlg, (HDC)wParam));
    case WM_LBUTTONDOWN:
        return(dlg->WMLButtonDown(hwndDlg, wParam, LOWORD(lParam), HIWORD(lParam)));
    case WM_LBUTTONUP:
        return(dlg->WMLButtonUp(hwndDlg, wParam, LOWORD(lParam), HIWORD(lParam)));
    case WM_MOUSEMOVE:
        return(dlg->WMMouseMove(hwndDlg, wParam, LOWORD(lParam), HIWORD(lParam)));
    default:
        return(0);
    }
}

CEditImageDlg::CEditImageDlg(HINSTANCE hInstance)
{
    m_hInstance = hInstance;
    m_IcoImg = NULL;
    m_pImage = NULL;

    HDC hdc = GetDC(0);
    int bps = GetDeviceCaps(hdc, BITSPIXEL);
    ReleaseDC(0, hdc);

    UINT ilflgs = ILC_COLOR | ILC_MASK;
    if(bps > 4) ilflgs = ILC_COLOR8 | ILC_MASK;
    if(bps > 8) ilflgs = ILC_COLOR16 | ILC_MASK;
    if(bps > 16) ilflgs = ILC_COLOR24 | ILC_MASK;
    if(bps > 24) ilflgs = ILC_COLOR32 | ILC_MASK;

    const int icowidth = 24;
    const int icoheight = 22;
    m_hImgs = ImageList_Create(icowidth, icoheight, ilflgs, 2, 2);

    HICON hIco = (HICON)LoadImage(m_hInstance, _T("ZOOMINICO"), IMAGE_ICON,
        icowidth, icoheight, LR_DEFAULTCOLOR);
    ImageList_ReplaceIcon(m_hImgs, -1, hIco);
    hIco = (HICON)LoadImage(m_hInstance, _T("ZOOMOUTICO"), IMAGE_ICON,
        icowidth, icoheight, LR_DEFAULTCOLOR);
    ImageList_ReplaceIcon(m_hImgs, -1, hIco);
}

CEditImageDlg::~CEditImageDlg()
{
    if(m_pImage) free(m_pImage);
    ImageList_Destroy(m_hImgs);
}

BOOL CEditImageDlg::ShowModal(HWND hWnd, LPICONIMAGE pIcoImg, int iSize)
{
    m_IcoImg = pIcoImg;
    if(m_pImage) free(m_pImage);
    m_pImage = (LPBITMAPINFO)malloc(sizeof(BITMAPINFOHEADER) +
        2*m_IcoImg->icHeader.biWidth*m_IcoImg->icHeader.biHeight);

    m_cTransp = 0x00C0C040;
    GetEditImage(m_IcoImg, m_pImage, m_cTransp);
    m_iSelColor = -2;
    m_ptDown.x = -1;
    m_bMouseDrag = FALSE;
    m_bHasChanged = FALSE;
    return(DialogBoxParam(m_hInstance, _T("EDITIMGDLG"), hWnd, EditImgDlgProc,
        (LPARAM)this) == IDOK && m_bHasChanged);
}

INT_PTR CEditImageDlg::WMInitDialog(HWND hWnd, HWND hwndFocus, LPARAM lInitParam)
{
    SetWindowLongPtr(hWnd, DWLP_USER, lInitParam);
    HWND tlb = GetDlgItem(hWnd, EDTIMGDLG_CTLR_TOOLBAR);
    SendMessage(tlb, TB_SETIMAGELIST, 0, (LPARAM)m_hImgs);

    const int numbtns = 2;
    TBBUTTON btns[numbtns];
    int idx = 0;
    /*btns[idx].iBitmap = -1;
    btns[idx].idCommand = EDTIMGDLG_CTLR_SELPIXELBNT;
    btns[idx].fsState = TBSTATE_ENABLED | TBSTATE_CHECKED;
    btns[idx].fsStyle = TBSTYLE_BUTTON | TBSTYLE_CHECKGROUP;
    btns[idx].dwData = 0;
    btns[idx++].iString = -1;

    btns[idx].iBitmap = -1;
    btns[idx].idCommand = EDTIMGDLG_CTLR_SELCOLORBNT;
    btns[idx].fsState = TBSTATE_ENABLED;
    btns[idx].fsStyle = TBSTYLE_BUTTON | TBSTYLE_CHECKGROUP;
    btns[idx].dwData = 0;
    btns[idx++].iString = -1;

    btns[idx].iBitmap = -1;
    btns[idx].idCommand = -1;
    btns[idx].fsState = 0;
    btns[idx].fsStyle = TBSTYLE_SEP;
    btns[idx].dwData = 0;
    btns[idx++].iString = -1;*/

    btns[idx].iBitmap = 0;
    btns[idx].idCommand = EDTIMGDLG_CTLR_ZOOMINBNT;
    btns[idx].fsState = TBSTATE_ENABLED;
    btns[idx].fsStyle = TBSTYLE_BUTTON;
    btns[idx].dwData = 0;
    btns[idx++].iString = -1;

    btns[idx].iBitmap = 1;
    btns[idx].idCommand = EDTIMGDLG_CTLR_ZOOMOUTBNT;
    btns[idx].fsState = TBSTATE_ENABLED;
    btns[idx].fsStyle = TBSTYLE_BUTTON;
    btns[idx].dwData = 0;
    btns[idx++].iString = -1;
    SendMessage(tlb, TB_ADDBUTTONS, numbtns, (LPARAM)btns);
    SendMessage(tlb, TB_AUTOSIZE, 0, 0);

    SetupPalette(hWnd);
    return(1);
}

INT_PTR CEditImageDlg::WMCommand(HWND hWnd, WORD wNotifyCode, WORD wID, HWND hwndCtl)
{
    switch(wID)
    {
    case IDOK:
        EndDialog(hWnd, 1);
        return(1);
    case IDCANCEL:
        EndDialog(hWnd, 0);
        return(1);
    case EDTIMGDLG_CTLR_ZOOMINBNT:
        return(ZoomIn(hWnd, hwndCtl));
    case EDTIMGDLG_CTLR_ZOOMOUTBNT:
        return(ZoomOut(hWnd, hwndCtl));
    default:
        return(0);
    }
}

INT_PTR CEditImageDlg::WMNotify(HWND hWnd, int idCtrl, LPNMHDR pnmh)
{
    //if(idCtrl == EDTIMGDLG_CTLR_TOOLBAR)
    {
        switch(pnmh->code)
        {
        case TTN_NEEDTEXT:
            return(NFToolBarNeedText(hWnd, (LPTOOLTIPTEXT)pnmh));
        default:
            return(0);
        }
    }
    return(0);
}

INT_PTR CEditImageDlg::WMPaint(HWND hWnd, HDC hdc)
{
    RECT R;
    if(GetUpdateRect(hWnd, &R, FALSE))
    {
        PAINTSTRUCT ps;
        HDC ldc = BeginPaint(hWnd, &ps);
        DrawPalette(hWnd, ldc);
        DrawImage(hWnd, ldc);
        //DefWindowProc(hWnd, WM_PAINT, (WPARAM)ldc, 0);
        EndPaint(hWnd, &ps);
    }
    return(0);
}

INT_PTR CEditImageDlg::WMLButtonDown(HWND hWnd, WPARAM fwKeys, WORD xPos, WORD yPos)
{
    m_ptDown.x = xPos;
    m_ptDown.y = yPos;
    m_bMouseDrag = FALSE;
    return(0);
}

INT_PTR CEditImageDlg::WMLButtonUp(HWND hWnd, WPARAM fwKeys, WORD xPos, WORD yPos)
{
    if(m_bMouseDrag)
    {
    }
    else LButtonClick(hWnd);
    m_ptDown.x = -1;
    m_bMouseDrag = FALSE;
    return(0);
}

INT_PTR CEditImageDlg::WMMouseMove(HWND hWnd, WPARAM fwKeys, WORD xPos, WORD yPos)
{
    if(m_bMouseDrag)
    {
    }
    else if(m_ptDown.x > -1)
    {
        m_bMouseDrag = abs(xPos - m_ptDown.x) > giDragThreshold |
            abs(yPos - m_ptDown.y) > giDragThreshold;
    }
    else
    {
    }
    return(0);
}

INT_PTR CEditImageDlg::NFToolBarNeedText(HWND hWnd, LPTOOLTIPTEXT pttt)
{
    pttt->hinst = m_hInstance;
    switch(pttt->hdr.idFrom)
    {
    case EDTIMGDLG_CTLR_SELPIXELBNT:
        pttt->lpszText = MAKEINTRESOURCE(IDS_SEL_PIXEL);
        break;
    case EDTIMGDLG_CTLR_SELCOLORBNT:
        pttt->lpszText = MAKEINTRESOURCE(IDS_SEL_COLOR);
        break;
    case EDTIMGDLG_CTLR_ZOOMINBNT:
        pttt->lpszText = MAKEINTRESOURCE(IDS_ZOOM_IN);
        break;
    case EDTIMGDLG_CTLR_ZOOMOUTBNT:
        pttt->lpszText = MAKEINTRESOURCE(IDS_ZOOM_OUT);
        break;
    }
    return(0);
}

void CEditImageDlg::SetupPalette(HWND hWnd)
{
    BITMAPINFO *pbi = (BITMAPINFO*)m_IcoImg;
    int columns = 8;
    if(pbi->bmiHeader.biBitCount > 4) columns = 16;

    RECT R;
    GetClientRect(GetDlgItem(hWnd, EDTIMGDLG_CTLR_TOOLBAR), &R);

    m_ptPalOrig.x = 4 + xpDx;
    m_ptPalOrig.y = R.bottom + 16 + xpDx;

    R.top = R.bottom + xpDx;
    R.left = xpDx;
    R.right = columns*giPalSize + 9;
    m_iPalCount = pbi->bmiHeader.biClrUsed;
    if(m_iPalCount < 1)
    {
        m_iPalCount = 1;
        for(int i = 0; i < pbi->bmiHeader.biBitCount; i++) m_iPalCount *= 2;
    }
    R.bottom = 21 + giPalSize + giPalSize*m_iPalCount/ columns;
    SetWindowPos(GetDlgItem(hWnd, EDTIMGDLG_CTLR_GBPALETTE), NULL, R.left,
        R.top, R.right, R.bottom, SWP_NOZORDER);
    SetWindowPos(GetDlgItem(hWnd, IDOK), NULL, 4, R.top + R.bottom + 4, 0, 0,
        SWP_NOSIZE);
    if(columns > 8)
        SetWindowPos(GetDlgItem(hWnd, IDCANCEL), NULL, 84,
            R.top + R.bottom + 4, 0, 0, SWP_NOSIZE);
    else
        SetWindowPos(GetDlgItem(hWnd, IDCANCEL), NULL, 4,
            R.top + R.bottom + 32, 0, 0, SWP_NOSIZE);

    m_iZoom = 4;
    R.right += 2*xpDx;

    m_ptImgOrig.x = R.right + 4;
    m_ptImgOrig.y = R.top + 16;

    SetWindowPos(GetDlgItem(hWnd, EDTIMGDLG_CTLR_GBIMAGE), NULL, R.right,
        R.top, 0, 0, SWP_NOSIZE);
    SetupImage(hWnd);
    return;
}

void CEditImageDlg::SetupImage(HWND hWnd)
{
    BITMAPINFO *pbi = (BITMAPINFO*)m_IcoImg;
    int dx = m_iZoom*pbi->bmiHeader.biWidth;
    int dy = m_iZoom*pbi->bmiHeader.biHeight/2;
    if(dx < 36) dx = 36;
    SetWindowPos(GetDlgItem(hWnd, EDTIMGDLG_CTLR_GBIMAGE), NULL, 0,
        0, dx + 9, dy + 21, SWP_NOMOVE);

    RECT R;
    GetWindowRect(GetDlgItem(hWnd, IDCANCEL), &R);
    MapWindowPoints(0, hWnd, (LPPOINT)&R, 2);
    int minh = R.bottom;
    GetWindowRect(GetDlgItem(hWnd, EDTIMGDLG_CTLR_GBIMAGE), &R);
    MapWindowPoints(0, hWnd, (LPPOINT)&R, 2);
    if(minh > R.bottom) R.bottom = minh;
    R.left = 0;
    R.top = 0;
    R.right += xpDx;
    R.bottom += xpDx;
    DWORD st = (DWORD)GetWindowLongPtr(hWnd, GWL_STYLE);
    DWORD ex = (DWORD)GetWindowLongPtr(hWnd, GWL_EXSTYLE);
    AdjustWindowRectEx(&R, st, FALSE, ex);
    SetWindowPos(hWnd, NULL, 0, 0, R.right - R.left, R.bottom - R.top, SWP_NOMOVE);
    return;
}

void CEditImageDlg::DrawPalette(HWND hWnd, HDC hdc)
{
    BITMAPINFO *pbi = (BITMAPINFO*)m_IcoImg;
    int columns = 8;
    if(pbi->bmiHeader.biBitCount > 4) columns = 16;

    HPEN hpn = CreatePen(PS_SOLID, 0, 0x000000FF);
    HPEN hpnhigh = CreatePen(PS_SOLID, 2, 0x00000000);
    HPEN prevpn = (HPEN)SelectObject(hdc, hpn);

    HBRUSH hbr = CreateSolidBrush(m_cTransp);
    HBRUSH prevbr = (HBRUSH)SelectObject(hdc, hbr);
    RGBQUAD *prgb = pbi->bmiColors;
    PALETTEENTRY pal[16];
    int dx = m_ptPalOrig.x;
    int dy = m_ptPalOrig.y;

    if(m_iSelColor == -1) SelectObject(hdc, hpnhigh);
    Rectangle(hdc, dx, dy, dx + giPalSize, dy + giPalSize);
    if(m_iSelColor == -1) SelectObject(hdc, hpn);

    hpn = CreatePen(PS_SOLID, 0, 0x00C0C0C0);
    DeleteObject(SelectObject(hdc, hpn));

    for(int i = 0; i < m_iPalCount; i++)
    {
        if(i % columns < 1)
        {
            dx = m_ptPalOrig.x;
            dy += giPalSize;
        }
        else dx += giPalSize;
        if(m_iSelColor == i) SelectObject(hdc, hpnhigh);
        hbr = CreateSolidBrush(RGB(prgb->rgbRed, prgb->rgbGreen, prgb->rgbBlue));
        DeleteObject(SelectObject(hdc, hbr));
        Rectangle(hdc, dx, dy, dx + giPalSize, dy + giPalSize);
        if(m_iSelColor == i) SelectObject(hdc, hpn);
        prgb++;
    }
    DeleteObject(hpnhigh);
    DeleteObject(SelectObject(hdc, prevbr));
    DeleteObject(SelectObject(hdc, prevpn));
    return;
}

void CEditImageDlg::DrawImage(HWND hWnd, HDC hdc)
{
    int srcw = m_pImage->bmiHeader.biWidth;
    int srch = m_pImage->bmiHeader.biHeight;
    StretchDIBits(hdc, m_ptImgOrig.x, m_ptImgOrig.y, m_iZoom*srcw, m_iZoom*srch,
        0, 0, srcw, srch, m_pImage->bmiColors, m_pImage, DIB_RGB_COLORS,
        SRCCOPY);
    return;
}

INT_PTR CEditImageDlg::ZoomIn(HWND hWnd, HWND hwndCtl)
{
    if(m_iZoom < 2) m_iZoom = 2;
    else m_iZoom += 2;
    HWND tb = GetDlgItem(hWnd, EDTIMGDLG_CTLR_TOOLBAR);
    SendMessage(tb, TB_SETSTATE, EDTIMGDLG_CTLR_ZOOMOUTBNT,
        MAKELONG(TBSTATE_ENABLED, 0));
    if(m_iZoom > 6)
        SendMessage(tb, TB_SETSTATE, EDTIMGDLG_CTLR_ZOOMINBNT, MAKELONG(0, 0));
    SetupImage(hWnd);
    return(1);
}

INT_PTR CEditImageDlg::ZoomOut(HWND hWnd, HWND hwndCtl)
{
    m_iZoom -= 2;
    HWND tb = GetDlgItem(hWnd, EDTIMGDLG_CTLR_TOOLBAR);
    SendMessage(tb, TB_SETSTATE, EDTIMGDLG_CTLR_ZOOMINBNT,
        MAKELONG(TBSTATE_ENABLED, 0));
    if(m_iZoom < 2)
    {
        SendMessage(tb, TB_SETSTATE, EDTIMGDLG_CTLR_ZOOMOUTBNT, MAKELONG(0, 0));
        m_iZoom = 1;
    }
    SetupImage(hWnd);
    return(1);
}

void CEditImageDlg::LButtonClick(HWND hWnd)
{
    /*TBBUTTONINFO bi;
    bi.cbSize = sizeof(TBBUTTONINFO);
    bi.dwMask = TBIF_STATE;
    if(SendDlgItemMessage(hWnd, EDTIMGDLG_CTLR_TOOLBAR, TB_GETBUTTONINFO,
        EDTIMGDLG_CTLR_SELPIXELBNT, (LPARAM)&bi) < 0) return;

    if(bi.fsState & TBSTATE_CHECKED) SelImageClick(hWnd);
    else*/ SelPaletteClick(hWnd);
    return;
}

void CEditImageDlg::SelImageClick(HWND hWnd)
{
    int iPal = IsPalettePoint(hWnd, &m_ptDown);
    int iImg = IsImagePoint(hWnd, &m_ptDown);
    return;
}

void CEditImageDlg::SelPaletteClick(HWND hWnd)
{
    RECT R;
    int iPal = IsPalettePoint(hWnd, &m_ptDown);
    if(iPal > -2)
    {
        m_iSelColor = iPal;
        GetWindowRect(GetDlgItem(hWnd, EDTIMGDLG_CTLR_GBPALETTE), &R);
        MapWindowPoints(0, hWnd, (LPPOINT)&R, 2);
        InvalidateRect(hWnd, &R, TRUE);
        return;
    }
    int iImg = IsImagePoint(hWnd, &m_ptDown);
    if(iImg > -1 && m_iSelColor > -2)
    {
        m_bHasChanged = TRUE;
        BYTE *pb = (BYTE*)m_pImage->bmiColors;
        if(m_iSelColor < 0)
        {
            pb[4*iImg] = (m_cTransp >> 16) & 0xFF;
            pb[4*iImg + 1] = (m_cTransp >> 8) & 0xFF;
            pb[4*iImg + 2] = m_cTransp & 0xFF;
            pb[4*iImg + 3] |= 1;
        }
        else
        {
            RGBQUAD *pcol = &m_IcoImg->icColors[m_iSelColor];
            pb[4*iImg] = pcol->rgbBlue;
            pb[4*iImg + 1] = pcol->rgbGreen;
            pb[4*iImg + 2] = pcol->rgbRed;
            pb[4*iImg + 3] &= ~1;
        }
        R.left = m_ptDown.x - m_iZoom - 2;
        R.right = m_ptDown.x + m_iZoom + 2;
        R.top = m_ptDown.y - m_iZoom - 2;
        R.bottom = m_ptDown.y + m_iZoom + 2;
        InvalidateRect(hWnd, &R, FALSE);
    }
    return;
}

int CEditImageDlg::IsPalettePoint(HWND hWnd, LPPOINT ppt)
{
    int row = (ppt->y - m_ptPalOrig.y);
    int col = (ppt->x - m_ptPalOrig.x);
    if(row < 0 || col < 0) return(-2);
    row /= giPalSize;
    col /= giPalSize;
    if(row == 0)
    {
        if(col == 0) return(-1);
        else return(-2);
    }
    int columns = 8;
    if(m_iPalCount > 16) columns = 16;
    int rows = m_iPalCount/columns + 1;
    if(row >= rows || col >= columns) return(-2);
    return((row - 1)*columns + col);
}

int CEditImageDlg::IsImagePoint(HWND hWnd, LPPOINT ppt)
{
    int row = (ppt->y - m_ptImgOrig.y);
    int col = (ppt->x - m_ptImgOrig.x);
    if(row < 0 || col < 0) return(-1);
    row /= m_iZoom;
    col /= m_iZoom;
    if(row >= m_pImage->bmiHeader.biWidth ||
        col >= m_pImage->bmiHeader.biHeight) return(-1);
    return((m_pImage->bmiHeader.biHeight - row - 1)*m_pImage->bmiHeader.biWidth + col);
}

void CEditImageDlg::SaveImage()
{
    SaveEditImage(m_IcoImg, m_pImage);
    return;
}
