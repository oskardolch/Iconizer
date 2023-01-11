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

//#include <windows.h>
//#include <commctrl.h>
#include <tchar.h>
#include "MainForm.hpp"

#ifdef _MSCOMP_
#define ICC_STANDARD_CLASSES 0x00004000
#endif

INT_PTR CALLBACK MainDlgProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    CMainForm *mw = NULL;
    if(uMsg == WM_INITDIALOG) mw = (CMainForm*)lParam;
    else mw = (CMainForm*)GetWindowLongPtr(hwndDlg, DWLP_USER);

    switch(uMsg)
    {
    case WM_INITDIALOG:
        return(mw->WMInitDialog(hwndDlg, (HWND)wParam, lParam));
    case WM_COMMAND:
        return(mw->WMCommand(hwndDlg, HIWORD(wParam), LOWORD(wParam), (HWND)lParam));
    case WM_SYSCOMMAND:
        return(mw->WMSysCommand(hwndDlg, wParam, LOWORD(lParam), HIWORD(lParam)));
    case WM_DESTROY:
        return(mw->WMDestroy(hwndDlg));
    case WM_CLOSE:
        return(mw->WMClose(hwndDlg));
    default:
        if(mw)
            return(mw->WMUnknownMessage(hwndDlg, uMsg, wParam, lParam));
        return(0);
    }
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
    PSTR szCmdLine, int iCmdShow)
{
    MSG msg;
    BOOL res;

    CoInitialize(NULL);

    INITCOMMONCONTROLSEX ictr;
    ictr.dwSize = sizeof(INITCOMMONCONTROLSEX);
    ictr.dwICC = ICC_STANDARD_CLASSES | ICC_BAR_CLASSES;
    InitCommonControlsEx(&ictr);

    HWND wnd = NULL;
    HACCEL accel = NULL;
    CMainForm *mw = new CMainForm(hInstance, MainDlgProc, &wnd, &accel);

    BOOL Finish = false;

    while(!Finish)
    {
        res = GetMessage(&msg, 0, 0, 0);
        if(!TranslateAccelerator(wnd, accel, &msg) && !IsDialogMessage(wnd, &msg))
        {
            switch (res)
            {
            case -1:
                MessageBox(wnd, _T("Error?"), _T("Debug"), MB_OK);
            case 0:
                Finish = true;
            default:
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
        }
    }

    CoUninitialize();

    delete mw;
    ExitProcess(0);
    return (0);
}
