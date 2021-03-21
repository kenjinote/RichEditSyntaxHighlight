#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <windows.h>
#include <Richedit.h>

#define DEFAULT_DPI 96
#define SCALEX(X) MulDiv(X, uDpiX, DEFAULT_DPI)
#define SCALEY(Y) MulDiv(Y, uDpiY, DEFAULT_DPI)
#define POINT2PIXEL(PT) MulDiv(PT, uDpiY, 72)

TCHAR szClassName[] = TEXT("Window");
static HFONT hFont;

BOOL GetScaling(HWND hWnd, UINT* pnX, UINT* pnY)
{
	BOOL bSetScaling = FALSE;
	const HMONITOR hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST);
	if (hMonitor)
	{
		HMODULE hShcore = LoadLibrary(TEXT("SHCORE"));
		if (hShcore)
		{
			typedef HRESULT __stdcall GetDpiForMonitor(HMONITOR, int, UINT*, UINT*);
			GetDpiForMonitor* fnGetDpiForMonitor = reinterpret_cast<GetDpiForMonitor*>(GetProcAddress(hShcore, "GetDpiForMonitor"));
			if (fnGetDpiForMonitor)
			{
				UINT uDpiX, uDpiY;
				if (SUCCEEDED(fnGetDpiForMonitor(hMonitor, 0, &uDpiX, &uDpiY)) && uDpiX > 0 && uDpiY > 0)
				{
					*pnX = uDpiX;
					*pnY = uDpiY;
					bSetScaling = TRUE;
				}
			}
			FreeLibrary(hShcore);
		}
	}
	if (!bSetScaling)
	{
		HDC hdc = GetDC(NULL);
		if (hdc)
		{
			*pnX = GetDeviceCaps(hdc, LOGPIXELSX);
			*pnY = GetDeviceCaps(hdc, LOGPIXELSY);
			ReleaseDC(NULL, hdc);
			bSetScaling = TRUE;
		}
	}
	if (!bSetScaling)
	{
		*pnX = DEFAULT_DPI;
		*pnY = DEFAULT_DPI;
		bSetScaling = TRUE;
	}
	return bSetScaling;
}

BOOL IsHighlightText(LPCWSTR lpszText)
{
	LPCWSTR lpszHighlightText[] = {
		L"あいうえお",
		L"hello",
		L"aaa",
		L"abc",
		L"book",
		L"apple"
	};
	for (auto i : lpszHighlightText) {
		if (lstrcmpi(lpszText, i) == 0) {
			return TRUE;
		}
	}
	return FALSE;
}

WNDPROC DefaultRichEditProc;
LRESULT CALLBACK RichEditProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
	case WM_MOUSEWHEEL:
		if (GET_KEYSTATE_WPARAM(wParam) & MK_CONTROL) {
			return 0;
		}
		break;
	case WM_PAINT:
	{
		HideCaret(hWnd);
		CallWindowProc(DefaultRichEditProc, hWnd, msg, wParam, lParam);
		RECT rect;
		SendMessage(hWnd, EM_GETRECT, 0, (LPARAM)&rect);
		LONG_PTR eax = SendMessage(hWnd, EM_CHARFROMPOS, 0, (LPARAM)&rect);
		eax = SendMessage(hWnd, EM_LINEFROMCHAR, eax, 0);
		TEXTRANGE txtrange = {};
		txtrange.chrg.cpMin = (LONG)SendMessage(hWnd, EM_LINEINDEX, eax, 0);
		txtrange.chrg.cpMax = (LONG)SendMessage(hWnd, EM_CHARFROMPOS, 0, (LPARAM)&rect.right);
		LONG_PTR FirstChar = txtrange.chrg.cpMin;
		HRGN hRgn = CreateRectRgn(rect.left, rect.top, rect.right, rect.bottom);
		HDC hdc = GetDC(hWnd);
		SetBkMode(hdc, TRANSPARENT);
		HRGN hOldRgn = (HRGN)SelectObject(hdc, hRgn);
		DWORD dwTextSize = txtrange.chrg.cpMax - txtrange.chrg.cpMin;
		if (dwTextSize > 0) {
			LPWSTR pBuffer = (LPWSTR)GlobalAlloc(0, (dwTextSize + 1) * sizeof(WCHAR));
			if (pBuffer) {
				txtrange.lpstrText = pBuffer;
				if (SendMessage(hWnd, EM_GETTEXTRANGE, 0, (LPARAM)&txtrange) > 0) {
					LPWSTR lpszStart = wcsstr(txtrange.lpstrText, L"[");
					while (lpszStart) {
						while (*(lpszStart + 1) == L'[') {
							lpszStart++;
						}
						LPWSTR lpszEnd = wcsstr(lpszStart, L"]");
						if (lpszEnd) {
							const int nTextLen = (int)(lpszEnd - lpszStart);
							LPWSTR lpszText = (LPWSTR)GlobalAlloc(0, nTextLen * sizeof(WCHAR));
							if (lpszText) {
								(void)lstrcpyn(lpszText, lpszStart + 1, nTextLen);
								if (IsHighlightText(lpszText)) {
									SendMessage(hWnd, EM_POSFROMCHAR, (WPARAM)&rect, lpszStart - txtrange.lpstrText + FirstChar);
									COLORREF colOld = SetTextColor(hdc, RGB(128, 128, 255));
									HFONT hOldFont = (HFONT)SelectObject(hdc, hFont);
									DrawText(hdc, lpszStart, nTextLen + 1, &rect, 0);
									SelectObject(hdc, hOldFont);
								}
								GlobalFree(lpszText);
							}
						}
						else {
							break;
						}
						lpszStart = wcsstr(lpszEnd, L"[");
					}
				}
				GlobalFree(pBuffer);
			}
		}
		SelectObject(hdc, hOldRgn);
		DeleteObject(hRgn);
		ReleaseDC(hWnd, hdc);
		ShowCaret(hWnd);
		return 0;
	}
	break;
	default:
		break;
	}
	return CallWindowProc(DefaultRichEditProc, hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hEdit;
	static UINT uDpiX = DEFAULT_DPI, uDpiY = DEFAULT_DPI;
	switch (msg)
	{
	case WM_CREATE:
		LoadLibrary(L"riched20");
		hEdit = CreateWindowEx(WS_EX_CLIENTEDGE, RICHEDIT_CLASSW, 0, WS_VISIBLE | WS_CHILD | WS_VSCROLL | WS_HSCROLL | ES_MULTILINE | ES_NOHIDESEL, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		SendMessage(hEdit, EM_SETEDITSTYLE, SES_EMULATESYSEDIT, SES_EMULATESYSEDIT);
		SendMessage(hEdit, EM_LIMITTEXT, -1, 0);
		DefaultRichEditProc = (WNDPROC)SetWindowLongPtr(hEdit, GWLP_WNDPROC, (LONG_PTR)RichEditProc);
		SendMessage(hWnd, WM_APP, 0, 0);
		break;
	case WM_APP:
		GetScaling(hWnd, &uDpiX, &uDpiY);
		DeleteObject(hFont);
		hFont = CreateFontW(-POINT2PIXEL(32), 0, 0, 0, FW_NORMAL, 0, 0, 0, SHIFTJIS_CHARSET, 0, 0, 0, 0, L"Yu Gothic UI");
		SendMessage(hEdit, WM_SETFONT, (WPARAM)hFont, 0);
		break;
	case WM_SETFOCUS:
		SetFocus(hEdit);
		break;
	case WM_SIZE:
		MoveWindow(hEdit, POINT2PIXEL(10), POINT2PIXEL(10), LOWORD(lParam) - POINT2PIXEL(20), HIWORD(lParam) - POINT2PIXEL(20), TRUE);
		break;
	case WM_NCCREATE:
		{
			const HMODULE hModUser32 = GetModuleHandle(TEXT("user32"));
			if (hModUser32)
			{
				typedef BOOL(WINAPI* fnTypeEnableNCScaling)(HWND);
				const fnTypeEnableNCScaling fnEnableNCScaling = (fnTypeEnableNCScaling)GetProcAddress(hModUser32, "EnableNonClientDpiScaling");
				if (fnEnableNCScaling)
				{
					fnEnableNCScaling(hWnd);
				}
			}
		}
		return DefWindowProc(hWnd, msg, wParam, lParam);
	case WM_DPICHANGED:
		SendMessage(hWnd, WM_APP, 0, 0);
		break;
	case WM_DESTROY:
		DeleteObject(hFont);
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("Window"),
		WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}
