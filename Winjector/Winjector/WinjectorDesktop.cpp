#include <Windows.h>
#include <new>
#include <ShObjIdl.h>

// Define a structure to hold some state information.
struct StateInfo {
	const char* selectedAppName = "";
};


LRESULT CALLBACK WindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void OnSize(HWND hwnd, UINT flag, int width, int height);
inline StateInfo* GetAppState(HWND hWnd);

int WINAPI WinMain(
	HINSTANCE hInstance,		// handle to the instance
	HINSTANCE hPrevInstance,	// no meaning; used in 16-bit Windows and is now always zero
	LPSTR lpCmdLine,			// contains the command line args in a string
	int nCmdShow)				// flag to indicate whether the app is minimised, maximised, or regular
{
	HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	if (FAILED(hr))
	{
		return 0;
	}

	// Register the window class
	const wchar_t CLASS_NAME[] = L"Winjector Window Class";

	WNDCLASS wndClass = {};
	wndClass.lpfnWndProc = WindowProc;		// The window procedure defines most of the behavior of the window
	wndClass.hInstance = hInstance;			// Handle to the application instance
	wndClass.lpszClassName = CLASS_NAME;	// String to identify the window class


	// Register the window class with the OS
	RegisterClass(&wndClass);


	// Create pointer to app data on the heap for the window to use 
	StateInfo* pState = new (std::nothrow) StateInfo;
	if (pState == NULL)
	{
		return 0;
	}

	// Create the window instance
	HWND hWnd = CreateWindowEx(
		0,						// Optional window styles
		CLASS_NAME,				// Same name as window class
		L"Winjector Tool",		// Example
		WS_OVERLAPPEDWINDOW,	// Window style

		// Position
		CW_USEDEFAULT, CW_USEDEFAULT,

		// Size
		CW_USEDEFAULT, CW_USEDEFAULT,

		NULL,		// Parent window
		NULL,		// Menu
		hInstance,	// Window instance handle
		pState		// Parameters (extra application data)
	);

	if (hWnd == NULL)
	{
		return 0;
	}

	// Create the FileOpenDialog object. The Class ID and Interface ID are defined in
	// The same header as the IFileOpenDialog interface itself (ShObjIdl.h).
	IFileOpenDialog* pFileOpen;

	hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL,
		IID_IFileOpenDialog, reinterpret_cast<void**>(&pFileOpen));

	if (SUCCEEDED(hr))
	{
		// Show the Open dialog box.
		hr = pFileOpen->Show(NULL);
	}


	// Display the window
	ShowWindow(hWnd, nCmdShow);
	UpdateWindow(hWnd);

	// Begin the message loop
	MSG msg = {};
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	CoUninitialize();
	return (int)msg.wParam;
}


LRESULT CALLBACK WindowProc(
	HWND hWnd,						// handle to the window
	UINT uMsg,						// the message code; for example, a WM_SIZE message indicated the window was resized.
	WPARAM wParam, LPARAM lParam)	// additional data that pertains to the message. The exact meaning depends on the message code.
{
	StateInfo* pApplicationState;

	if (uMsg == WM_CREATE)
	{
		CREATESTRUCT* pCreate = reinterpret_cast<CREATESTRUCT*>(lParam);

		pApplicationState = reinterpret_cast<StateInfo*>(pCreate->lpCreateParams);

		// Set the pointer in the instance data for the window.
		// it can then later be retrieved anytime by a call to GetWindowLongPtr
		SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR)pApplicationState);
	}
	else
	{
		// For any message that comes in on the queue, we want to use the current instance data
		pApplicationState = GetAppState(hWnd);
	}

	switch (uMsg)
	{
	case WM_PAINT:
		{
			PAINTSTRUCT paintStruct;
			HDC hdc = BeginPaint(hWnd, &paintStruct);

			// (All painting begins here, between the call to BeginPaint and EndPaint).

			// pass in the entire update region as the 2nd parameter (area to paint)
			FillRect(hdc, &paintStruct.rcPaint, (HBRUSH)(COLOR_WINDOW + 1));

			EndPaint(hWnd, &paintStruct);
		}
		break;


	case WM_SIZE:
		{
			int width = LOWORD(lParam);
			int height = HIWORD(lParam);

			OnSize(hWnd, (UINT)wParam, width, height);
		}
		break;


	case WM_DESTROY:
		// Stop listening and addressing messages on the queue; 0 is the exit-code
		PostQuitMessage(0);
		return 0;
	}

	// Any other unhandled messages go to the DefWindowProc. 
	// This function performs the default action for the message, which varies by message type.
	return DefWindowProc(hWnd, uMsg, wParam, lParam);
}


void OnSize(HWND hWnd, UINT flag, int width, int height)
{
	// Handle resizing
}


// Helper function for getting the instance data of the current window class
inline StateInfo* GetAppState(HWND hWnd)
{
	LONG_PTR lp = GetWindowLongPtr(hWnd, GWLP_USERDATA);
	StateInfo* pState = reinterpret_cast<StateInfo*>(lp);
	return pState;
}
