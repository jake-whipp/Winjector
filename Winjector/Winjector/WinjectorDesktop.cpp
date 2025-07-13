#include <Windows.h>
#include <new>
#include <ShObjIdl.h>
#include <TlHelp32.h>
#include <atlbase.h>
#include <string>
#include <vector>

#define WS_FIXED (WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX )
#define APP_WIDTH 500
#define APP_HEIGHT 500
#define OPEN_FILE_BUTTON 1001
#define PROCESS_LIST_VIEW 1002

// Structure to hold the information about the application's state.
struct StateInfo {
	HWND hProcessListView = NULL;

	std::wstring dllPath = L"";
	std::wstring dllDisplayName = L"";
};

struct ProcessInfo {
	std::wstring processName = L"";
	DWORD processID = NULL;
};

LRESULT CALLBACK WindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void AddAppControls(HWND hWnd, StateInfo* pAppState);
void OpenFile(HWND hWnd, StateInfo* pAppState);
void UpdateProcessList(HWND hListView);
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

	// Create the application window instance
	HWND hWnd = CreateWindowEx(
		0,						// Optional window styles
		CLASS_NAME,				// Same name as window class
		L"Winjector Tool",		// Window name
		WS_FIXED,				// Window style

		// Position
		CW_USEDEFAULT, CW_USEDEFAULT,

		// Size
		APP_WIDTH, APP_HEIGHT,

		NULL,		// Parent window
		NULL,		// Menu
		hInstance,	// Window instance handle
		pState		// Parameters (extra application data)
	);

	if (hWnd == NULL)
	{
		return 0;
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
	StateInfo* pApplicationState = nullptr;
		

	if (uMsg == WM_CREATE)
	{
		CREATESTRUCT* pCreate = reinterpret_cast<CREATESTRUCT*>(lParam);

		pApplicationState = reinterpret_cast<StateInfo*>(pCreate->lpCreateParams);

		// Set the pointer in the instance data for the window.
		// it can then later be retrieved anytime by a call to GetWindowLongPtr
		SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR)pApplicationState);

		AddAppControls(hWnd, pApplicationState);
	}
	else
	{
		// For any other message that comes in on the queue, we want to use the current instance data
		pApplicationState = GetAppState(hWnd);
	}

	switch (uMsg)
	{
	case WM_COMMAND:
		{
			switch (LOWORD(wParam))
			{
			case OPEN_FILE_BUTTON:
				OpenFile(hWnd, pApplicationState);
				break;
			}

			break;
		}

	case WM_PAINT:
		{
			PAINTSTRUCT paintStruct;
			HDC hdc = BeginPaint(hWnd, &paintStruct);

			// (All painting begins here, between the call to BeginPaint and EndPaint).

			// pass in the entire update region as the 2nd parameter (area to paint)
			FillRect(hdc, &paintStruct.rcPaint, (HBRUSH)(COLOR_WINDOW + 1));

			//TextOutW(hdc, 200, 10, L"Winjector", strlen("Winjector"));

			// Show currently selected DLL
			StateInfo* pState = GetAppState(hWnd);
			if (pState)
			{
				TextOutW(hdc, 200, 
					320, 
					pState->dllDisplayName.c_str(), 
					(int)pState->dllDisplayName.length()
				);
			}

			EndPaint(hWnd, &paintStruct);
		}
		break;


	case WM_SIZE:
		{
			int width = LOWORD(lParam);
			int height = HIWORD(lParam);
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


int GetActiveProcesses(HWND hWnd, std::vector<ProcessInfo>* processesVector)
{
	HANDLE hProcessSnapshot = NULL;

	PROCESSENTRY32 processEntry = { NULL };
	processEntry.dwSize = sizeof(PROCESSENTRY32);

	// Create the snapshot and check if valid
	hProcessSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);

	if (hProcessSnapshot == INVALID_HANDLE_VALUE)
	{
		MessageBox(hWnd, L"Error: process snapshot has invalid handle", L"Error", MB_OK | MB_ICONERROR);
		return -1;
	}

	// Walk the snapshot and build a vector of process information
	std::vector<ProcessInfo> processes;

	if (Process32First(hProcessSnapshot, &processEntry))
	{
		do
		{
			ProcessInfo p;
			p.processID = processEntry.th32ProcessID;
			p.processName = processEntry.szExeFile;

			processes.push_back(p);
		} while (Process32Next(hProcessSnapshot, &processEntry));
	}

	// Cleanup
	CloseHandle(hProcessSnapshot);
	
	*processesVector = processes;

	return 0;
}

void UpdateProcessList(HWND hListView)
{
	std::vector<ProcessInfo> processes;

	if (GetActiveProcesses(hListView, &processes) == -1)
	{
		return;
	}

	LVITEM listViewItem = { 0 };
	listViewItem.mask = LVIF_TEXT | LVIF_IMAGE;

	for (auto p : processes)
	{
		listViewItem.pszText = (LPWSTR)p.processName.c_str();
		listViewItem.iImage = 0;
		listViewItem.iItem = 0;

		int row = ListView_InsertItem(hListView, &listViewItem);

		std::wstring pidText = std::to_wstring(p.processID);
		ListView_SetItemText(hListView, row, 1, (LPWSTR)pidText.c_str());
	}
}


void OpenFile(HWND hWnd, StateInfo* pAppState)
{
	// Create the FileOpenDialog object. The Class ID and Interface ID are defined in
	// The same header as the IFileOpenDialog interface itself (ShObjIdl.h).
	CComPtr<IFileOpenDialog> pFileOpen;
	HRESULT hr = CoCreateInstance(__uuidof(FileOpenDialog), NULL, CLSCTX_ALL, IID_PPV_ARGS(&pFileOpen));

	if (FAILED(hr))
	{
		MessageBox(hWnd, L"Could not open the FileOpenDialog.", L"Error", MB_OK | MB_ICONERROR);
		return;
	}

	// Restrict types of selected files to DLL-only
	const COMDLG_FILTERSPEC fileTypes[] =
	{
		{ L"DLL Files (*.dll)", L"*.dll" }
	};

	pFileOpen->SetFileTypes(ARRAYSIZE(fileTypes), fileTypes);

	// Show the Open dialog box.
	hr = pFileOpen->Show(hWnd);

	CComPtr<IShellItem> pItem;
	hr = pFileOpen->GetResult(&pItem);

	if (FAILED(hr))
	{
		return;
	}

	LPWSTR pszFilePath = nullptr;
	hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);

	if (SUCCEEDED(hr))
	{
		//MessageBox(hWnd, pszFilePath, L"Selected File", MB_OK);
		pAppState->dllPath = pszFilePath;

		CoTaskMemFree(pszFilePath);

		LPWSTR pszFileName = nullptr;
		hr = pItem->GetDisplayName(SIGDN_NORMALDISPLAY, &pszFileName);

		if (SUCCEEDED(hr))
		{
			pAppState->dllDisplayName = pszFileName;
			CoTaskMemFree(pszFileName);
		}

		// Trigger a redraw so the selected DLL text updates
		InvalidateRect(hWnd, NULL, TRUE);
	}
}

// Helper function for getting the instance data of the current window class
inline StateInfo* GetAppState(HWND hWnd)
{
	LONG_PTR lp = GetWindowLongPtr(hWnd, GWLP_USERDATA);
	StateInfo* pState = reinterpret_cast<StateInfo*>(lp);
	return pState;
}

void AddAppControls(HWND hWnd, StateInfo* pAppState)
{
	// Create other windows for the application
	
	// Select DLL Button
	CreateWindow(
		L"BUTTON",
		L"Select DLL",
		WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
		30, 320,	// Coordinates
		150, 30,	// Size
		hWnd,
		(HMENU)OPEN_FILE_BUTTON,
		(HINSTANCE)GetWindowLongPtr(hWnd, GWLP_HINSTANCE),
		NULL
	);


	// Select process ListView
	HWND hProcessListView = CreateWindowEx(0,
		WC_LISTVIEW,
		L"Process select",
		WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL,
		30, 50,		// Coordinates
		420, 250,	// Size
		hWnd,
		(HMENU)PROCESS_LIST_VIEW,
		(HINSTANCE)GetWindowLongPtr(hWnd, GWLP_HINSTANCE),
		NULL
	);

	// Create columns
	LVCOLUMN listViewColumn = { 0 };
	listViewColumn.mask = LVCF_TEXT | LVCF_WIDTH;

	// Column 1 (Process)
	listViewColumn.pszText = (LPWSTR)L"Process Name";
	listViewColumn.cx = 120;
	ListView_InsertColumn(hProcessListView, 0, &listViewColumn);

	// Column 2 (PID)
	listViewColumn.pszText = (LPWSTR)L"PID";
	listViewColumn.cx = 80;
	ListView_InsertColumn(hProcessListView, 1, &listViewColumn);

	// Set default image for all processes
	HIMAGELIST hImageList = ImageList_Create(16, 16, ILC_COLOR32 | ILC_MASK, 1, 1);

	HICON hDefaultIcon = LoadIcon(NULL, IDI_APPLICATION);
	int defaultIconIndex = ImageList_AddIcon(hImageList, hDefaultIcon);

	ListView_SetImageList(hProcessListView, hImageList, LVSIL_SMALL);

	// Populate the listview for the first time
	pAppState->hProcessListView = hProcessListView;
	UpdateProcessList(hProcessListView);
}