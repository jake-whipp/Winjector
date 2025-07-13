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
#define SELECT_PROCESS_BUTTON 1003
#define INJECT_DLL_BUTTON 1004

struct ProcessInfo {
	std::wstring processName;
	DWORD processID;

	ProcessInfo()
	{
		processName = L"";
		processID = NULL;
	}
};

// Structure to hold the information about the application's state.
struct StateInfo {
	HWND hProcessListView = NULL;
	ProcessInfo selectedProcess = ProcessInfo();

	std::wstring dllPath = L"";
	std::wstring dllDisplayName = L"";
};


LRESULT CALLBACK WindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void AddAppControls(HWND hWnd, StateInfo* pAppState);
void OpenFile(HWND hWnd, StateInfo* pAppState);
void UpdateProcessList(HWND hListView);
inline StateInfo* GetAppState(HWND hWnd);
DWORD Inject(HWND hWnd, StateInfo* pAppState);


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

	// Create the base window
	HWND hWnd = CreateWindowEx(
		0,						// Optional window styles
		CLASS_NAME,				// Same name as window class
		L"Winjector Tool",		// Window name
		WS_FIXED,				// Window style

		// Position
		CW_USEDEFAULT, CW_USEDEFAULT,

		// Size
		APP_WIDTH, APP_HEIGHT,

		NULL,		// No parent window
		NULL,		
		hInstance,	// Window instance handle
		pState		// Extra application data
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
				{
					OpenFile(hWnd, pApplicationState);
					break;
				}
				

			case INJECT_DLL_BUTTON:
				{
					int result = Inject(hWnd, pApplicationState);

					// According to MSFT doc in the case of GetExitCodeThread: "if the function succeeds, the return value is nonzero."
					// further: "if the function fails, the return value is zero".
					if (result == 0)
					{
						MessageBox(hWnd, L"DLL injected, but returned an error code.", L"Error", MB_OK | MB_ICONERROR);
					}
					// -1 means there was an error with the injection process (which we have already covered inside the function)
					else if (result > 0) 
					{
						MessageBox(hWnd, L"DLL was injected successfully.", L"Success", MB_OK | MB_ICONINFORMATION);
					}
					break;
				}
				
			}

			break;
		}

	case WM_NOTIFY:
		{
			LPNMHDR pNotifMsgHeader = (LPNMHDR)lParam;

			if (pNotifMsgHeader->idFrom == PROCESS_LIST_VIEW 
				&& pNotifMsgHeader->code == NM_DBLCLK)
			{
				NMLISTVIEW* pNMListView = (NMLISTVIEW*)lParam;

				int iItem = pNMListView->iItem;
				if (iItem >= 0)
				{
					LVITEM listViewItem = { 0 };
					listViewItem.iItem = iItem;
					listViewItem.mask = LVIF_PARAM;

					if (ListView_GetItem(pNotifMsgHeader->hwndFrom, &listViewItem))
					{
						// Update the PID of the state tracker
						DWORD processID = (DWORD)listViewItem.lParam;
						pApplicationState->selectedProcess.processID = processID;

						// Update the Process Name of the state tracker (for displaying to user)
						wchar_t buffer[100] = L"No Application";
						LPWSTR procName = buffer;
						ListView_GetItemText(pNotifMsgHeader->hwndFrom, iItem, 0, procName, 260);

						pApplicationState->selectedProcess.processName = procName;

						// Invoke text update for selected process
						InvalidateRect(hWnd, NULL, TRUE);
					}
				}
			}
		}

	case WM_PAINT:
		{
			PAINTSTRUCT paintStruct;
			HDC hdc = BeginPaint(hWnd, &paintStruct);

			// (All painting begins here, between the call to BeginPaint and EndPaint).

			// pass in the entire update region as the 2nd parameter (area to paint)
			FillRect(hdc, &paintStruct.rcPaint, (HBRUSH)(COLOR_WINDOW + 1));

			//TextOutW(hdc, 200, 10, L"Winjector", strlen("Winjector"));

			
			StateInfo* pState = GetAppState(hWnd);
			if (pState)
			{
				// Show currently selected DLL
				TextOut(hdc, 200, 
					320, 
					pState->dllDisplayName.c_str(), 
					(int)pState->dllDisplayName.length()
				);

				// Show currently selected process
				std::wstring processDisplayText = pState->selectedProcess.processName +
					L" (" + std::to_wstring(pState->selectedProcess.processID) + L")";

				TextOut(hdc, 
					200,
					360,
					processDisplayText.c_str(),
					(int)processDisplayText.length()
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
	listViewItem.mask = LVIF_TEXT | LVIF_IMAGE | LVIF_PARAM;

	for (auto p : processes)
	{
		listViewItem.pszText = (LPWSTR)p.processName.c_str();
		listViewItem.iImage = 0;
		listViewItem.iItem = 0;
		listViewItem.lParam = p.processID;

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

DWORD Inject(HWND hWnd, StateInfo* pAppState)
{
	if (pAppState->selectedProcess.processID == 0)
	{
		MessageBox(hWnd, L"You need to select a process before injecting", L"Error", MB_OK | MB_ICONERROR);
		return -1;
	}
	
	if (pAppState->dllPath == L"")
	{
		MessageBox(hWnd, L"You need to select a DLL before injecting", L"Error", MB_OK | MB_ICONERROR);
		return -1;
	}

	HANDLE hProcess = OpenProcess(PROCESS_ALL_ACCESS, true, pAppState->selectedProcess.processID);


	// Allocate extra memory in the process so we can store the path of our DLL.
	LPVOID pDllPathAddress = VirtualAllocEx(
		hProcess,				
		NULL,					// Don't care path is allocated
		(SIZE_T)(pAppState->dllPath.length() + 1) * sizeof(wchar_t), // wchar takes 2 bytes per character
		MEM_COMMIT,				// we want to reserve and access the memory immediately
		PAGE_READWRITE			// read and write permissions on that memory
	);

	if (pDllPathAddress == nullptr)
	{
		MessageBox(hWnd, L"Could not allocate extra memory to process", L"Error", MB_OK | MB_ICONERROR);
		return -1;
	}

	// Write DLL path
	SIZE_T noBytesWritten = 0;

	WriteProcessMemory(hProcess, 
		pDllPathAddress, 
		(LPWSTR)pAppState->dllPath.c_str(),			// buffer to write
		(SIZE_T)(pAppState->dllPath.length() + 1) * sizeof(wchar_t),
		&noBytesWritten
	);

	if (noBytesWritten == 0)
	{
		MessageBox(hWnd, L"Failed to inject DLL", L"Error", MB_OK | MB_ICONERROR);
		return -1;
	}

	// We want to get a pointer to the LoadLibraryA method, but it exists inside the kernel32 dll, 
	// so we firstly need a handle to the Kernel32 DLL.
	HMODULE kernel32base = GetModuleHandle(L"kernel32.dll");
	FARPROC pLoadLibrary = GetProcAddress(kernel32base, "LoadLibraryW");

	// Create the remote thread to execute the LoadLibrary in
	HANDLE thread = CreateRemoteThread(
		hProcess,
		NULL,
		0,
		(LPTHREAD_START_ROUTINE)pLoadLibrary,	// ptr to LoadLibraryA
		pDllPathAddress,						// parameters to the call
		0,
		NULL
	);

	if (thread == nullptr)
	{
		MessageBox(hWnd, L"Failed to inject DLL", L"Error", MB_OK | MB_ICONERROR);
		return -1;
	}


	// Wait for the LoadLibrary to finish executing in the remote thread
	
	// According to MSFT doc in the case of GetExitCodeThread: "if the function succeeds, the return value is nonzero."
	// further: "if the function fails, the return value is zero".
	DWORD exitCode = 1;
	WaitForSingleObject(thread, INFINITE);
	GetExitCodeThread(thread, &exitCode);

	// Free the extra allocated memory, and close open handles to the process
	VirtualFreeEx(hProcess, pDllPathAddress, 0, MEM_RELEASE);
	CloseHandle(thread);
	CloseHandle(hProcess);

	return exitCode;
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

	// Select Process Button
	CreateWindow(
		L"BUTTON",
		L"Select Process",
		WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
		30, 360,	// Coordinates
		150, 30,	// Size
		hWnd,
		(HMENU)SELECT_PROCESS_BUTTON,
		(HINSTANCE)GetWindowLongPtr(hWnd, GWLP_HINSTANCE),
		NULL
	);

	// Inject Button
	CreateWindow(
		L"BUTTON",
		L"Inject DLL",
		WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
		30, 400,	// Coordinates
		150, 30,	// Size
		hWnd,
		(HMENU)INJECT_DLL_BUTTON,
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