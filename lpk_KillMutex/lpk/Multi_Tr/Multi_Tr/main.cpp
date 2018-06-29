#include <Windows.h>
#include <stdio.h>
#include <winternl.h>

#pragma comment(linker, "/SECTION:.text,REW" )
#pragma comment(linker, "/EXPORT:LpkTabbedTextOut=_AheadLib_LpkTabbedTextOut,@1")
#pragma comment(linker, "/EXPORT:LpkDllInitialize=_AheadLib_LpkDllInitialize,@2")
#pragma comment(linker, "/EXPORT:LpkDrawTextEx=_AheadLib_LpkDrawTextEx,@3")
#pragma comment(linker, "/EXPORT:LpkEditControl=_AheadLib_LpkEditControl,@4")
#pragma comment(linker, "/EXPORT:LpkExtTextOut=_AheadLib_LpkExtTextOut,@5")
#pragma comment(linker, "/EXPORT:LpkGetCharacterPlacement=_AheadLib_LpkGetCharacterPlacement,@6")
#pragma comment(linker, "/EXPORT:LpkGetTextExtentExPoint=_AheadLib_LpkGetTextExtentExPoint,@7")
#pragma comment(linker, "/EXPORT:LpkInitialize=_AheadLib_LpkInitialize,@8")
#pragma comment(linker, "/EXPORT:LpkPSMTextOut=_AheadLib_LpkPSMTextOut,@9")
#pragma comment(linker, "/EXPORT:LpkUseGDIWidthCache=_AheadLib_LpkUseGDIWidthCache,@10")
#pragma comment(linker, "/EXPORT:ftsWordBreak=_AheadLib_ftsWordBreak,@11")

typedef LONG							NTSTATUS;
//#define NT_SUCCESS(Status)				((NTSTATUS)(Status) >= 0)
#define SystemHandleInformation			0x10
#define STATUS_INFO_LENGTH_MISMATCH     ((NTSTATUS)0xC0000004L)

typedef struct _SYSTEM_HANDLE_INFORMATION {
    ULONG       ProcessId;
    UCHAR       ObjectTypeNumber;
    UCHAR       Flags;
    USHORT      Handle;
    PVOID       Object;
    ACCESS_MASK GrantedAccess;
} SYSTEM_HANDLE_INFORMATION, *PSYSTEM_HANDLE_INFORMATION;

typedef struct _SYSTEM_HANDLE_INFORMATION_EX {
	ULONG NumberOfHandles;
	SYSTEM_HANDLE_INFORMATION Information[1];
} SYSTEM_HANDLE_INFORMATION_EX, *PSYSTEM_HANDLE_INFORMATION_EX;

typedef NTSTATUS (__stdcall *NTQUERYSYSTEMINFORMATION)(
													   IN	DWORD				  SystemInformationClass,  //就是enum類型的
													   OUT	PVOID                 SystemInformation,
													   IN	ULONG                 Length,
													   OUT	PULONG                ReturnLength
													   );

void	WINAPI		MyRemoveMutexThread	( LPVOID lpParam );
NTQUERYSYSTEMINFORMATION	NtQuerySysteminformation;
BOOL GetCurrentProcessName(LPTSTR lpstrName);   

#define EXTERNC extern "C"
#define NAKED __declspec(naked)
#define EXPORT __declspec(dllexport)

#define ALCPP EXPORT NAKED
#define ALSTD EXTERNC EXPORT NAKED void __stdcall
#define ALCFAST EXTERNC EXPORT NAKED void __fastcall
#define ALCDECL EXTERNC NAKED void __cdecl

namespace AheadLib
{
	HMODULE m_hModule = NULL;
	DWORD m_dwReturn[11] = {0};

	// 加載原始模塊
	inline BOOL WINAPI Load()
	{
		TCHAR tzPath[MAX_PATH];
		TCHAR tzTemp[MAX_PATH * 2];
		GetSystemDirectory((LPSTR)tzPath,MAX_PATH);
		strcat_s(tzPath,"\\lpk.dll");
		m_hModule = LoadLibrary(tzPath);
		if (m_hModule == NULL)
		{
			wsprintf(tzTemp, TEXT("無法加載 %s，程序無法正常運行。"), tzPath);
			MessageBox(NULL, tzTemp, TEXT("AheadLib"), MB_ICONSTOP);
		}

		return (m_hModule != NULL);	
	}
		
	inline VOID WINAPI Free()
	{
		if (m_hModule)
		{
			FreeLibrary(m_hModule);
		}
	}

	FARPROC WINAPI GetAddress(PCSTR pszProcName)
	{
		FARPROC fpAddress;
		CHAR szProcName[16];
		TCHAR tzTemp[MAX_PATH];

		fpAddress = GetProcAddress(m_hModule, pszProcName);
		if (fpAddress == NULL)
		{
			if (HIWORD(pszProcName) == 0)
			{
				wsprintf(szProcName, "%d", pszProcName);
				pszProcName = szProcName;
			}

			wsprintf(tzTemp, TEXT("無法找到函數 %hs，程序無法正常運行。"), pszProcName);
			MessageBox(NULL, tzTemp, TEXT("AheadLib"), MB_ICONSTOP);
			ExitProcess(-2);
		}

		return fpAddress;
	}
}
using namespace AheadLib;

__inline BOOL MyInitNtApi()
{
	HMODULE hNtDll;

	if ( ( hNtDll = GetModuleHandle("ntdll") ) == NULL)
		return FALSE;

	if ( ( NtQuerySysteminformation = (NTQUERYSYSTEMINFORMATION)GetProcAddress(hNtDll,"NtQuerySystemInformation") ) == NULL )
		return FALSE;

	return TRUE;
}

__inline VOID MyCloseSomeMutex(USHORT hInputMutex)
{
	if ( !MyInitNtApi() ){
		return;
	}
	
	
	DWORD dwBufferLength = 327680;
	PBYTE pBuffer = new BYTE[dwBufferLength];

	if ( !pBuffer )
		return;

	BOOL		bNtFunSuccess     = FALSE;
    NTSTATUS    status;

	while(1){
		
		//NT_SUCCESS
		if ( ( status = NtQuerySysteminformation( SystemHandleInformation, pBuffer, dwBufferLength, NULL ) ) >= 0 ){
			bNtFunSuccess = TRUE;
			break;
		}
		
		if ( status == STATUS_INFO_LENGTH_MISMATCH  ){
			delete[] pBuffer;
			dwBufferLength *= 2;
			pBuffer = new BYTE[dwBufferLength];
			if ( !pBuffer )
				break;
		}
		else{
			delete[] pBuffer;
			break;
		}
		
	}

	if ( !bNtFunSuccess )
		return;

	ULONG uHandleCount = *(PULONG)pBuffer;
	PSYSTEM_HANDLE_INFORMATION pInfo = (PSYSTEM_HANDLE_INFORMATION)(pBuffer+4);
	ULONG		i;
	PVOID		Object					= NULL;
	UCHAR       ObjectTypeNumber		= 0;

	ULONG pId = GetCurrentProcessId();

	for (i=0; i< uHandleCount; pInfo++,i++){
		if (pInfo->ProcessId == pId && pInfo->Handle == hInputMutex){
			ObjectTypeNumber = pInfo->ObjectTypeNumber;
			Object = pInfo->Object;
			break;
		}

	}

	if ( !Object ){ 
		delete[] pBuffer;
		return;
	}

	pInfo = (PSYSTEM_HANDLE_INFORMATION)(pBuffer+4);

	for (i=0; i< uHandleCount; pInfo++,i++){
		if ( pInfo->ProcessId == pId  && pInfo->ObjectTypeNumber == ObjectTypeNumber && pInfo->Object == Object ){
			CloseHandle((HANDLE)pInfo->Handle);
		}
	}
	delete[] pBuffer;
}
BOOL GetCurrentProcessName(LPTSTR lpstrName)   
{   
    int i=0;   
    TCHAR lpstrFullPath[1024]=TEXT("");   
    GetModuleFileName(GetModuleHandle(NULL),lpstrFullPath,1024);   
    for(i=lstrlen(lpstrFullPath);i>0;i--)   
    {   
        if (lpstrFullPath[i] == '\\')   
        {   
            break;   
        }   
    }   
    lstrcpy(lpstrName,lpstrFullPath+i+1);   
    return TRUE;   
}   

void WINAPI MyRemoveMutexThread(LPVOID lpParam )
{
	HANDLE hMutex;
	while(1){
		Sleep(1000);
		hMutex = OpenMutex (MUTEX_ALL_ACCESS,FALSE,"rhaon_co_kr__TalesRunnerMutex");
		if (   hMutex == (HANDLE)NULL || hMutex == (HANDLE)ERROR_FILE_NOT_FOUND)
			Sleep(1000);
		else
			break;  
	}

	__try{
		MyCloseSomeMutex((USHORT)hMutex);
	}
	
	__except(EXCEPTION_EXECUTE_HANDLER){
		return;
	}
}

BOOL WINAPI DllMain(HMODULE hModule, DWORD dwReason, PVOID pvReserved)
{
    if (dwReason == DLL_PROCESS_ATTACH)
    {
		DWORD dwThreadID;
		HANDLE hThread;
		
        DisableThreadLibraryCalls(hModule);
		char buf[64];
		HWND hWnd= FindWindow(NULL,"Tales Runner");
		if (hWnd)
		{
			_snprintf(buf, sizeof(buf), "Tales Runner");
			SetWindowText(hWnd, buf);
		}
		CloseHandle( CreateThread(NULL,0,(LPTHREAD_START_ROUTINE)MyRemoveMutexThread,NULL,NULL,NULL) );
		return Load();
    }
    else if (dwReason == DLL_PROCESS_DETACH)
    {
        Free();
    }

    return TRUE;
}

bool EnableMultiTrgame()
{
		return true;
}

ALCDECL AheadLib_LpkTabbedTextOut(void)
{
	// 保存返回地址
	__asm POP m_dwReturn[0 * TYPE long];

	// 調用原始函數
	GetAddress("LpkTabbedTextOut")();

	// 轉跳到返回地址
	__asm JMP m_dwReturn[0 * TYPE long];
}

ALCDECL AheadLib_LpkDllInitialize(void)
{
	// 保存返回地址
	__asm POP m_dwReturn[1 * TYPE long];

	// 調用原始函數
	GetAddress("LpkDllInitialize")();

	// 轉跳到返回地址
	__asm JMP m_dwReturn[1 * TYPE long];
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 導出函數
ALCDECL AheadLib_LpkDrawTextEx(void)
{

	// 保存返回地址
	__asm POP m_dwReturn[2 * TYPE long];

	// 調用原始函數
	GetAddress("LpkDrawTextEx")();

	// 轉跳到返回地址
	__asm JMP m_dwReturn[2 * TYPE long];
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 導出函數
 ALCDECL AheadLib_LpkEditControl(void)
 {

 	// 保存返回地址
 	__asm POP m_dwReturn[3 * TYPE long];
 
 	// 調用原始函數
 	GetAddress("LpkEditControl")();

 	// 轉跳到返回地址
 	__asm JMP m_dwReturn[3 * TYPE long];
 }


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 導出函數
ALCDECL AheadLib_LpkExtTextOut(void)
{

	// 保存返回地址
	__asm POP m_dwReturn[4 * TYPE long];

	// 調用原始函數
	GetAddress("LpkExtTextOut")();

	// 轉跳到返回地址
	__asm JMP m_dwReturn[4 * TYPE long];
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 導出函數
ALCDECL AheadLib_LpkGetCharacterPlacement(void)
{

	// 保存返回地址
	__asm POP m_dwReturn[5 * TYPE long];

	// 調用原始函數
	GetAddress("LpkGetCharacterPlacement")();

	// 轉跳到返回地址
	__asm JMP m_dwReturn[5 * TYPE long];
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 導出函數
ALCDECL AheadLib_LpkGetTextExtentExPoint(void)
{

	// 保存返回地址
	__asm POP m_dwReturn[6 * TYPE long];

	// 調用原始函數
	GetAddress("LpkGetTextExtentExPoint")();

	// 轉跳到返回地址
	__asm JMP m_dwReturn[6 * TYPE long];
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 導出函數
ALCDECL AheadLib_LpkInitialize(void)
{
	// 保存返回地址
	__asm POP m_dwReturn[7 * TYPE long];

	// 調用原始函數
	GetAddress("LpkInitialize")();
	
	// 轉跳到返回地址
	__asm JMP m_dwReturn[7 * TYPE long];
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 導出函數
ALCDECL AheadLib_LpkPSMTextOut(void)
{

	// 保存返回地址
	__asm POP m_dwReturn[8 * TYPE long];

	// 調用原始函數
	GetAddress("LpkPSMTextOut")();

	// 轉跳到返回地址
	__asm JMP m_dwReturn[8 * TYPE long];
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 導出函數
ALCDECL AheadLib_LpkUseGDIWidthCache(void)
{

	// 保存返回地址
	__asm POP m_dwReturn[9 * TYPE long];

	// 調用原始函數
	GetAddress("LpkUseGDIWidthCache")();

	// 轉跳到返回地址
	__asm JMP m_dwReturn[9 * TYPE long];
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 導出函數
ALCDECL AheadLib_ftsWordBreak(void)
{

	// 保存返回地址
	__asm POP m_dwReturn[10 * TYPE long];

	// 調用原始函數
	GetAddress("ftsWordBreak")();

	// 轉跳到返回地址
	__asm JMP m_dwReturn[10 * TYPE long];
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
