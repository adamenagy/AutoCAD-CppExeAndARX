// CppDriveAutoCAD.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "CppDriveAutoCAD.h"

#import "C:\Program Files\Common Files\Autodesk Shared\acax20enu.tlb" 
#import "C:\Program Files\Common Files\Autodesk Shared\axdb20enu.tlb"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// The one and only application object

CWinApp theApp;

using namespace std;

struct MyMessageFilter : public IMessageFilter
{
	// IMessageFilter
	DWORD STDMETHODCALLTYPE HandleInComingCall(
		DWORD dwCallType, HTASK hTaskCaller,
		DWORD dwTickCount, LPINTERFACEINFO lpInterfaceInfo
		)
	{
		printf(">> MyMessageFilter::HandleInComingCall\n");
		return 0; // SERVERCALL_ISHANDLED
	}

	DWORD STDMETHODCALLTYPE RetryRejectedCall(
		HTASK hTaskCallee, DWORD dwTickCount, DWORD dwRejectType
		)
	{
		printf(">> MyMessageFilter::RetryRejectedCall\n");
		return 1000; // Retry in a second
	}

	DWORD STDMETHODCALLTYPE MessagePending(
		HTASK hTaskCallee, DWORD dwTickCount, DWORD dwPendingType
		)
	{
		printf(">> MyMessageFilter::MessagePending\n");
		return 1; // PENDINGMSG_WAITNOPROCESS
	}

	// IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface( 
		/* [in] */ REFIID riid,
		/* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject)
	{
		printf(">> MyMessageFilter::QueryInterface\n");
		return S_OK;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef( void)
	{
		printf(">> MyMessageFilter::AddRef\n");
		return 0;
	}

	virtual ULONG STDMETHODCALLTYPE Release( void)
	{
		printf(">> MyMessageFilter::Release\n");
		return 0;
	}
};

void driveAutoCAD()
{
	HRESULT hr;

	CoInitialize(NULL);

	MyMessageFilter myMF;
	LPMESSAGEFILTER oldMF;
	CoRegisterMessageFilter(&myMF, &oldMF);

	AutoCAD::IAcadApplicationPtr m_acPtr;
	CLSID clsid;
	hr = ::CLSIDFromProgID(L"AutoCAD.Application.20", &clsid);
	printf("After CLSIDFromProgID; hr = 0x%08lx\n", hr);

	for (int i = 0; i < 10; i++)
	{ 
		printf("i = %i\n", i);
		printf("Before CreateInstance\n");
		hr = m_acPtr.CreateInstance(clsid,NULL,CLSCTX_LOCAL_SERVER);
		printf("After CreateInstance; hr = 0x%08lx\n", hr);

		BSTR caption;
		hr = m_acPtr->get_Caption(&caption);
		printf("After get_Caption; caption = %S; hr = 0x%08lx\n", caption, hr);
		
		// just for debugging
		//m_acPtr->Visible = VARIANT_TRUE; 

		// The path of the arx is already added to the 
		// "Options >> Files >> Trusted Locations" inside AutoCAD 
		// to avoid the "File Loading - Security Concern" dialog
		// MODIFY below path to make sure it's correct
		hr = m_acPtr->raw_LoadArx(_bstr_t("C:\\Users\\adamnagy\\Documents\\Visual Studio 2012\\Projects\\ArxSleepTest\\x64\\Debug\\ArxSleepTest.arx"));
		printf("After LoadArx; hr = 0x%08lx\n", hr);
		AutoCAD::IAcadDocumentPtr m_docPtr; 
		CComVariant vtOptional;
		vtOptional.vt = VT_ERROR;
		vtOptional.scode = DISP_E_PARAMNOTFOUND;
		// MODIFY below path to make sure it's correct
		hr = m_acPtr->Documents->raw_Open(_bstr_t("C:\\Users\\adamnagy\\Documents\\Visual Studio 2012\\Projects\\CppDriveAutoCAD\\rectangle.dwg"), vtOptional, vtOptional, &m_docPtr); 
		printf("After Documents->Open; hr = 0x%08lx\n", hr);
		hr = m_docPtr->raw_SendCommand(_bstr_t("_MyCommand ")); // note the space at the end
		printf("After SendCommand; hr = 0x%08lx\n", hr);

		// does not seem to work :-/
		//hr = m_acPtr->raw_Quit();
		// use this instead
		hr = m_docPtr->raw_SendCommand(_bstr_t("_.QUIT ")); // note the space at the end
		printf("After QUIT; hr = 0x%08lx\n", hr);

		m_acPtr = NULL;
	}

	CoUninitialize(); 

	// Wait for user to press enter
	printf("Press enter to finish...");
	getchar(); 
}

int _tmain(int argc, TCHAR* argv[], TCHAR* envp[])
{
	int nRetCode = 0;

	HMODULE hModule = ::GetModuleHandle(NULL);

	if (hModule != NULL)
	{
		// initialize MFC and print and error on failure
		if (!AfxWinInit(hModule, NULL, ::GetCommandLine(), 0))
		{
			// TODO: change error code to suit your needs
			_tprintf(_T("Fatal Error: MFC initialization failed\n"));
			nRetCode = 1;
		}
		else
		{
			// TODO: code your application's behavior here.
			driveAutoCAD();
		}
	}
	else
	{
		// TODO: change error code to suit your needs
		_tprintf(_T("Fatal Error: GetModuleHandle failed\n"));
		nRetCode = 1;
	}

	return nRetCode;
}
