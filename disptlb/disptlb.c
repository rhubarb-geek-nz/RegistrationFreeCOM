/***********************************
 * Copyright (c) 2024 Roger Brown.
 * Licensed under the MIT License.
 ****/

#include <windows.h>
#include <objbase.h>

static HMODULE globalModule;
static OLECHAR globalModuleFileName[260];

STDAPI DllRegisterServer(void)
{
	ITypeLib* piTypeLib = NULL;
	HRESULT hr = LoadTypeLibEx(globalModuleFileName, REGKIND_REGISTER, &piTypeLib);

	if (hr >= 0)
	{
		ITypeLib_Release(piTypeLib);
	}

	return hr;
}

STDAPI DllUnregisterServer(void)
{
	ITypeLib* piTypeLib = NULL;
	HRESULT hr = LoadTypeLibEx(globalModuleFileName, REGKIND_NONE, &piTypeLib);

	if (hr >= 0)
	{
		TLIBATTR* attr = NULL;

		hr = ITypeLib_GetLibAttr(piTypeLib, &attr);

		if (hr >= 0)
		{
			hr = UnRegisterTypeLib(&(attr->guid), attr->wMajorVerNum, attr->wMinorVerNum, attr->lcid, attr->syskind);

			ITypeLib_ReleaseTLibAttr(piTypeLib, attr);
		}

		ITypeLib_Release(piTypeLib);
	}

	return hr;
}

BOOL APIENTRY DllMain(HMODULE hModule,
	DWORD  ul_reason_for_call,
	LPVOID lpReserved
)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		globalModule = hModule;
		DisableThreadLibraryCalls(hModule);

		if (!GetModuleFileNameW(globalModule, globalModuleFileName, sizeof(globalModuleFileName) / sizeof(globalModuleFileName[0])))
		{
			return FALSE;
		}

		break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
		break;
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}
