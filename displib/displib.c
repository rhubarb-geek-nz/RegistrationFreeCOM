/***********************************
 * Copyright (c) 2024 Roger Brown.
 * Licensed under the MIT License.
 ****/

#include <displib.h>

static LONG globalUsage;
static HMODULE globalModule;
static OLECHAR globalModuleFileName[260];

static const CLSID CLSID_CHelloWorld = { 0x49ef0168, 0x2765, 0x4932, 0xbe, 0x4c, 0xe2, 0x1e, 0x0d, 0x7a, 0x55, 0x4f };

typedef struct CHelloWorldData
{
	IUnknown IUnknown;
	IHelloWorld IHelloWorld;
	LONG lUsage;
	IUnknown* lpOuter;
	ITypeInfo* piTypeInfo;
	IUnknown* punkMarshal;
} CHelloWorldData;

#define GetBaseObjectPtr(x,y,z)     (x *)(((char *)(void *)z)-(size_t)(char *)(&(((x*)NULL)->y)))

static STDMETHODIMP CHelloWorld_IUnknown_QueryInterface(IUnknown* pThis, REFIID riid, void** ppvObject)
{
	HRESULT hr = E_NOINTERFACE;
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IUnknown, pThis);

	if (IsEqualIID(riid, &IID_IDispatch) || IsEqualIID(riid, &IID_IHelloWorld))
	{
		IUnknown_AddRef(pData->lpOuter);

		*ppvObject = &(pData->IHelloWorld);

		hr = S_OK;
	}
	else
	{
		if (IsEqualIID(riid, &IID_IUnknown))
		{
			InterlockedIncrement(&pData->lUsage);

			*ppvObject = &(pData->IUnknown);

			hr = S_OK;
		}
		else
		{
			if (pData->punkMarshal && IsEqualIID(riid, &IID_IMarshal))
			{
				hr = IUnknown_QueryInterface(pData->punkMarshal, riid, ppvObject);
			}
			else
			{
				*ppvObject = NULL;
			}
		}
	}

	return hr;
}

static STDMETHODIMP_(ULONG) CHelloWorld_IUnknown_AddRef(IUnknown* pThis)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IUnknown, pThis);
	return InterlockedIncrement(&pData->lUsage);
}

static STDMETHODIMP_(ULONG) CHelloWorld_IUnknown_Release(IUnknown* pThis)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IUnknown, pThis);
	LONG res = InterlockedDecrement(&pData->lUsage);

	if (!res)
	{
		if (pData->piTypeInfo) IUnknown_Release(pData->piTypeInfo);
		if (pData->punkMarshal) IUnknown_Release(pData->punkMarshal);
		LocalFree(pData);
		InterlockedDecrement(&globalUsage);
	}

	return res;
}

static STDMETHODIMP CHelloWorld_IHelloWorld_QueryInterface(IHelloWorld* pThis, REFIID riid, void** ppvObject)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IHelloWorld, pThis);
	return IUnknown_QueryInterface(pData->lpOuter, riid, ppvObject);
}

static STDMETHODIMP_(ULONG) CHelloWorld_IHelloWorld_AddRef(IHelloWorld* pThis)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IHelloWorld, pThis);
	return IUnknown_AddRef(pData->lpOuter);
}

static STDMETHODIMP_(ULONG) CHelloWorld_IHelloWorld_Release(IHelloWorld* pThis)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IHelloWorld, pThis);
	return IUnknown_Release(pData->lpOuter);
}

static STDMETHODIMP CHelloWorld_IHelloWorld_GetTypeInfoCount(IHelloWorld* pThis, UINT* pctinfo)
{
	*pctinfo = 1;
	return S_OK;
}

static STDMETHODIMP CHelloWorld_IHelloWorld_GetTypeInfo(IHelloWorld* pThis, UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IHelloWorld, pThis);
	if (iTInfo) return DISP_E_BADINDEX;
	ITypeInfo_AddRef(pData->piTypeInfo);
	*ppTInfo = pData->piTypeInfo;

	return S_OK;
}

static STDMETHODIMP CHelloWorld_IHelloWorld_GetIDsOfNames(IHelloWorld* pThis, REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IHelloWorld, pThis);

	if (!IsEqualIID(riid, &IID_NULL))
	{
		return DISP_E_UNKNOWNINTERFACE;
	}

	return DispGetIDsOfNames(pData->piTypeInfo, rgszNames, cNames, rgDispId);
}

static STDMETHODIMP CHelloWorld_IHelloWorld_Invoke(IHelloWorld* pThis, DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IHelloWorld, pThis);

	if (!IsEqualIID(riid, &IID_NULL))
	{
		return DISP_E_UNKNOWNINTERFACE;
	}

	return DispInvoke(pThis, pData->piTypeInfo, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
}

static STDMETHODIMP CHelloWorld_IHelloWorld_GetMessage(IHelloWorld* pThis, int Hint, BSTR* lpMessage)
{
	*lpMessage = SysAllocString(Hint < 0 ? L"Goodbye, Cruel World" : L"Hello World");

	return S_OK;
}

static IUnknownVtbl CHelloWorld_IUnknownVtbl =
{
	CHelloWorld_IUnknown_QueryInterface,
	CHelloWorld_IUnknown_AddRef,
	CHelloWorld_IUnknown_Release
};

static IHelloWorldVtbl CHelloWorld_IHelloWorldVtbl =
{
	CHelloWorld_IHelloWorld_QueryInterface,
	CHelloWorld_IHelloWorld_AddRef,
	CHelloWorld_IHelloWorld_Release,
	CHelloWorld_IHelloWorld_GetTypeInfoCount,
	CHelloWorld_IHelloWorld_GetTypeInfo,
	CHelloWorld_IHelloWorld_GetIDsOfNames,
	CHelloWorld_IHelloWorld_Invoke,
	CHelloWorld_IHelloWorld_GetMessage
};

static STDMETHODIMP CClassObject_CHelloWorld_IClassFactory_QueryInterface(IClassFactory* pThis, REFIID riid, void** ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IClassFactory))
	{
		InterlockedIncrement(&globalUsage);

		*ppvObject = pThis;

		return S_OK;
	}

	return E_NOINTERFACE;
}

static STDMETHODIMP_(ULONG) CClassObject_CHelloWorld_IClassFactory_AddRef(IClassFactory* pThis)
{
	return InterlockedIncrement(&globalUsage);
}

static STDMETHODIMP_(ULONG) CClassObject_CHelloWorld_IClassFactory_Release(IClassFactory* pThis)
{
	return InterlockedDecrement(&globalUsage);
}

static STDMETHODIMP CClassObject_CHelloWorld_IClassFactory_CreateInstance(IClassFactory* pThis, LPVOID punk, REFIID riid, void** ppvObject)
{
	HRESULT hr = E_FAIL;

	if (punk == NULL || IsEqualIID(riid, &IID_IUnknown))
	{
		ITypeLib* piTypeLib = NULL;

		hr = LoadTypeLibEx(globalModuleFileName, REGKIND_NONE, &piTypeLib);

		if (SUCCEEDED(hr))
		{
			CHelloWorldData* pData = LocalAlloc(LMEM_ZEROINIT, sizeof(*pData));

			if (pData)
			{
				IUnknown* p = &(pData->IUnknown);
				InterlockedIncrement(&globalUsage);

				pData->IUnknown.lpVtbl = &CHelloWorld_IUnknownVtbl;
				pData->IHelloWorld.lpVtbl = &CHelloWorld_IHelloWorldVtbl;

				pData->lUsage = 1;
				pData->lpOuter = punk ? punk : p;

				hr = ITypeLib_GetTypeInfoOfGuid(piTypeLib, &IID_IHelloWorld, &pData->piTypeInfo);

				if (SUCCEEDED(hr))
				{
					if (punk)
					{
						hr = S_OK;

						*ppvObject = p;
					}
					else
					{
						hr = CoCreateFreeThreadedMarshaler(p, &pData->punkMarshal);

						if (SUCCEEDED(hr))
						{
							hr = IUnknown_QueryInterface(p, riid, ppvObject);
						}

						IUnknown_Release(p);
					}
				}
				else
				{
					IUnknown_Release(p);
				}
			}

			if (piTypeLib)
			{
				ITypeLib_Release(piTypeLib);
			}
		}
	}

	return hr;
}

static STDMETHODIMP CClassObject_CHelloWorld_IClassFactory_LockServer(IClassFactory* pThis, BOOL bLock)
{
	if (bLock)
	{
		InterlockedIncrement(&globalUsage);
	}
	else
	{
		InterlockedDecrement(&globalUsage);
	}

	return S_OK;
}

static IClassFactoryVtbl CClassObject_CHelloWorld_IClassFactoryVtbl =
{
	CClassObject_CHelloWorld_IClassFactory_QueryInterface,
	CClassObject_CHelloWorld_IClassFactory_AddRef,
	CClassObject_CHelloWorld_IClassFactory_Release,
	CClassObject_CHelloWorld_IClassFactory_CreateInstance,
	CClassObject_CHelloWorld_IClassFactory_LockServer
};

static struct CClassObject {
	IClassFactoryVtbl* lpVtbl;
} CClassObject_CHelloWorld = { &CClassObject_CHelloWorld_IClassFactoryVtbl };

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppvObject)
{
	if (IsEqualCLSID(rclsid, &CLSID_CHelloWorld))
	{
		return IUnknown_QueryInterface((IUnknown*)&CClassObject_CHelloWorld, riid, ppvObject);
	}

	return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllCanUnloadNow(void)
{
	return globalUsage ? S_FALSE : S_OK;
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
