/***********************************
 * Copyright (c) 2024 Roger Brown.
 * Licensed under the MIT License.
 ****/

#include <objbase.h>
#include <stdio.h>

int main(int argc, char** argv)
{
	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

	if (hr >= 0)
	{
		CLSID clsid;

		BSTR app = SysAllocString(L"RhubarbGeekNz.HelloWorld");

		hr = CLSIDFromProgID(app, &clsid);

		SysFreeString(app);

		if (hr >= 0)
		{
			IDispatch* dispatch = NULL;

			hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void**)&dispatch);

			if (hr >= 0)
			{
				LPOLESTR names[] = {
					SysAllocString(L"GetMessage")
				};
				DISPID dispId[sizeof(names) / sizeof(names[0])];

				hr = dispatch->GetIDsOfNames(IID_NULL, names, sizeof(names) / sizeof(names[0]), LOCALE_USER_DEFAULT, dispId);

				int i = sizeof(names) / sizeof(names[0]);

				while (i--)
				{
					SysFreeString(names[i]);
				}

				if (hr >= 0)
				{
					WORD flags = DISPATCH_METHOD;
					DISPPARAMS params;
					VARIANT result;
					EXCEPINFO ex;
					UINT argErr = 0;
					VARIANTARG args[1];

					ZeroMemory(&ex, sizeof(ex));
					ZeroMemory(&params, sizeof(params));

					VariantInit(args);
					args[0].vt = VT_I4;
					args[0].intVal = 1;

					params.cArgs = 1;
					params.rgvarg = args;

					VariantInit(&result);

					hr = dispatch->Invoke(dispId[0], IID_NULL, LOCALE_USER_DEFAULT, flags, &params, &result, &ex, &argErr);

					if (hr >= 0 && result.vt == VT_BSTR)
					{
						printf("%S\n", result.bstrVal);
					}

					VariantClear(&result);
				}

				dispatch->Release();
			}
		}

		CoUninitialize();
	}

	if (hr < 0)
	{
		fprintf(stderr, "0x%lx\n", (long)hr);
	}

	return hr < 0;
}