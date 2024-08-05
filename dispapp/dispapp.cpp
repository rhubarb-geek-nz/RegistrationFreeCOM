/***********************************
 * Copyright (c) 2024 Roger Brown.
 * Licensed under the MIT License.
 ****/

#include<windows.h>
#undef GetMessage
#include <displib.h>
#include <stdio.h>

int main(int argc, char** argv)
{
	const char* op = "CoInitializeEx";
	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

	if (SUCCEEDED(hr))
	{
		op = "CLSIDFromProgID";
		BSTR app = SysAllocString(L"RhubarbGeekNz.RegistrationFreeCOM");
		CLSID clsid;

		hr = CLSIDFromProgID(app, &clsid);

		SysFreeString(app);

		if (SUCCEEDED(hr))
		{
			op = "CoCreateInstance";
			IHelloWorld* helloWorld = NULL;

			hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IHelloWorld, (void**)&helloWorld);

			if (SUCCEEDED(hr))
			{
				op = "GetMessage";
				int hint = 1;

				while (hint < 6)
				{
					BSTR bstr = NULL;

					hr = helloWorld->GetMessage(hint, &bstr);

					if (SUCCEEDED(hr))
					{
						printf("%d %S\n", hint, bstr);

						SysFreeString(bstr);
					}

					hint++;
				}

				helloWorld->Release();
			}
		}

		CoUninitialize();
	}

	if (FAILED(hr))
	{
		fprintf(stderr, "%s 0x%lx\n", op, (long)hr);
	}

	return hr < 0;
}
