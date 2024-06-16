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
	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

	if (hr >= 0)
	{
		BSTR app = SysAllocString(L"RhubarbGeekNz.HelloWorld");
		CLSID clsid;

		hr = CLSIDFromProgID(app, &clsid);

		SysFreeString(app);

		if (hr >= 0)
		{
			IHelloWorld* helloWorld = NULL;

			hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IHelloWorld, (void**)&helloWorld);

			if (hr >= 0)
			{
				BSTR bstr = NULL;

				hr = helloWorld->GetMessage(1, &bstr);

				if (hr >= 0)
				{
					printf("%S\n", bstr);

					SysFreeString(bstr);
				}

				helloWorld->Release();
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
