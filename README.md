# rhubarb-geek-nz/RegistrationFreeCOM

Demonstration of Registration-Free COM object.

[displib.idl](displib/displib.idl) defines the dual-interface for a simple inprocess server.

[displib.c](displib/displib.c) implements the interface.

[dispapp.cpp](dispapp/dispapp.cpp) creates an instance with [CoCreateInstance](https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cocreateinstance) and uses it to get a message to display.

[dispapp.manifest](dispapp/dispapp.manifest) provides the side-by-side assembly binding.

[disptlb.nuspec](disptlb/disptlb.nuspec) is used for `NuGet` packaging of the dlls.

[dispnet.cs](dispnet/dispnet.cs) demonstrates using [System.Activator.CreateInstance](https://learn.microsoft.com/en-us/dotnet/api/system.activator.createinstance) to create the instance.

[package.ps1](package.ps1) is used to automate the building of multiple architectures and create the `NuGet` package.

[dispps1.cs](dispps1/dispps1.cs) demonstration of loading COM object into a PowerShell cmdlet.
