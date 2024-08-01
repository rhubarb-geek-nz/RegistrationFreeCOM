// Copyright (c) 2024 Roger Brown.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Xml;

namespace RhubarbGeekNzRegistrationFreeCOM
{
    sealed public class ModuleAssembly : IModuleAssemblyInitializer, IModuleAssemblyCleanup
    {
        private readonly static Dictionary<IntPtr, List<uint>> registrations = new Dictionary<IntPtr, List<uint>>();

        public void OnImport()
        {
            string xmlns = "urn:schemas-microsoft-com:asm.v1";
            Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");
            var GetHINSTANCE = typeof(Marshal).GetMethod("GetHINSTANCE");

            foreach (var loadedModule in GetType().Assembly.GetLoadedModules())
            {
                XmlDocument document = new XmlDocument();

                {
                    IntPtr hInstance = (IntPtr)GetHINSTANCE.Invoke(null, new object[] { loadedModule });
                    IntPtr hResource = FindResource(hInstance, (IntPtr)2, (IntPtr)24);
                    int dwSize = SizeofResource(hInstance, hResource);
                    IntPtr ptr = LoadResource(hInstance, hResource);
                    byte[] data = new byte[dwSize];
                    Marshal.Copy(ptr, data, 0, dwSize);
                    document.Load(new MemoryStream(data));
                }

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(new NameTable());
                nsmgr.AddNamespace("m", xmlns);

                foreach (XmlNode node in document.SelectNodes("/m:assembly/m:dependency/m:dependentAssembly/m:assemblyIdentity", nsmgr))
                {
                    XmlElement dependency = node as XmlElement;
                    string name = dependency.GetAttribute("name");
                    string version = dependency.GetAttribute("version");
                    string type = dependency.GetAttribute("type");

                    if ("win32".Equals(type) && RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    {
                        string moduleDir = Path.GetDirectoryName(loadedModule.FullyQualifiedName);
                        string manifestPath = Path.Combine(moduleDir, name + ".manifest");
                        XmlDocument doc = new XmlDocument();

                        using (var fs = File.OpenRead(manifestPath))
                        {
                            doc.Load(fs);
                        }

                        var identity = doc.SelectSingleNode("/m:assembly/m:assemblyIdentity", nsmgr) as XmlElement;

                        string identityName = identity.GetAttribute("name");
                        string identityVersion = identity.GetAttribute("version");

                        if (name.Equals(identityName) && version.Equals(identityVersion))
                        {
                            foreach (XmlNode file in doc.SelectNodes("/m:assembly/m:file", nsmgr))
                            {
                                string dllname = file.Attributes["name"].Value;
                                string arch = RuntimeInformation.ProcessArchitecture.ToString().ToLower();
                                string path = String.Join(Path.DirectorySeparatorChar.ToString(), new string[] { moduleDir, "win-" + arch, dllname });
                                IntPtr hModule = CoLoadLibrary(path, 0);

                                if (hModule == IntPtr.Zero)
                                {
                                    throw new Win32Exception(Marshal.GetLastWin32Error(), $"Failed to load {path}");
                                }

                                List<uint> registeredClasses = new List<uint>();
                                IntPtr intPtr = GetProcAddress(hModule, "DllGetClassObject");
                                DllGetClassObjectDelegate dllGetClassObjectDelegate = Marshal.GetDelegateForFunctionPointer(intPtr, typeof(DllGetClassObjectDelegate)) as DllGetClassObjectDelegate;

                                foreach (XmlNode classNode in file.SelectNodes("m:comClass", nsmgr))
                                {
                                    Guid clsid = new Guid(classNode.Attributes["clsid"].Value);
                                    object classObject;

                                    Marshal.ThrowExceptionForHR(dllGetClassObjectDelegate(clsid, IID_IUnknown, out classObject));

                                    uint dwRegisterClass;
                                    CoRegisterClassObject(ref clsid, classObject, CLSCTX.CLSCTX_INPROC_SERVER, REGCLS.REGCLS_MULTIPLEUSE, out dwRegisterClass);

                                    registeredClasses.Add(dwRegisterClass);
                                }

                                registrations.Add(hModule, registeredClasses);
                            }
                        }
                    }
                }
            }
        }

        public void OnRemove(PSModuleInfo psModuleInfo)
        {
            foreach (var kvp in registrations.ToArray())
            {
                IntPtr hModule = kvp.Key;
                var list = kvp.Value;

                foreach (var dwRegisterClass in list.ToArray())
                {
                    list.Remove(dwRegisterClass);
                    CoRevokeClassObject(dwRegisterClass);
                }

                if (list.Count == 0)
                {
                    registrations.Remove(hModule);
                    IntPtr intPtr = GetProcAddress(hModule, "DllCanUnloadNow");
                    DllCanUnloadNowDelegate dllCanUnloadNowDelegate = Marshal.GetDelegateForFunctionPointer(intPtr, typeof(DllCanUnloadNowDelegate)) as DllCanUnloadNowDelegate;
                    if (0 == dllCanUnloadNowDelegate())
                    {
                        CoFreeLibrary(hModule);
                    }
                }
            }
        }

        [DllImport("ole32.dll", PreserveSig = false)]
        private static extern void CoRegisterClassObject([In] ref Guid rclsid, [MarshalAs(UnmanagedType.IUnknown)] object pUnk, CLSCTX dwClsContext, REGCLS flags, out uint lpdwRegister);

        [DllImport("ole32.dll", PreserveSig = false)]
        private static extern void CoRevokeClassObject(uint dwRegister);

        [Flags]
        enum CLSCTX : uint { CLSCTX_INPROC_SERVER = 1 }

        [Flags]
        enum REGCLS : uint { REGCLS_MULTIPLEUSE = 1 }

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        public delegate int DllGetClassObjectDelegate(
            [MarshalAs(UnmanagedType.LPStruct)]
            Guid rclsid,
            [MarshalAs(UnmanagedType.LPStruct)]
            Guid riid,
            [MarshalAs(UnmanagedType.IUnknown, IidParameterIndex=1)]
            out object ppv
        );

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        public delegate int DllCanUnloadNowDelegate();

        [DllImport("ole32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr CoLoadLibrary(string lpszLibName, uint dwFlags);

        [DllImport("ole32.dll", SetLastError = true)]
        private static extern int CoFreeLibrary(IntPtr hModule);

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
        private static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

        [DllImport("kernel32.dll")]
        static extern IntPtr FindResource(IntPtr hModule, IntPtr lpType, IntPtr lpName);

        [DllImport("kernel32.dll")]
        static extern int SizeofResource(IntPtr hModule, IntPtr hResource);

        [DllImport("kernel32.dll")]
        static extern IntPtr LoadResource(IntPtr hModule, IntPtr hResource);
    }

    [Cmdlet(VerbsCommon.New, "RegistrationFreeCOM")]
    [OutputType(typeof(IHelloWorld))]
    sealed public class NewRegistrationFreeCOM : PSCmdlet
    {
        protected override void ProcessRecord()
        {
            WriteObject(new CHelloWorld());
        }
    }

    [Cmdlet(VerbsDiagnostic.Test, "RegistrationFreeCOM")]
    [OutputType(typeof(string))]
    sealed public class TestRegistrationFreeCOM : PSCmdlet
    {
        [Parameter(Position = 0)]
        public int Hint;

        protected override void ProcessRecord()
        {
            WriteObject(new CHelloWorld().GetMessage(Hint));
        }
    }
}
