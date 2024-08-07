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
using System.Xml.Serialization;

namespace RhubarbGeekNzRegistrationFreeCOM
{
    [XmlRoot("assembly", Namespace = "urn:schemas-microsoft-com:asm.v1")]
    public class XmlAssembly
    {
        [XmlElement("file")]
        public XmlAssemblyFile[] Files;
    }

    public class XmlAssemblyFile
    {
        [XmlAttribute("name")]
        public string Name { get; set; }
        [XmlElement("comClass")]
        public XmlAssemblyComClass[] ComClasses;
    }

    public class XmlAssemblyComClass
    {
        [XmlAttribute("clsid")]
        public string Clsid;
    }

    sealed public class ModuleAssembly : IModuleAssemblyInitializer, IModuleAssemblyCleanup
    {
        private readonly static Dictionary<IntPtr, List<uint>> registrations = new Dictionary<IntPtr, List<uint>>();

        public void OnImport()
        {
            var GetHINSTANCE = typeof(Marshal).GetMethod("GetHINSTANCE");

            if (GetHINSTANCE != null)
            {
                Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

                foreach (var loadedModule in GetType().Assembly.GetLoadedModules())
                {
                    string moduleDir = Path.GetDirectoryName(loadedModule.FullyQualifiedName);
                    XmlAssembly manifest;

                    {
                        IntPtr hInstance = (IntPtr)GetHINSTANCE.Invoke(null, new object[] { loadedModule });
                        IntPtr hResource = FindResource(hInstance, (IntPtr)2, (IntPtr)24);
                        int dwSize = SizeofResource(hInstance, hResource);
                        IntPtr ptr = LoadResource(hInstance, hResource);
                        byte[] data = new byte[dwSize];
                        Marshal.Copy(ptr, data, 0, dwSize);

                        XmlSerializer xmlSerializer = new XmlSerializer(typeof(XmlAssembly));
                        manifest = (XmlAssembly)xmlSerializer.Deserialize(new MemoryStream(data));
                    }

                    foreach (var file in manifest.Files)
                    {
                        string arch = RuntimeInformation.ProcessArchitecture.ToString().ToLower();
                        string path = String.Join(Path.DirectorySeparatorChar.ToString(), new string[] { moduleDir, "win-" + arch, file.Name });
                        IntPtr hModule = LoadLibraryW(path);

                        if (hModule == IntPtr.Zero)
                        {
                            throw new Win32Exception(Marshal.GetLastWin32Error(), $"Failed to load {path}");
                        }

                        List<uint> registeredClasses;

                        if (registrations.TryGetValue(hModule, out registeredClasses))
                        {
                            FreeLibrary(hModule);
                        }
                        else
                        {
                            registeredClasses = new List<uint>();
                            registrations.Add(hModule, registeredClasses);
                        }

                        IntPtr intPtr = GetProcAddress(hModule, "DllGetClassObject");
                        DllGetClassObjectDelegate dllGetClassObjectDelegate = Marshal.GetDelegateForFunctionPointer(intPtr, typeof(DllGetClassObjectDelegate)) as DllGetClassObjectDelegate;

                        foreach (var comClass in file.ComClasses)
                        {
                            Guid clsid = new Guid(comClass.Clsid);
                            object classObject;

                            Marshal.ThrowExceptionForHR(dllGetClassObjectDelegate(clsid, IID_IUnknown, out classObject));

                            uint dwRegisterClass;
                            CoRegisterClassObject(ref clsid, classObject, CLSCTX.CLSCTX_INPROC_SERVER, REGCLS.REGCLS_MULTIPLEUSE, out dwRegisterClass);

                            registeredClasses.Add(dwRegisterClass);
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
                    IntPtr intPtr = GetProcAddress(hModule, "DllCanUnloadNow");
                    DllCanUnloadNowDelegate dllCanUnloadNowDelegate = Marshal.GetDelegateForFunctionPointer(intPtr, typeof(DllCanUnloadNowDelegate)) as DllCanUnloadNowDelegate;
                    if (0 == dllCanUnloadNowDelegate())
                    {
                        registrations.Remove(hModule);
                        FreeLibrary(hModule);
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

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr LoadLibraryW(string lpszLibName);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern int FreeLibrary(IntPtr hModule);

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
