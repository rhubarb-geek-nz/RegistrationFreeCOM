// Copyright (c) 2024 Roger Brown.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Management.Automation;
using System.Runtime.InteropServices;

namespace RhubarbGeekNzRegistrationFreeCOM
{
    public class displib : IModuleAssemblyInitializer, IModuleAssemblyCleanup
    {
        private static uint dwRegisterClass;
        private static bool isRegistered;
        private static IntPtr hModule = IntPtr.Zero;
        private static readonly Dictionary<Architecture, string> archDirectories = new Dictionary<Architecture, string>(){
            {Architecture.Arm, "win-arm"},
            {Architecture.Arm64, "win-arm64"},
            {Architecture.X86, "win-x86"},
            {Architecture.X64, "win-x64"}
        };

        public void OnImport()
        {
            if (hModule == IntPtr.Zero)
            {
                string archDir;

                if (archDirectories.TryGetValue(RuntimeInformation.ProcessArchitecture, out archDir))
                {
                    string dllName = GetType().Assembly.Location;
                    string dirName = Path.GetDirectoryName(dllName);
                    string path = String.Join(Path.DirectorySeparatorChar.ToString(), new string[] { dirName, archDir, "displib.dll" });

                    hModule = CoLoadLibrary(path, 0);

                    if (hModule == IntPtr.Zero)
                    {
                        throw new Win32Exception(Marshal.GetLastWin32Error(), $"Failed to load {path}");
                    }
                }
            }

            if (!isRegistered && hModule != IntPtr.Zero)
            {
                IntPtr intPtr = GetProcAddress(hModule, "DllGetClassObject");

                if (intPtr != IntPtr.Zero)
                {
                    Guid CLSID_CHelloWorld = new Guid("49ef0168-2765-4932-be4c-e21e0d7a554f");
                    Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");
                    object classObject;

                    DllGetClassObjectDelegate dllGetClassObjectDelegate = Marshal.GetDelegateForFunctionPointer(intPtr, typeof(DllGetClassObjectDelegate)) as DllGetClassObjectDelegate;

                    Marshal.ThrowExceptionForHR(dllGetClassObjectDelegate(CLSID_CHelloWorld, IID_IUnknown, out classObject));

                    CoRegisterClassObject(ref CLSID_CHelloWorld, classObject, CLSCTX.CLSCTX_INPROC_SERVER, REGCLS.REGCLS_MULTIPLEUSE, out dwRegisterClass);

                    isRegistered = true;
                }
            }
        }

        public void OnRemove(PSModuleInfo psModuleInfo)
        {
            if (isRegistered)
            {
                isRegistered = false;

                CoRevokeClassObject(dwRegisterClass);
            }

            IntPtr hInstance = hModule;
            hModule = IntPtr.Zero;

            if (hInstance != IntPtr.Zero)
            {
                CoFreeLibrary(hInstance);
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

        [DllImport("ole32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr CoLoadLibrary(string lpszLibName, uint dwFlags);

        [DllImport("ole32.dll", SetLastError = true)]
        private static extern int CoFreeLibrary(IntPtr hModule);

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
        private static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);
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
