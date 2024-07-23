// Copyright (c) 2024 Roger Brown.
// Licensed under the MIT License.

using System;
using System.Management.Automation;
using System.Runtime.InteropServices;

namespace RhubarbGeekNzRegistrationFreeCOM
{
    public class displib : IModuleAssemblyInitializer, IModuleAssemblyCleanup
    {
        internal static Guid CLSID_CHelloWorld = Guid.Parse("49ef0168-2765-4932-be4c-e21e0d7a554f");
        private static Guid IID_IUnknown = Guid.Parse("00000000-0000-0000-C000-000000000046");
        private static uint dwRegisterClass;
        private static bool isRegistered;

        public void OnImport()
        {
            object classObject;

            DllGetClassObject(ref CLSID_CHelloWorld, ref IID_IUnknown, out classObject);

            CoRegisterClassObject(ref CLSID_CHelloWorld, classObject, CLSCTX.CLSCTX_INPROC_SERVER, REGCLS.REGCLS_MULTIPLEUSE, out dwRegisterClass);

            isRegistered = true;
        }

        public void OnRemove(PSModuleInfo psModuleInfo)
        {
            if (isRegistered)
            {
                isRegistered = false;

                CoRevokeClassObject(dwRegisterClass);
            }
        }

        [DllImport("displib", PreserveSig = false)]
        static extern void DllGetClassObject([In] ref Guid rclsid, [In] ref Guid riid, [MarshalAs(UnmanagedType.Interface)][Out] out object ppv);

        [DllImport("ole32.dll", PreserveSig = false)]
        static extern void CoRegisterClassObject([In] ref Guid rclsid, [MarshalAs(UnmanagedType.IUnknown)] object pUnk, CLSCTX dwClsContext, REGCLS flags, out uint lpdwRegister);

        [DllImport("ole32.dll", PreserveSig = false)]
        static extern void CoRevokeClassObject(uint dwRegister);

        [Flags]
        enum CLSCTX : uint { CLSCTX_INPROC_SERVER = 1 }

        [Flags]
        enum REGCLS : uint { REGCLS_MULTIPLEUSE = 1 }
    }

    [Cmdlet(VerbsCommon.New, "RegistrationFreeCOM")]
    [OutputType(typeof(object))]
    sealed public class NewRegistrationFreeCOM : PSCmdlet
    {
        protected override void ProcessRecord()
        {
            WriteObject(Activator.CreateInstance(Type.GetTypeFromCLSID(displib.CLSID_CHelloWorld)));
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
            WriteObject((Activator.CreateInstance(Type.GetTypeFromCLSID(displib.CLSID_CHelloWorld)) as IHelloWorld).GetMessage(Hint));
        }
    }
}
