/***********************************
 * Copyright (c) 2024 Roger Brown.
 * Licensed under the MIT License.
 ****/

using RhubarbGeekNz.RegistrationFreeCOM;
using System;
using System.Runtime.InteropServices;

namespace dispnet
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Guid IID_IUnknown = Guid.Parse("00000000-0000-0000-C000-000000000046");
            Guid CLSID_CHelloWorld = Guid.Parse("49ef0168-2765-4932-be4c-e21e0d7a554f");
            object classObject;

            DllGetClassObject(ref CLSID_CHelloWorld, ref IID_IUnknown, out classObject);

            uint dwRegister;

            CoRegisterClassObject(ref CLSID_CHelloWorld, classObject, CLSCTX.CLSCTX_INPROC_SERVER, REGCLS.REGCLS_MULTIPLEUSE, out dwRegister);

            Type type = Type.GetTypeFromCLSID(CLSID_CHelloWorld);

            IHelloWorld obj = Activator.CreateInstance(type) as IHelloWorld;

            var result = obj.GetMessage(1);

            Console.WriteLine($"{result}");

            CoRevokeClassObject(dwRegister);
        }

        [DllImport("displib.dll", PreserveSig = false)]
        static extern void DllGetClassObject([In] ref Guid rclsid,[In] ref Guid riid,[MarshalAs(UnmanagedType.Interface)][Out] out object ppv);

        [DllImport("ole32.dll", PreserveSig = false)]
        public static extern void CoRegisterClassObject([In] ref Guid rclsid, [MarshalAs(UnmanagedType.IUnknown)] object pUnk, CLSCTX dwClsContext, REGCLS flags, out uint lpdwRegister);

        [DllImport("ole32.dll", PreserveSig = false)]
        public static extern void CoRevokeClassObject(uint dwRegister);
            
        [Flags]
        public enum CLSCTX : uint { CLSCTX_INPROC_SERVER = 1 }

        [Flags]
        public enum REGCLS : uint { REGCLS_MULTIPLEUSE = 1 }
    }
}
