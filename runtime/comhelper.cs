using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.CustomMarshalers;
using System.Text;

namespace Python.Runtime
{
    class COMHelper
    {
        internal static IntPtr GetInstHandle(object ob, Type t)
        {
            if (!Marshal.IsComObject(ob))
            {
                ClassBase cc = ClassManager.GetClass(t);
                CLRObject co = CLRObject.GetInstance(ob, cc.tpHandle);
                return co.pyHandle;
            }

            Type type = GetManagedType(ob, t);

            bool isEnumerable = false;
            bool isNotifyProperty = false;
            foreach (Type intf in type.GetInterfaces())
            {
                if (intf.Name == "IEnumerable" || intf.Name == "IList")
                    isEnumerable = true;

                if (intf.Name == "INotifyPropertyChanged")
                    isNotifyProperty = true;
            }


            System.Diagnostics.Debug.WriteLine(type.Name + " Enumerable: " + isEnumerable);


            if (isEnumerable && !isNotifyProperty)
            {
            }

            {
                ClassBase cc = ClassManager.GetClass(type);
                CLRObject co = CLRObject.GetInstance(ob, cc.tpHandle);
                return co.pyHandle;
            }
        }

        internal static Type GetManagedType(object obj, Type type)
        {
            if (!Marshal.IsComObject(obj))
                return type;

            if (!DispatchUtility.ImplementsIDispatch(obj))
                return type;

            try
            {
                Type manType = DispatchUtility.GetType(obj, true);
                System.Diagnostics.Debug.WriteLine("Got managed type " + type.FullName + " from ComObject.");

                return manType;
            }
            catch (Exception ex) { }

            return type;
        }
    }

    #region Bill Menees' DispatchUtility
    /// <summary>
    /// Authored by Bill Menees
    /// https://stackoverflow.com/a/14208030/8520655
    /// </summary>
    public static class DispatchUtility
    {
        private const int S_OK = 0; //From WinError.h
        private const int LOCALE_SYSTEM_DEFAULT = 2 << 10; //From WinNT.h == 2048 == 0x800

        public static bool ImplementsIDispatch(object obj)
        {
            bool result = obj is IDispatchInfo;
            return result;
        }

        public static Type GetType(object obj, bool throwIfNotFound)
        {
            RequireReference(obj, "obj");
            Type result = GetType((IDispatchInfo)obj, throwIfNotFound);
            return result;
        }

        public static bool TryGetDispId(object obj, string name, out int dispId)
        {
            RequireReference(obj, "obj");
            bool result = TryGetDispId((IDispatchInfo)obj, name, out dispId);
            return result;
        }

        public static object Invoke(object obj, int dispId, object[] args)
        {
            string memberName = "[DispId=" + dispId + "]";
            object result = Invoke(obj, memberName, args);
            return result;
        }

        public static object Invoke(object obj, string memberName, object[] args)
        {
            RequireReference(obj, "obj");
            Type type = obj.GetType();
            object result = type.InvokeMember(memberName,
                BindingFlags.InvokeMethod | BindingFlags.GetProperty,
                null, obj, args, null);
            return result;
        }

        private static void RequireReference<T>(T value, string name) where T : class
        {
            if (value == null)
            {
                throw new ArgumentNullException(name);
            }
        }

        private static Type GetType(IDispatchInfo dispatch, bool throwIfNotFound)
        {
            RequireReference(dispatch, "dispatch");

            Type result = null;
            int typeInfoCount;
            int hr = dispatch.GetTypeInfoCount(out typeInfoCount);
            if (hr == S_OK && typeInfoCount > 0)
            {
                result = dispatch.GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT);
            }

            if (result == null && throwIfNotFound)
            {
                // If the GetTypeInfoCount called failed, throw an exception for that.
                Marshal.ThrowExceptionForHR(hr);

                // Otherwise, throw the same exception that Type.GetType would throw.
                throw new TypeLoadException();
            }

            return result;
        }

        private static bool TryGetDispId(IDispatchInfo dispatch, string name, out int dispId)
        {
            RequireReference(dispatch, "dispatch");
            RequireReference(name, "name");

            bool result = false;

            Guid iidNull = Guid.Empty;
            int hr = dispatch.GetDispId(ref iidNull, ref name, 1, LOCALE_SYSTEM_DEFAULT, out dispId);

            const int DISP_E_UNKNOWNNAME = unchecked((int)0x80020006); //From WinError.h
            const int DISPID_UNKNOWN = -1; //From OAIdl.idl
            if (hr == S_OK)
            {
                result = true;
            }
            else if (hr == DISP_E_UNKNOWNNAME && dispId == DISPID_UNKNOWN)
            {
                result = false;
            }
            else
            {
                Marshal.ThrowExceptionForHR(hr);
            }

            return result;
        }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("00020400-0000-0000-C000-000000000046")]
        private interface IDispatchInfo
        {
            int GetTypeInfoCount(out int typeInfoCount);


            [return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalTypeRef = typeof(TypeToTypeInfoMarshaler))]
            Type GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);

            int GetDispId(ref Guid riid, ref string name, int nameCount, int lcid, out int dispId);
        }
    }

    #endregion
}
