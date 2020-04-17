using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

using COM = System.Runtime.InteropServices.ComTypes;

namespace Python.Runtime
{

    internal class COMHelper
    {
        private static Dictionary<Guid, Type> s_Cache;

        static COMHelper()
        {
            s_Cache = new Dictionary<Guid, Type>();
        }


        internal static Type GetManagedType(object ob, Type t)
        {
            if (ob == null)
                return t;

            if (!Marshal.IsComObject(ob))
                return t;

            IDispatch disp = GetIDispatch(ob);

            if (disp == null)
                return t;

            ITypeInfo ti = GetTypeInfo(disp);

            if (ti == null)
                return t;

            Type managedType = GetCOMObjectType(ti);

            if (managedType == null)
                return t;

            return managedType;
        }

        #region Get TypeInfo/IDispatch/Type

        // Modified code from https://stackoverflow.com/a/10883851/8520655

        internal static Func<ITypeInfo, Guid> GetTypeInfoGuid = (Func<ITypeInfo, Guid>)Delegate.CreateDelegate(typeof(Func<ITypeInfo, Guid>), typeof(Marshal).GetMethod("GetTypeInfoGuid", BindingFlags.NonPublic | BindingFlags.Static, null, new[] { typeof(ITypeInfo) }, null), true);
        public static Type GetCOMObjectType(ITypeInfo info)
        {
            Guid guid = GetTypeInfoGuid(info);
            Type result = null;

            //If we've found it previously, no need to look again
            #region Check Type Cache

            s_Cache.TryGetValue(guid, out result);
            if (result != null)
                return result;

            #endregion

            //Works for most of the cases
            #region Check AssemblyManager Assemblies

            AssemblyName[] names = AssemblyManager.ListAssemblies();
            foreach (AssemblyName an in names)
            {
                Assembly a = Assembly.Load(an);
                Type[] aTypes = a.GetTypes();
                foreach (Type t in aTypes)
                {
                    if (t.IsInterface && t.IsImport && t.GUID == guid && !t.Name.StartsWith("_"))
                    {
                        result = t;

                        List<Type> possible = aTypes.Where(x => x.Name.Equals(t.Name + "Class", StringComparison.Ordinal)).ToList<Type>();
                        if (possible.Count > 0)
                        {
                            s_Cache.Add(guid, result);
                            return result;
                        }
                    }
                }
            }

            if (result != null)
            {
                s_Cache.Add(guid, result);
                return result;
            }

            #endregion
            
            //Extended check of assemblies
            #region Check CurrentDomain Assemblies

            Assembly[] assemblies = AppDomain.CurrentDomain.GetAssemblies();
            foreach (Assembly a in assemblies)
            {
                Type[] aTypes = a.GetTypes();
                foreach (Type t in aTypes)
                {
                    if (t.IsInterface && t.IsImport && t.GUID == guid && !t.Name.StartsWith("_"))
                    {
                        result = t;

                        List<Type> possible = aTypes.Where(x => x.Name.Equals(t.Name + "Class", StringComparison.Ordinal)).ToList<Type>();
                        if (possible.Count > 0)
                        {
                            s_Cache.Add(guid, result);
                            return result;
                        }
                    }
                }
            }

            if (result != null)
            {
                s_Cache.Add(guid, result);
                return result;
            }

            #endregion

            //Almost never works for my cases but worth a shot i guess
            #region Check Type CLSID

            result = Type.GetTypeFromCLSID(guid);

            if (result != null)
            {
                s_Cache.Add(guid, result);
                return result;
            }

            #endregion

            //Reliable but incredibly slow, like really really really slow
            #region Get Type for ITypeInfo

            IntPtr pInfo = Marshal.GetIUnknownForObject(info);
            result =  Marshal.GetTypeForITypeInfo(pInfo);

            s_Cache.Add(guid, result);
            return result;

            #endregion
        }

        private static ITypeInfo GetTypeInfo(IDispatch disp)
        {
            const int LOCALE_SYSTEM_DEFAULT = 2 << 10; //From WinNT.h == 2048 == 0x800
            return disp.GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT);
        }

        private static IDispatch GetIDispatch(object ob)
        {
            try
            {
                return (IDispatch)ob;
            }
            catch (InvalidCastException ex)
            {
                return null;
            }
        }

        #endregion

        #region IDispatch

        // Modified code from https://stackoverflow.com/a/10883851/8520655

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("00020400-0000-0000-C000-000000000046")]
        private interface IDispatch
        {
            int GetTypeInfoCount();

            [return: MarshalAs(UnmanagedType.Interface)]
            ITypeInfo GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);

            void GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
                [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        }

        #endregion
    }
}
