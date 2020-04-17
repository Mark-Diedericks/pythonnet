using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Python.Runtime
{
    /// <summary>
    /// Provides the implementation for reflected interface types. Managed
    /// interfaces are represented in Python by actual Python type objects.
    /// Each of those type objects is associated with an instance of this
    /// class, which provides the implementation for the Python type.
    /// </summary>
    internal class InterfaceObject : ClassBase
    {
        internal ConstructorInfo ctor;

        internal InterfaceObject(Type tp) : base(tp)
        {
            var coclass = (CoClassAttribute)Attribute.GetCustomAttribute(tp, cc_attr);
            if (coclass != null)
            {
                ctor = coclass.CoClass.GetConstructor(Type.EmptyTypes);
            }
        }

        private static Type cc_attr;

        static InterfaceObject()
        {
            cc_attr = typeof(CoClassAttribute);
        }

        /// <summary>
        /// Implements __new__ for reflected interface types.
        /// </summary>
        public static IntPtr tp_new(IntPtr tp, IntPtr args, IntPtr kw)
        {
            var self = (InterfaceObject)GetManagedObject(tp);
            var nargs = Runtime.PyTuple_Size(args);
            Type type = self.type;
            object obj;

            if (nargs == 1)
            {
                IntPtr inst = Runtime.PyTuple_GetItem(args, 0);
                var co = GetManagedObject(inst) as CLRObject;

                if (co == null || !type.IsInstanceOfType(co.inst))
                {
                    Exceptions.SetError(Exceptions.TypeError, $"object does not implement {type.Name}");
                    return IntPtr.Zero;
                }

                obj = co.inst;
            }

            else if (nargs == 0 && self.ctor != null)
            {
                obj = self.ctor.Invoke(null);

                if (obj == null || !type.IsInstanceOfType(obj))
                {
                    Exceptions.SetError(Exceptions.TypeError, "CoClass default constructor failed");
                    return IntPtr.Zero;
                }
            }

            else
            {
                Exceptions.SetError(Exceptions.TypeError, "interface takes exactly one argument");
                return IntPtr.Zero;
            }

            return CLRObject.GetInstHandle(obj, self.pyHandle);
        }

        #region COM Binding

        public override IntPtr type_subscript(IntPtr idx)
        {
            // If this type is the Array type, the [<type>] means we need to
            // construct and return an array type of the given element type.
            if (type == typeof(Array))
            {
                if (Runtime.PyTuple_Check(idx))
                {
                    return Exceptions.RaiseTypeError("type expected");
                }
                var c = GetManagedObject(idx) as ClassBase;
                Type t = c != null ? c.type : Converter.GetTypeByAlias(idx);
                if (t == null)
                {
                    return Exceptions.RaiseTypeError("type expected");
                }
                Type a = t.MakeArrayType();
                ClassBase o = ClassManager.GetClass(a);
                Runtime.XIncref(o.pyHandle);
                return o.pyHandle;
            }

            // If there are generics in our namespace with the same base name
            // as the current type, then [<type>] means the caller wants to
            // bind the generic type matching the given type parameters.
            Type[] types = Runtime.PythonArgsToTypeArray(idx);
            if (types == null)
            {
                return Exceptions.RaiseTypeError("type(s) expected");
            }

            Type gtype = AssemblyManager.LookupType($"{type.FullName}`{types.Length}");
            if (gtype != null)
            {
                var g = ClassManager.GetClass(gtype) as GenericType;
                return g.type_subscript(idx);
                //Runtime.XIncref(g.pyHandle);
                //return g.pyHandle;
            }
            return Exceptions.RaiseTypeError("unsubscriptable object");
        }

        public static IntPtr mp_subscript(IntPtr ob, IntPtr idx)
        {
            //ManagedType self = GetManagedObject(ob);
            IntPtr tp = Runtime.PyObject_TYPE(ob);
            var cls = (ClassBase)GetManagedObject(tp);

            if (cls.indexer == null || !cls.indexer.CanGet)
            {
                Exceptions.SetError(Exceptions.TypeError, "unindexable object");
                return IntPtr.Zero;
            }

            // Arg may be a tuple in the case of an indexer with multiple
            // parameters. If so, use it directly, else make a new tuple
            // with the index arg (method binders expect arg tuples).
            IntPtr args = idx;
            var free = false;

            if (!Runtime.PyTuple_Check(idx))
            {
                args = Runtime.PyTuple_New(1);
                Runtime.XIncref(idx);
                Runtime.PyTuple_SetItem(args, 0, idx);
                free = true;
            }

            IntPtr value;

            try
            {
                value = cls.indexer.GetItem(ob, args);
            }
            finally
            {
                if (free)
                {
                    Runtime.XDecref(args);
                }
            }
            return value;
        }

        public static int mp_ass_subscript(IntPtr ob, IntPtr idx, IntPtr v)
        {
            //ManagedType self = GetManagedObject(ob);
            IntPtr tp = Runtime.PyObject_TYPE(ob);
            var cls = (ClassBase)GetManagedObject(tp);

            if (cls.indexer == null || !cls.indexer.CanSet)
            {
                Exceptions.SetError(Exceptions.TypeError, "object doesn't support item assignment");
                return -1;
            }

            // Arg may be a tuple in the case of an indexer with multiple
            // parameters. If so, use it directly, else make a new tuple
            // with the index arg (method binders expect arg tuples).
            IntPtr args = idx;
            var free = false;

            if (!Runtime.PyTuple_Check(idx))
            {
                args = Runtime.PyTuple_New(1);
                Runtime.XIncref(idx);
                Runtime.PyTuple_SetItem(args, 0, idx);
                free = true;
            }

            // Get the args passed in.
            var i = Runtime.PyTuple_Size(args);
            IntPtr defaultArgs = cls.indexer.GetDefaultArgs(args);
            var numOfDefaultArgs = Runtime.PyTuple_Size(defaultArgs);
            var temp = i + numOfDefaultArgs;
            IntPtr real = Runtime.PyTuple_New(temp + 1);
            for (var n = 0; n < i; n++)
            {
                IntPtr item = Runtime.PyTuple_GetItem(args, n);
                Runtime.XIncref(item);
                Runtime.PyTuple_SetItem(real, n, item);
            }

            // Add Default Args if needed
            for (var n = 0; n < numOfDefaultArgs; n++)
            {
                IntPtr item = Runtime.PyTuple_GetItem(defaultArgs, n);
                Runtime.XIncref(item);
                Runtime.PyTuple_SetItem(real, n + i, item);
            }
            // no longer need defaultArgs
            Runtime.XDecref(defaultArgs);
            i = temp;

            // Add value to argument list
            Runtime.XIncref(v);
            Runtime.PyTuple_SetItem(real, i, v);

            try
            {
                cls.indexer.SetItem(ob, real);
            }
            finally
            {
                Runtime.XDecref(real);

                if (free)
                {
                    Runtime.XDecref(args);
                }
            }

            if (Exceptions.ErrorOccurred())
            {
                return -1;
            }

            return 0;
        }

        #endregion
    }
}
