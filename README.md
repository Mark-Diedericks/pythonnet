# pythonnet
Custom build of pythonnet 2.4.0 - Usage of COM types through managed wrappers.
Original project at http://pythonnet.github.io  and  https://github.com/pythonnet/pythonnet

Very basic changes to original code to allow for embedded PythonNET to use ComObjects given that an Interop Assembly exists and is loaded.

Code will simply attempt to find the managed type which represents the given ComObject and use that to define functionality.

Additionally, InterfaceObjects are now indexable as that did not exist in the original PythonNet project.

In my case this was needed for the Microsoft.Office.Interop.Excel.Range Com Interface, as it is enumerable.
