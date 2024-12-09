# Spire.Xls and Spire.Doc - Sequential Document Generation Failure

This is an MRE for a bug in Spire.Xls 13.9.0 & 14.7.3 / Spire.Doc 12.71. (or some combination thereof). In some circumstances, generating a .doc file followed by a .xls file will result in a crash. In our testing, we found that generating a Word file, then converting _two or more_ Excel files to a pdf, then generating another Word document results in an unrecoverable crash. The trick also seems to be that both file kinds need to have images included in them.

This bug is present when using either of Spire.Xls 13.9.0 or 14.7.3.

```
Unhandled Exception: System.InvalidCastException: Arg_InvalidCastException
   at System.Runtime.TypeCast.CheckCastClass(MethodTable*, Object) + 0x2b
   at SkiaSharp.SKAbstractManagedStream.ReadInternal(IntPtr s, Void* context, Void* buffer, IntPtr size) + 0xcc
   at Spire.Doc.Base!<BaseAddress>+0x29119d0
```

To run:

```
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
chmod u+x crash.py
./crash.py
```

This code was run on Ubuntu 22.04.1 LTS with Python 3.9.16.

## With Docker

This repo also contains a Dockerfile to illustrate the crash. To run it,

```
docker build -t crash .
docker run crash
```

In this case the result is the same:

```
Unhandled Exception: System.InvalidCastException: Arg_InvalidCastException
   at System.Runtime.TypeCast.CheckCastClass(MethodTable*, Object) + 0x2b
   at SkiaSharp.SKAbstractManagedStream.ReadInternal(IntPtr s, Void* context, Void* buffer, IntPtr size) + 0xcc
   at Spire.Doc.Base!<BaseAddress>+0x29119d0
Aborted (core dumped)
```

## Bug with Spire.Xls 14.9.3

A different bug is present with the same code when using Spire.Xls 14.9.3. In this case, the error is

```
Unhandled Exception: System.TypeInitializationException: TypeInitialization_Type_NoTypeAvailable
 ---> System.InvalidOperationException: The version of the native libSkiaSharp library (88.1) is incompatible with this version of SkiaSharp. Supported versions of the native libSkiaSharp library are in the range [80.3, 81.0).
   at SkiaSharp.SkiaSharpVersion.CheckNativeLibraryCompatible(Version, Version, Boolean) + 0x271
   at SkiaSharp.SKObject..cctor() + 0x25
   at System.Runtime.CompilerServices.ClassConstructorRunner.EnsureClassConstructorRun(StaticClassConstructionContext*) + 0xb9
   Exception_EndOfInnerExceptionStack
   at System.Runtime.CompilerServices.ClassConstructorRunner.EnsureClassConstructorRun(StaticClassConstructionContext*) + 0xaf
   at System.Runtime.CompilerServices.ClassConstructorRunner.CheckStaticClassConstructionReturnNonGCStaticBase(StaticClassConstructionContext*, IntPtr) + 0x9
   at SkiaSharp.SKObject.DeregisterHandle(IntPtr, SKObject) + 0x14
   at SkiaSharp.SKObject.set_Handle(IntPtr) + 0x24
   at SkiaSharp.SKNativeObject.Finalize() + 0x10
   at System.Runtime.__Finalizer.DrainQueue() + 0x66
   at System.Runtime.__Finalizer.ProcessFinalizers() + 0x3e
[1]    13175 IOT instruction (core dumped)  ./crash_change.py
```

## Bug with Spire.Xls 14.9.3 and Spire.Doc 12.12.0

There is a new bug present when using Spre.Xls 14.9.3 and Spire.Doc 12.12.0 together.

```
Error generating document: Arg_NullReferenceException:   at Spire.Xls.AOT.NLWorkbook.Workbook_get_Worksheets(IntPtr, IntPtr) + 0x47
Traceback (most recent call last):
  File "/home/eric/documents/spire-failure-doc-then-excel-mwe/./crash_change.py", line 40, in simulated_generate_documents
    f()
  File "/home/eric/documents/spire-failure-doc-then-excel-mwe/./crash_change.py", line 71, in generate_xlsx_docs
    sheet = wb.Worksheets[0]
  File "/home/eric/documents/spire-failure-doc-then-excel-mwe/venv/lib/python3.9/site-packages/spire/xls/Workbook.py", line 2081, in Worksheets
    intPtr = CallCFunction(GetDllLibXls().Workbook_get_Worksheets, self.Ptr)
  File "/home/eric/documents/spire-failure-doc-then-excel-mwe/venv/lib/python3.9/site-packages/spire/xls/common/__init__.py", line 113, in CallCFunction
    raise SpireException(info)
spire.xls.common.SpireException: Arg_NullReferenceException:   at Spire.Xls.AOT.NLWorkbook.Workbook_get_Worksheets(IntPtr, IntPtr) + 0x47
```

This one seems to occur roughly 3 in every 5 (6 / 20) times running the program.

| Run | Crash? |
| --- | ------ |
| 1   | N      |
| 2   | N      |
| 3   | N      |
| 4   | N      |
| 5   | N      |
| 6   | N      |
| 7   | N      |
| 8   | Y      |
| 9   | Y      |
| 10  | Y      |
| 11  | N      |
| 12  | N      |
| 13  | Y      |
| 14  | Y      |
| 15  | N      |
| 16  | N      |
| 17  | Y      |
| 18  | N      |
| 19  | N      |
| 20  | N      |
