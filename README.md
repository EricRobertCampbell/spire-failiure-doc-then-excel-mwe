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

In this case the result is slightly more informative:

```
Unhandled Exception: System.TypeInitializationException: TypeInitialization_Type_NoTypeAvailable
   ---> System.TypeInitializationException: TypeInitialization_Type_NoTypeAvailable
   ---> System.TypeInitializationException: TypeInitialization_Type_NoTypeAvailable
   ---> System.DllNotFoundException: DllNotFound_Linux, libSkiaSharp,
   libfontconfig.so.1: cannot open shared object file: No such file or directory                                                                                                                                                          liblibSkiaSharp.so: cannot open shared object file: No such file or directory                                                                                                                                                              libSkiaSharp: cannot open shared object file: No such file or directory                                                                                                                                                        liblibSkiaSharp: cannot open shared object file: No such file or directory                                                                                                                                                                                                                                                                                                                                                                                                               at System.Runtime.InteropServices.NativeLibrary.LoadLibErrorTracker.Throw(String) + 0x4e                                                                                                                                                   at Internal.Runtime.CompilerHelpers.InteropHelpers.FixupModuleCell(InteropHelpers.ModuleFixupCell*) + 0x10f                                                                                                                                at Internal.Runtime.CompilerHelpers.InteropHelpers.ResolvePInvokeSlow(InteropHelpers.MethodFixupCell*) + 0x35                                                                                                                              at SkiaSharp.SkiaApi.sk_colortype_get_default_8888() + 0x2d                                                                                                                                                                                at SkiaSharp.SKImageInfo..cctor() + 0x2b
   at System.Runtime.CompilerServices.ClassConstructorRunner.EnsureClassConstructorRun(StaticClassConstructionContext*) + 0xb9
   Exception_EndOfInnerExceptionStack
   at System.Runtime.CompilerServices.ClassConstructorRunner.EnsureClassConstructorRun(StaticClassConstructionContext*) + 0x153
   at System.Runtime.CompilerServices.ClassConstructorRunner.CheckStaticClassConstructionReturnNonGCStaticBas
e(StaticClassConstructionContext*, IntPtr) + 0x9
   at sprmtx..cctor() + 0x690
   at System.Runtime.CompilerServices.ClassConstructorRunner.EnsureClassConstructorRun(StaticClassConstructionContext*) + 0xb9
   Exception_EndOfInnerExceptionStack
   at System.Runtime.CompilerServices.ClassConstructorRunner.EnsureClassConstructorRun(StaticClassConstructionContext*) + 0x153
   at System.Runtime.CompilerServices.ClassConstructorRunner.CheckStaticClassConstructionReturnGCStaticBase(StaticClassConstructionContext*, Object) + 0x9
   at Spire.Xls.Core.Spreadsheet.XlsPageSetupBase.PaperSizeEntry..ctor(Double, Double, MeasureUnits) + 0x1e
   at Spire.Xls.Core.Spreadsheet.XlsPageSetupBase..cctor() + 0x6c
   at System.Runtime.CompilerServices.ClassConstructorRunner.EnsureClassConstructorRun(StaticClassConstructionContext*) + 0xb9
   Exception_EndOfInnerExceptionStack
   at System.Runtime.CompilerServices.ClassConstructorRunner.EnsureClassConstructorRun(StaticClassConstructionContext*) + 0x153
   at System.Runtime.CompilerServices.ClassConstructorRunner.CheckStaticClassConstructionReturnNonGCStaticBase(StaticClassConstructionContext*, IntPtr) + 0x9
   at Spire.Xls.Core.Spreadsheet.XlsWorksheet.InitializeCollections() + 0xdd
   at Spire.Xls.Core.Spreadsheet.XlsWorksheet..ctor(Object) + 0x82
   at Spire.Xls.Core.Spreadsheet.Collections.XlsWorksheetsCollection.Add(String) + 0x190
   at Spire.Xls.Core.Spreadsheet.XlsWorkbook.spra(Int32) + 0x461
   at Spire.Xls.Workbook..ctor() + 0x44
   at Spire.Xls.AOT.NLWorkbook.CreateWorkbook() + 0x2b
Aborted (core dumped)
```
