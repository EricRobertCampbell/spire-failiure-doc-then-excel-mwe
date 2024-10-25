Spire.Xls and Spire.Doc - Sequential Document Generation Failure

This is an MRE for a bug in Spire.Xls 13.9.0 / Spire.Doc 12.71. (or some combination thereof). In some circumstances, generating a .doc file followed by a .xls file will result in a crash.

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
