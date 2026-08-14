using System;
using System.Runtime.InteropServices;

namespace RecallTape.OneNote
{
    // Hand-declared COM interfaces.
    //
    // WHY not a COMReference / PIA: TlbImp.exe ships with the Windows SDK, which is not a
    // prerequisite we want on a build box (this machine has the .NET SDK and no Visual Studio,
    // and CI shouldn't need one either). These three interfaces are the entire surface Office
    // needs from us, so declaring them is cheaper than the dependency.
    //
    // Every IID below was read off THIS machine, not recalled:
    //   IDTExtensibility2    - NOT hand-declared: comes from the embedded Extensibility PIA (see csproj)
    //   IRibbonExtensibility - HKLM\SOFTWARE\Classes\Interface (Office 16.0 Object Library, v2.8)
    //   IRibbonControl       - same

    // MUST be dual, and the method order MUST match the type library exactly.
    //
    // WHY: IDTExtensibility2 is registered for cross-process marshaling by the universal
    // typelib marshaler (HKCR\Interface\{B65AD801-...}\ProxyStubClsid32 = {00020424-...}),
    // which drives the interface through its VTABLE. Declared IsIDispatch, the CCW exposes no
    // vtable past IDispatch, so OneNote activates the add-in successfully and then calls into
    // nothing -- LoadBehavior silently drops 3 -> 2 with OnConnection never running. Measured:
    // the assembly loaded fine per the CLR bind log, and no log line was ever written.
    [ComImport, Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [TypeIdentifier]
    public interface IRibbonExtensibility
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }

    [ComImport, Guid("000C0395-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonControl
    {
        string Id { get; }
        object Context { get; }
        string Tag { get; }
    }

    // NO IRibbonUI HERE, deliberately.
    //
    // It was declared, with the correct IID read off this machine, and onLoad then never fired -
    // silently. Office binds ribbon callbacks over IDispatch and drops the ones it cannot marshal
    // without a word, so the add-in looked fine and the ribbon simply never handed itself over.
    // AddIn.OnRibbonLoad takes an OBJECT and calls InvalidateControl by name instead. Do not
    // "tidy that up" into a typed interface.
}
