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
    //   IDTExtensibility2    - byte-searched out of C:\Program Files\Common Files\DESIGNER\MSADDNDR.OLB
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
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom);
        [DispId(2)] void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom);
        [DispId(3)] void OnAddInsUpdate(ref Array custom);
        [DispId(4)] void OnStartupComplete(ref Array custom);
        [DispId(5)] void OnBeginShutdown(ref Array custom);
    }

    [ComImport, Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
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

    public enum ext_ConnectMode { ext_cm_AfterStartup = 0, ext_cm_Startup = 1, ext_cm_External = 2, ext_cm_CommandLine = 3 }
    public enum ext_DisconnectMode { ext_dm_HostShutdown = 0, ext_dm_UserClosed = 1 }

    // OneNote PageInfo. piSelection is the one that matters: it marks the current selection
    // inside the returned page XML, which is the foothold the whole product stands on.
    public static class PageInfo
    {
        public const int Basic = 0;
        public const int BinaryData = 1;
        public const int Selection = 2;
        public const int BinaryDataSelection = 3;
    }

    public static class HierarchyScope
    {
        public const int Self = 0, Children = 1, Notebooks = 2, Sections = 3, Pages = 4;
    }
}
