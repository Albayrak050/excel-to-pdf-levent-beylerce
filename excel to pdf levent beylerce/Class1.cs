using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

class FileOpenDialogHelper
{
    [ComImport]
    [Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
    private class FileOpenDialog { }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("42f85136-db7e-439c-85f1-e4075d135fc8")]
    private interface IFileOpenDialog
    {
        void Show(IntPtr parent);
        void SetFileTypes(uint cFileTypes, [In] ref COMDLG_FILTERSPEC rgFilterSpec);
        void SetFileTypeIndex(uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise();
        void Unadvise();
        void SetOptions(uint fos);
        void GetOptions(out uint pfos);
        void SetDefaultFolder();
        void SetFolder();
        void GetFolder();
        void GetCurrentSelection();
        void SetFileName([In, MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName();
        void SetTitle([In, MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel();
        void SetFileNameLabel();
        void GetResult();
        void AddPlace();
        void SetDefaultExtension();
        void Close(int hr);
        void SetClientGuid();
        void ClearClientData();
        void SetFilter();
        void GetResults(out IShellItemArray ppenum);   // *** BİZE LAZIM OLAN ***
        void GetSelectedItems();
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct COMDLG_FILTERSPEC
    {
        public string pszName;
        public string pszSpec;
    }

    [ComImport]
    [Guid("b63ea76d-1f85-456f-a19c-48159efa858b")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItemArray
    {
        void BindToHandler();
        void GetPropertyStore();
        void GetPropertyDescriptionList();
        void GetAttributes();
        void GetCount(out uint pdwNumItems);
        void GetItemAt(uint dwIndex, out IShellItem ppsi);
    }

    [ComImport]
    [Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem
    {
        void BindToHandler();
        void GetParent();
        void GetDisplayName(uint sigdnName, out IntPtr ppszName);
        void GetAttributes();
        void Compare();
    }

    private const uint SIGDN_FILESYSPATH = 0x80058000;
    private const uint FOS_ALLOWMULTISELECT = 0x200;

    public static string[] ShowDialog(string filter, string title)
    {
        var dialog = (IFileOpenDialog)new FileOpenDialog();

        var fs = new COMDLG_FILTERSPEC
        {
            pszName = "Files",
            pszSpec = filter
        };
        dialog.SetFileTypes(1, ref fs);

        dialog.SetOptions(FOS_ALLOWMULTISELECT);
        dialog.SetTitle(title);

        dialog.Show(IntPtr.Zero);

        dialog.GetResults(out IShellItemArray items);
        items.GetCount(out uint count);

        List<string> files = new List<string>();

        for (uint i = 0; i < count; i++)
        {
            items.GetItemAt(i, out IShellItem item);
            item.GetDisplayName(SIGDN_FILESYSPATH, out IntPtr ppszName);

            string filename = Marshal.PtrToStringUni(ppszName);
            Marshal.FreeCoTaskMem(ppszName);

            files.Add(filename);
        }

        return files.ToArray();  // *** %100 SEÇİM SIRASI ***
    }
}
