using System;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace SessionMinutesPortableLauncher
{
    [Flags]
    internal enum FileOpenOptions : uint
    {
        PickFolders = 0x00000020,
        ForceFileSystem = 0x00000040,
        NoChangeDirectory = 0x00000008,
        PathMustExist = 0x00000800,
        DontAddToRecent = 0x02000000
    }

    internal enum ShellItemDisplayName : uint
    {
        FileSystemPath = 0x80058000
    }

    [ComImport]
    [Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IShellItem
    {
        void BindToHandler(IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid,
            [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem parent);
        void GetDisplayName(ShellItemDisplayName sigdnName, out IntPtr name);
        void GetAttributes(uint mask, out uint attributes);
        void Compare(IShellItem item, uint hint, out int order);
    }

    [ComImport]
    [Guid("42F85136-DB7E-439C-85F1-E4075D135FC8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IFileDialog
    {
        [PreserveSig] int Show(IntPtr owner);
        void SetFileTypes(uint count, IntPtr filters);
        void SetFileTypeIndex(uint index);
        void GetFileTypeIndex(out uint index);
        void Advise(IntPtr events, out uint cookie);
        void Unadvise(uint cookie);
        void SetOptions(FileOpenOptions options);
        void GetOptions(out FileOpenOptions options);
        void SetDefaultFolder(IShellItem folder);
        void SetFolder(IShellItem folder);
        void GetFolder(out IShellItem folder);
        void GetCurrentSelection(out IShellItem item);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string name);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string name);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string title);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string label);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string label);
        void GetResult(out IShellItem item);
        void AddPlace(IShellItem item, uint alignment);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string extension);
        void Close(int error);
        void SetClientGuid([MarshalAs(UnmanagedType.LPStruct)] Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr filter);
    }

    [ComImport]
    [Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
    internal class NativeFileOpenDialog
    {
    }

    internal static class Program
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetProcessDpiAwarenessContext(IntPtr value);

        [DllImport("shcore.dll")]
        private static extern int SetProcessDpiAwareness(int value);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetProcessDPIAware();

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SetCurrentProcessExplicitAppUserModelID(string appId);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string path,
            IntPtr bindContext,
            [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out IShellItem shellItem
        );

        [STAThread]
        private static void Main(string[] args)
        {
            EnableHighDpi();
            EnableApplicationIdentity();
            Environment.SetEnvironmentVariable(
                "SESSION_MINUTES_HOSTED",
                "1",
                EnvironmentVariableTarget.Process
            );
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);

            try
            {
                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string scriptPath = Path.Combine(baseDirectory, "app", "session_minutes_gui.ps1");
                if (!File.Exists(scriptPath))
                {
                    ShowError("تعذر العثور على ملفات البرنامج المحمولة. أعد فك ضغط الحزمة كاملة ثم شغّلها من جديد.");
                    return;
                }

                string projectRoot = ResolveProjectRoot(args);
                if (String.IsNullOrWhiteSpace(projectRoot))
                {
                    return;
                }

                Environment.CurrentDirectory = baseDirectory;
                RunGuiScript(scriptPath, projectRoot);
            }
            catch (Exception ex)
            {
                ShowError("تعذر تشغيل النسخة المحمولة:\r\n\r\n" + ex.Message);
            }
        }

        private static void EnableApplicationIdentity()
        {
            try
            {
                SetCurrentProcessExplicitAppUserModelID(
                    "SessionMinutes.Portable.Executive"
                );
            }
            catch { }
        }

        private static void RunGuiScript(string scriptPath, string projectRoot)
        {
            InitialSessionState initialState = InitialSessionState.CreateDefault();
            initialState.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;
            using (Runspace runspace = RunspaceFactory.CreateRunspace(initialState))
            {
                runspace.ApartmentState = System.Threading.ApartmentState.STA;
                runspace.ThreadOptions = PSThreadOptions.UseCurrentThread;
                runspace.Open();

                using (PowerShell shell = PowerShell.Create())
                {
                    shell.Runspace = runspace;
                    shell.AddCommand(scriptPath)
                        .AddParameter("ProjectRoot", projectRoot);
                    shell.Invoke();

                    if (shell.HadErrors)
                    {
                        StringBuilder details = new StringBuilder();
                        foreach (ErrorRecord error in shell.Streams.Error)
                        {
                            if (details.Length > 0)
                            {
                                details.AppendLine();
                            }
                            details.Append(error.ToString());
                        }
                        throw new InvalidOperationException(
                            details.Length > 0
                                ? details.ToString()
                                : "تعذر تشغيل واجهة البرنامج."
                        );
                    }
                }
            }
        }

        private static void EnableHighDpi()
        {
            try
            {
                if (SetProcessDpiAwarenessContext(new IntPtr(-4)))
                {
                    return;
                }
                if (Marshal.GetLastWin32Error() == 5)
                {
                    return;
                }
            }
            catch (EntryPointNotFoundException) { }
            catch (DllNotFoundException) { }

            try
            {
                int result = SetProcessDpiAwareness(2);
                if (result == 0 || result == unchecked((int)0x80070005))
                {
                    return;
                }
            }
            catch (EntryPointNotFoundException) { }
            catch (DllNotFoundException) { }

            try
            {
                SetProcessDPIAware();
            }
            catch { }
        }

        private static string ResolveProjectRoot(string[] args)
        {
            if (args != null && args.Length > 0)
            {
                if (!Directory.Exists(args[0]))
                {
                    ShowError("مجلد المشروع المحدد غير موجود:\r\n" + args[0]);
                    return null;
                }
                return Path.GetFullPath(args[0]);
            }

            IFileDialog dialog = null;
            IShellItem initialFolder = null;
            IShellItem selectedFolder = null;
            try
            {
                dialog = (IFileDialog)new NativeFileOpenDialog();
                dialog.SetTitle("اختر مجلد مشروع محاضر الجلسات");
                dialog.SetOkButtonLabel("اختيار هذا المجلد");
                dialog.SetFileNameLabel("مجلد المشروع");
                dialog.SetOptions(
                    FileOpenOptions.PickFolders |
                    FileOpenOptions.ForceFileSystem |
                    FileOpenOptions.PathMustExist |
                    FileOpenOptions.NoChangeDirectory |
                    FileOpenOptions.DontAddToRecent
                );

                string desktop = Environment.GetFolderPath(
                    Environment.SpecialFolder.DesktopDirectory
                );
                Guid shellItemGuid = typeof(IShellItem).GUID;
                SHCreateItemFromParsingName(
                    desktop,
                    IntPtr.Zero,
                    shellItemGuid,
                    out initialFolder
                );
                dialog.SetDefaultFolder(initialFolder);
                dialog.SetFolder(initialFolder);

                int result = dialog.Show(IntPtr.Zero);
                if (result == unchecked((int)0x800704C7))
                {
                    return null;
                }
                if (result != 0)
                {
                    Marshal.ThrowExceptionForHR(result);
                }

                dialog.GetResult(out selectedFolder);
                IntPtr pathPointer;
                selectedFolder.GetDisplayName(
                    ShellItemDisplayName.FileSystemPath,
                    out pathPointer
                );
                try
                {
                    string selectedPath = Marshal.PtrToStringUni(pathPointer);
                    return String.IsNullOrWhiteSpace(selectedPath)
                        ? null
                        : Path.GetFullPath(selectedPath);
                }
                finally
                {
                    Marshal.FreeCoTaskMem(pathPointer);
                }
            }
            catch (COMException ex)
            {
                ShowError(
                    "تعذر فتح نافذة اختيار المجلد الحديثة:\r\n" + ex.Message
                );
                return null;
            }
            finally
            {
                if (selectedFolder != null) Marshal.ReleaseComObject(selectedFolder);
                if (initialFolder != null) Marshal.ReleaseComObject(initialFolder);
                if (dialog != null) Marshal.ReleaseComObject(dialog);
            }
        }

        private static void ShowError(string message)
        {
            MessageBox.Show(
                message,
                "محاضر الجلسات - النسخة المحمولة",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading
            );
        }
    }
}
