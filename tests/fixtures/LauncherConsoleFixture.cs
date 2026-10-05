// Development-only Windows fixture. Compile with the existing Framework csc as
// a winexe; the application and its distributed launcher do not use this file.
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

public static class LauncherConsoleFixture {
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool AllocConsole();
    [DllImport("kernel32.dll")] static extern bool FreeConsole();
    [DllImport("kernel32.dll")] static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr window, int command);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool SetStdHandle(int kind, IntPtr handle);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool SetHandleInformation(IntPtr handle, uint mask, uint flags);
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern IntPtr CreateFileW(string path, uint access, uint sharing, IntPtr security, uint disposition, uint flags, IntPtr template);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr handle);
    [DllImport("kernel32.dll", SetLastError=true)] static extern uint GetConsoleProcessList(uint[] members, uint count);
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)] struct KeyEvent { public int Down; public ushort Repeat, Virtual, Scan; public char Character; public uint Controls; }
    [StructLayout(LayoutKind.Explicit, Size=20)] struct InputRecord { [FieldOffset(0)] public ushort Kind; [FieldOffset(4)] public KeyEvent Key; }
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern bool WriteConsoleInputW(IntPtr input, InputRecord[] records, uint length, out uint written);
    [StructLayout(LayoutKind.Sequential)] struct JobBasic { public long ProcessTime, JobTime; public uint Flags; public UIntPtr MinWorking, MaxWorking; public uint ActiveLimit; public UIntPtr Affinity; public uint Priority, Scheduling; }
    [StructLayout(LayoutKind.Sequential)] struct IoCounters { public ulong ReadOps, WriteOps, OtherOps, ReadBytes, WriteBytes, OtherBytes; }
    [StructLayout(LayoutKind.Sequential)] struct JobExtended { public JobBasic Basic; public IoCounters Io; public UIntPtr ProcessMemory, JobMemory, PeakProcessMemory, PeakJobMemory; }
    [DllImport("kernel32.dll", SetLastError=true)] static extern IntPtr CreateJobObjectW(IntPtr attributes, IntPtr name);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool SetInformationJobObject(IntPtr job, int kind, ref JobExtended limits, uint length);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
    static void Check(bool value, string operation) { if (!value) throw new Win32Exception(Marshal.GetLastWin32Error(), operation); }

    public static int Main() {
        // Capture the parent's pipes before console allocation replaces the
        // standard handles. Forward actual bytes without changing the encoding.
        Stream parentOutput = Console.OpenStandardOutput(), parentError = Console.OpenStandardError();
        StreamReader parentInput = new StreamReader(Console.OpenStandardInput());
        Process child = null;
        IntPtr input = IntPtr.Zero, job = IntPtr.Zero;
        bool allocated = false;
        int result = 1, inputEnabled = 1;
        Thread keyThread = null;
        Exception inputError = null;
        var record = new Dictionary<string, object>();
        try {
            job = CreateJobObjectW(IntPtr.Zero, IntPtr.Zero); Check(job != IntPtr.Zero, "CreateJobObjectW");
            var limits = new JobExtended(); limits.Basic.Flags = 0x2000; // KILL_ON_JOB_CLOSE
            Check(SetInformationJobObject(job, 9, ref limits, (uint)Marshal.SizeOf(typeof(JobExtended))), "SetInformationJobObject");
            using (var self = Process.GetCurrentProcess()) { Check(AssignProcessToJobObject(job, self.Handle), "AssignProcessToJobObject"); }
            record["PrivateJobAssignedBeforeCmd"] = true;
            // Only this fixture detaches; the Pester host's console is untouched.
            FreeConsole(); Check(AllocConsole(), "AllocConsole"); allocated = true;
            ShowWindow(GetConsoleWindow(), 0);
            record["PrivateConsoleAllocated"] = true;
            input = CreateFileW("CONIN$", 0xC0000000, 3, IntPtr.Zero, 3, 0, IntPtr.Zero);
            Check(input != new IntPtr(-1), "CONIN$");
            Check(SetHandleInformation(input, 1, 1), "SetHandleInformation");
            Check(SetStdHandle(-10, input), "SetStdHandle");

            var start = new ProcessStartInfo(Environment.GetEnvironmentVariable("WINIMG_TEST_PRIVATE_CONSOLE_CMD"), Environment.GetEnvironmentVariable("WINIMG_TEST_PRIVATE_CONSOLE_ARGUMENTS"));
            start.UseShellExecute = false; start.CreateNoWindow = false; start.WindowStyle = ProcessWindowStyle.Hidden;
            // CMD uses the private console input. Only output/error are pipes.
            start.RedirectStandardOutput = true; start.RedirectStandardError = true;
            child = Process.Start(start);
            record["ControllerId"] = Process.GetCurrentProcess().Id;
            record["CmdId"] = child.Id; record["CmdStartTicks"] = child.StartTime.ToUniversalTime().Ticks;
            var output = child.StandardOutput.BaseStream.CopyToAsync(parentOutput);
            var error = child.StandardError.BaseStream.CopyToAsync(parentError);
            string recordPath = Environment.GetEnvironmentVariable("WINIMG_TEST_PRIVATE_CONSOLE_RECORD");
            File.WriteAllText(recordPath + ".identity.json", new JavaScriptSerializer().Serialize(record), new UTF8Encoding(false));
            File.WriteAllText(recordPath + ".ready", "Owned CMD was created with private console input.", new UTF8Encoding(false));
            keyThread = new Thread(delegate() {
                try {
                    string line = parentInput.ReadLine();
                    if (line == null || Interlocked.CompareExchange(ref inputEnabled, 0, 1) != 1 || child.HasExited) return;
                    uint[] members = new uint[32];
                    uint count = GetConsoleProcessList(members, (uint)members.Length);
                    if (count != 2) throw new Exception("Private console must contain only this controller and its CMD after driver exit.");
                    uint[] exact = new uint[count]; Array.Copy(members, exact, count);
                    bool self = false, target = false;
                    foreach (uint member in exact) { if (member == Process.GetCurrentProcess().Id) self = true; if (member == child.Id) target = true; }
                    if (!self || !target) throw new Exception("Owned controller and CMD are absent from private console.");
                    InputRecord[] keys = new InputRecord[] {
                        new InputRecord { Kind = 1, Key = new KeyEvent { Down = 1, Repeat = 1, Virtual = 88, Character = 'x' } },
                        new InputRecord { Kind = 1, Key = new KeyEvent { Down = 0, Repeat = 1, Virtual = 88, Character = 'x' } }
                    };
                    uint written; Check(WriteConsoleInputW(input, keys, 2, out written) && written == 2, "WriteConsoleInputW");
                    lock (record) { record["MembersBeforeKey"] = exact; record["KeyRecordsWritten"] = written; }
                } catch (Exception ex) { lock (record) { inputError = ex; } }
            });
            keyThread.IsBackground = true; keyThread.Start();
            // This is CMD's wait, not a wait for the parent to supply a key.
            // A no-pause command must exit promptly even with parent stdin open.
            if (!child.WaitForExit(45000)) throw new Exception("Owned CMD exceeded the private console fixture's 45-second bound.");
            result = child.ExitCode;
            lock (record) { record["ExitCode"] = result; }
            if (!output.Wait(5000) || !error.Wait(5000)) throw new Exception("Owned CMD streams did not finish.");
            keyThread.Join(100);
            lock (record) { if (inputError != null) throw inputError; }
        } catch (Exception ex) {
            result = 1;
            lock (record) { record["HarnessError"] = ex.ToString(); }
        } finally {
            Interlocked.Exchange(ref inputEnabled, 0);
            lock (record) {
                record["ControllerExitCode"] = result; record["FinishedUtc"] = DateTime.UtcNow.ToString("o");
                File.WriteAllText(Environment.GetEnvironmentVariable("WINIMG_TEST_PRIVATE_CONSOLE_RECORD"), new JavaScriptSerializer().Serialize(record), new UTF8Encoding(false));
            }
            if (child != null) child.Dispose();
            if (input != IntPtr.Zero && input != new IntPtr(-1)) CloseHandle(input);
            if (allocated) FreeConsole();
            parentOutput.Flush(); parentError.Flush();
        }
        // Keep the private job handle open until process exit. Windows closes
        // it on normal return or parent force termination and kills this tree.
        // The fixture joins no parent/user processes and has no taskkill fallback.
        return result;
    }
}
