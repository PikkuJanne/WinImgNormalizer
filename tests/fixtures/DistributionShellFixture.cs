// Development-only instrumentation. ShellExecuteEx opens a real BAT shortcut;
// a private inherited job records actual children, executable hashes and exits.
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

public static class DistributionShellFixture {
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)] struct ShellInfo {
        public int Size; public uint Mask; public IntPtr Window;
        [MarshalAs(UnmanagedType.LPWStr)] public string Verb, File, Parameters, Directory;
        public int Show; public IntPtr Instance, IdList;
        [MarshalAs(UnmanagedType.LPWStr)] public string Class;
        public IntPtr ClassKey; public uint HotKey; public IntPtr Icon, Process;
    }
    [DllImport("shell32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern bool ShellExecuteExW(ref ShellInfo info);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool AllocConsole();
    [DllImport("kernel32.dll")] static extern bool FreeConsole();
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool AttachConsole(uint process);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool SetStdHandle(int kind,IntPtr handle);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool SetHandleInformation(IntPtr handle,uint mask,uint flags);
    [DllImport("kernel32.dll")] static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr window, int command);
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern IntPtr CreateFileW(string path,uint access,uint sharing,IntPtr security,uint disposition,uint flags,IntPtr template);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr handle);
    [StructLayout(LayoutKind.Sequential)] struct Coord { public short X,Y; public Coord(short x,short y) { X=x;Y=y; } }
    [StructLayout(LayoutKind.Sequential)] struct Rect { public short Left,Top,Right,Bottom; }
    [StructLayout(LayoutKind.Sequential)] struct ScreenInfo { public Coord Size,Cursor;public ushort Attributes;public Rect Window;public Coord Maximum; }
    [StructLayout(LayoutKind.Explicit, CharSet=CharSet.Unicode, Size=4)] struct CharInfo { [FieldOffset(0)] public char Character; [FieldOffset(2)] public ushort Attributes; }
    [DllImport("kernel32.dll",SetLastError=true)] static extern bool GetConsoleScreenBufferInfo(IntPtr screen,out ScreenInfo info);
    [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern bool ReadConsoleOutputW(IntPtr screen,[Out] CharInfo[] cells,Coord size,Coord start,ref Rect region);
    [DllImport("kernel32.dll",SetLastError=true)] static extern bool SetConsoleScreenBufferSize(IntPtr screen,Coord size);
    [StructLayout(LayoutKind.Sequential)] struct JobBasic { public long ProcessTime,JobTime;public uint Flags;public UIntPtr MinWorking,MaxWorking;public uint ActiveLimit;public UIntPtr Affinity;public uint Priority,Scheduling; }
    [StructLayout(LayoutKind.Sequential)] struct IoCounters { public ulong A,B,C,D,E,F; }
    [StructLayout(LayoutKind.Sequential)] struct JobExtended { public JobBasic Basic;public IoCounters Io;public UIntPtr A,B,C,D; }
    [StructLayout(LayoutKind.Sequential)] struct JobPort { public IntPtr Key,Port; }
    [DllImport("kernel32.dll",SetLastError=true)] static extern IntPtr CreateJobObjectW(IntPtr attributes,IntPtr name);
    [DllImport("kernel32.dll",SetLastError=true)] static extern bool SetInformationJobObject(IntPtr job,int kind,ref JobExtended info,uint size);
    [DllImport("kernel32.dll",SetLastError=true)] static extern bool SetInformationJobObject(IntPtr job,int kind,ref JobPort info,uint size);
    [DllImport("kernel32.dll",SetLastError=true)] static extern bool AssignProcessToJobObject(IntPtr job,IntPtr process);
    [DllImport("kernel32.dll",SetLastError=true)] static extern IntPtr CreateIoCompletionPort(IntPtr file,IntPtr port,UIntPtr key,uint threads);
    [DllImport("kernel32.dll",SetLastError=true)] static extern bool GetQueuedCompletionStatus(IntPtr port,out uint message,out UIntPtr key,out IntPtr pid,uint timeout);
    [DllImport("kernel32.dll",SetLastError=true)] static extern IntPtr OpenProcess(uint access,bool inherit,uint pid);
    [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern bool QueryFullProcessImageNameW(IntPtr process,uint flags,StringBuilder value,ref uint size);
    [DllImport("kernel32.dll")] static extern bool GetExitCodeProcess(IntPtr process,out uint code);
    [DllImport("kernel32.dll")] static extern uint WaitForSingleObject(IntPtr process,uint milliseconds);
    [DllImport("kernel32.dll")] static extern uint GetProcessId(IntPtr process);
    [DllImport("ntdll.dll")] static extern int NtQueryInformationProcess(IntPtr process,int kind,IntPtr buffer,int bytes,out int required);
    [StructLayout(LayoutKind.Sequential)] struct UnicodeString { public ushort Length,Maximum;public IntPtr Buffer; }
    static void Check(bool value,string operation) { if(!value) throw new Win32Exception(Marshal.GetLastWin32Error(),operation); }
    static string CommandLine(IntPtr process) {
        int size; NtQueryInformationProcess(process,60,IntPtr.Zero,0,out size);
        if(size<=0 || size>65536) return "";
        IntPtr buffer=Marshal.AllocHGlobal(size);
        try { if(NtQueryInformationProcess(process,60,buffer,size,out size)!=0)return "";var text=(UnicodeString)Marshal.PtrToStructure(buffer,typeof(UnicodeString));return Marshal.PtrToStringUni(text.Buffer,text.Length/2); }
        finally { Marshal.FreeHGlobal(buffer); }
    }
    static string Capture(IntPtr screen) {
        ScreenInfo info;if(!GetConsoleScreenBufferInfo(screen,out info)) return "";
        int width=info.Size.X,rows=info.Cursor.Y+1;
        if(width<1 || rows<1 || rows>5000 || (long)width*rows>1000000)return "";
        var cells=new CharInfo[width*rows];var region=new Rect{Left=0,Top=0,Right=(short)(width-1),Bottom=(short)(rows-1)};
        if(!ReadConsoleOutputW(screen,cells,new Coord((short)width,(short)rows),new Coord(0,0),ref region))return "";
        var text=new StringBuilder();for(int row=0;row<rows;row++){var line=new StringBuilder();for(int col=0;col<width;col++){var cell=cells[row*width+col];if((cell.Attributes&0x0200)==0)line.Append(cell.Character=='\0'?' ':cell.Character);}text.AppendLine(line.ToString().TrimEnd());}return text.ToString();
    }
    public static int Main(string[] args) {
        var record=new Dictionary<string,object>();var processes=new List<Dictionary<string,object>>();
        var handles=new Dictionary<uint,IntPtr>();var records=new Dictionary<uint,Dictionary<string,object>>();
        var hashes=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
        var hashThreads=new Dictionary<string,Thread>(StringComparer.OrdinalIgnoreCase);
        IntPtr job=IntPtr.Zero,port=IntPtr.Zero,screen=IntPtr.Zero,input=IntPtr.Zero,shellProcess=IntPtr.Zero;
        bool running=true;Thread observer=null;int exit=1;string captured="";string failure=null;
        string recordPath=args.Length==3?args[2]:(args.Length==4?args[3]:null);
        try {
            if(args.Length!=3 && args.Length!=4)throw new Exception("Expected shortcut, unrelated working directory, private record; or executable, arguments, working directory, record.");
            job=CreateJobObjectW(IntPtr.Zero,IntPtr.Zero);Check(job!=IntPtr.Zero,"CreateJobObject");
            var limits=new JobExtended();limits.Basic.Flags=0x2000;Check(SetInformationJobObject(job,9,ref limits,(uint)Marshal.SizeOf(typeof(JobExtended))),"Job limits");
            port=CreateIoCompletionPort(new IntPtr(-1),IntPtr.Zero,UIntPtr.Zero,1);Check(port!=IntPtr.Zero,"Completion port");
            var association=new JobPort{Key=new IntPtr(1),Port=port};Check(SetInformationJobObject(job,7,ref association,(uint)Marshal.SizeOf(typeof(JobPort))),"Job completion port");
            using(var self=Process.GetCurrentProcess())Check(AssignProcessToJobObject(job,self.Handle),"Assign private job");
            record["PrivateJobAssignedBeforeShell"]=true;
            FreeConsole();Check(AllocConsole(),"AllocConsole");ShowWindow(GetConsoleWindow(),0);
            screen=CreateFileW("CONOUT$",0xC0000000,3,IntPtr.Zero,3,0,IntPtr.Zero);Check(screen!=new IntPtr(-1),"CONOUT$");SetConsoleScreenBufferSize(screen,new Coord(180,5000));
            input=CreateFileW("CONIN$",0xC0000000,3,IntPtr.Zero,3,0,IntPtr.Zero);Check(input!=new IntPtr(-1),"CONIN$");
            // The controller itself was started with redirected parent pipes.
            // AllocConsole preserves STARTF_USESTDHANDLES in that case; replace
            // those inherited handles explicitly before ShellExecuteEx creates
            // CMD so the real child renders into our private console buffer.
            Check(SetHandleInformation(screen,1,1),"Inheritable console output");Check(SetHandleInformation(input,1,1),"Inheritable console input");
            Check(SetStdHandle(-10,input),"Console stdin");Check(SetStdHandle(-11,screen),"Console stdout");Check(SetStdHandle(-12,screen),"Console stderr");
            observer=new Thread(delegate(){
                try { while(running) {
                    uint message;UIntPtr key;IntPtr processId;
                    if(!GetQueuedCompletionStatus(port,out message,out key,out processId,30))continue;
                    uint pid=(uint)processId.ToInt64();if(pid==Process.GetCurrentProcess().Id)continue;
                    if(message==6) {
                        IntPtr process=OpenProcess(0x1000,false,pid);
                        if(process==IntPtr.Zero)throw new Win32Exception(Marshal.GetLastWin32Error(),"Could not inspect an actual owned child "+pid);
                        var path=new StringBuilder(32768);uint length=(uint)path.Capacity;Check(QueryFullProcessImageNameW(process,0,path,ref length),"Query owned image");
                        string imagePath=path.ToString();
                        // Open every reported child immediately. Hashing 32 MB
                        // inline delayed the next NEW_PROCESS notification past
                        // short probes' lifetimes on a busy hosted runner.
                        // Hash each observed path concurrently while the app
                        // retains its runtime read locks for this whole run.
                        if(!hashThreads.ContainsKey(imagePath)) {
                            var hashThread=new Thread(delegate(){try {
                                string hash;using(var algorithm=SHA256.Create())using(var stream=File.OpenRead(imagePath))hash=BitConverter.ToString(algorithm.ComputeHash(stream)).Replace("-","").ToLowerInvariant();
                                lock(hashes){hashes[imagePath]=hash;}
                            }catch(Exception error){failure=error.ToString();}});
                            hashThread.IsBackground=true;hashThreads[imagePath]=hashThread;hashThread.Start();
                        }
                        var item=new Dictionary<string,object>{{"Pid",pid},{"Path",imagePath},{"CommandLine",CommandLine(process)}};
                        lock(processes){processes.Add(item);handles[pid]=process;records[pid]=item;}
                    } else if(message==7 || message==8) {
                        lock(processes){if(handles.ContainsKey(pid)){uint code;Check(GetExitCodeProcess(handles[pid],out code),"Owned exit");records[pid]["ExitCode"]=code;}}
                    }
                }}catch(Exception error){failure=error.ToString();}
            });observer.IsBackground=true;observer.Priority=ThreadPriority.AboveNormal;observer.Start();
            if(args.Length==3) {
                var shell=new ShellInfo{Size=Marshal.SizeOf(typeof(ShellInfo)),Mask=0x40|0x100|0x400|0x8000,Verb="open",File=args[0],Directory=args[1],Show=0};
                Check(ShellExecuteExW(ref shell),"ShellExecuteEx real BAT shortcut");shellProcess=shell.Process;
                record["Entry"]="ShellExecuteEx BAT shortcut (drop-equivalent)";
            } else {
                var start=new ProcessStartInfo(args[0],args[1]);start.WorkingDirectory=args[2];start.UseShellExecute=false;start.CreateNoWindow=false;start.WindowStyle=ProcessWindowStyle.Hidden;
                var direct=Process.Start(start);shellProcess=OpenProcess(0x100000|0x1000,false,(uint)direct.Id);direct.Dispose();
                record["Entry"]="Direct process entry";
            }
            record["ShellReturnedProcessHandle"]=shellProcess!=IntPtr.Zero;
            if(shellProcess==IntPtr.Zero)throw new Exception("Shell returned no process handle; child instrumentation retained, harness requires bounded completion.");
            uint shellPid=GetProcessId(shellProcess);record["ShellProcessId"]=shellPid;
            var clock=Stopwatch.StartNew();bool attached=false;
            while(WaitForSingleObject(shellProcess,10)==258 && clock.ElapsedMilliseconds<120000) {
                string snapshot=Capture(screen);if(snapshot.Length>captured.Length)captured=snapshot;
                if(!attached && clock.ElapsedMilliseconds>300 && captured.Trim().Length==0) {
                    CloseHandle(screen);screen=IntPtr.Zero;FreeConsole();
                    if(AttachConsole(shellPid)) {ShowWindow(GetConsoleWindow(),0);screen=CreateFileW("CONOUT$",0xC0000000,3,IntPtr.Zero,3,0,IntPtr.Zero);}
                    attached=true;
                }
            }
            if(clock.ElapsedMilliseconds>=120000)throw new Exception("Shell child exceeded bounded run.");
            string final=Capture(screen);if(final.Length>captured.Length)captured=final;
            uint actual;Check(GetExitCodeProcess(shellProcess,out actual),"Shell child actual exit");record["ShellProcessExitCode"]=actual;
            Thread.Sleep(150);running=false;observer.Join(2000);
            foreach(var hashThread in hashThreads.Values)if(!hashThread.Join(10000))throw new Exception("Observed executable hash exceeded its bound.");
            if(failure!=null)throw new Exception(failure);
            lock(processes){
                foreach(var item in processes){string imagePath=(string)item["Path"];lock(hashes){item["Sha256"]=hashes[imagePath];}}
                foreach(var pair in handles){uint code;Check(GetExitCodeProcess(pair.Value,out code),"Final owned exit");records[pair.Key]["ExitCode"]=code;}
                bool observed=false;foreach(var item in processes)if(Convert.ToUInt32(item["Pid"])==shellPid)observed=true;
                if(!observed)throw new Exception("Shell process escaped owned job instrumentation.");
            }
            record["ObservedChildExitCode"]=actual;record["Console"]=captured;exit=(int)actual;
        }catch(Exception error){record["HarnessError"]=error.ToString();}
        finally {
            running=false;if(observer!=null)observer.Join(2000);
            record["Processes"]=processes;record["ControllerExitCode"]=exit;
            if(recordPath!=null)File.WriteAllText(recordPath,new JavaScriptSerializer{MaxJsonLength=4000000}.Serialize(record),new UTF8Encoding(false));
            foreach(var value in handles.Values)CloseHandle(value);
            if(shellProcess!=IntPtr.Zero)CloseHandle(shellProcess);if(screen!=IntPtr.Zero && screen!=new IntPtr(-1))CloseHandle(screen);if(input!=IntPtr.Zero && input!=new IntPtr(-1))CloseHandle(input);FreeConsole();
            // Job handle deliberately remains open until this controller exits.
        }
        return exit;
    }
}
