# ActiveScripting

How to run VBScript and get a result? You can RUN(‘file.vbs’) and parse output file to obtain a result, but there is another approach: [Active Scripting](https://en.wikipedia.org/wiki/Active_Scripting).
3 scripting engines are installed by default: VBScript, JScript and [ChakraJS](https://en.wikipedia.org/wiki/Chakra_(JScript_engine)), and can be easily used in Clarion applications.

VBScript, areas of use:
- Active Directory
- ADO
- Computer Hardware
- Group Policy
- IIS
- Logs
- Mathematics
- Messaging and Communication
- Microsoft Office
- Networking
- Operating System
- Other Directory Services
- Printing
- Security
- Service Packs and Hot Fixes
- Storage
- Terminal Server
- WMI

[Demo application](https://yadi.sk/d/pF49IbwH_JEh4g) allows to run scripts and evaluate expressions.

### Requirements
- C63. and higher, ABC/Clarion
- No blackbox
- No 3rd parties
- No dependencies (VC++, .NET and so on)

### Price
Core classes: $150  
Add-ons (WMI, ADO, RegEx, Windows Search): $25 for each add-on.

### Contacts
mikeduglas@yandex.ru

#### VBScript examples.
1. Calculator:
type an expression like "0.75 / 4" or "sqr(5) + log(13)" and press Evaluate expression button.

2. Reusable scripts:
paste following code (the Fibonacci numbers function) into script text box and press Run script button:

>     function fibonacci(limit)
>       dim a,b,c,res
>       a=0
>       b=1
>       res="Fibonacci numbers from 1 to "& limit & vbCrLf
> 
>       for i=1 to limit
>         c=a+b
>         a=b
>         b=c
>         res=res & c & vbCrLf
>       next
> 
>       fibonacci=res
>     end function

then you can evaluate "fibonacci" function many times with different arguments, for example type "fibonacci(10)" and press Evaluate expression button, 
next time type "fibonacci(20)" and press Evaluate expression button.

Same is true for JScript and [ChakraJS](https://en.wikipedia.org/wiki/Chakra_(JScript_engine)).

#### Practical use of Active Scripting: Regular Expressions.

Clarion has a very limited support of regular expressions. With javascript we can use the full power of regular expressions:

search for a match between a regular expression and a specified string
replace some or all matches
turn a string into an array of strings, by separating the string at each instance of a specified separator string
An example of replace feature:

>     var name1 = 'John Smith';
>     var re = /(\w+)\s(\w+)/;
>     var name2 = name1.replace(re, '$2, $1');  // expected result: 'Smith, John'
An example of split feature:

>     var s = 'Harry Trump ;Fred Barney; Helen Rigby ; Bill Abel ;Chris Hand ';
>     var re = /\s*(?:;|$)\s*/;
>     var arr = s.split(re);

I wrote a class that implements js regular expressions properties and methods:

- Test - returns true if match found
- Exec - returns an array of matches
- Match - returns an array of matches
- Search - returns the index of the first match
- Replace - returns a new string with some or all matches of a pattern replaced by a replacement
- ReplaceFunction - uses a function to be invoked to create the new substring
- Split - returns an array of strings, split at each point where the separator occurs in the given string
- LastIndex - the index at which to start the next match

so examples above can be written in Clarion like this:

>       re.CreateNew('(\w+)\s(\w+)')
>       name = re.Replace(s, '$2, $1')

and this

>       re.CreateNew('\s*(?:;|$)\s*')
>       len = re.Split(s)
>       LOOP i = 1 TO len
>         name = re.MatchedItem(i)
>       END

#### Practical use of Active Scripting: [ADO](https://en.wikipedia.org/wiki/ActiveX_Data_Objects).

For example, let’s see how to extract data from Microsoft Access database using VBScript (you can test it in the demo program):

>       ' open connection
>       Dim conn
>       Set conn = CreateObject("ADODB.Connection")
>       conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=address.mdb"
>       
>       ' execute SQL query and get recordset
>       Dim rs
>       Set rs = CreateObject("ADODB.Recordset")
>       rs.Open "SELECT * FROM address", conn
>       
>       Do Until rs.EOF
>         MsgBox "CustomerID="& rs.Fields("CustomerID") &"; Company="& rs.Fields("Company")
>         rs.MoveNext
>       Loop
>       
>       Set rs = nothing
>       Set conn = nothing

Here is same task implemented with helper class:

>       !- open connection
>       IF ado.conn.Open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=address.mdb') = adStateOpen
>       
>         !- execute SQL query and get recordset
>         IF ado.rs.Open('SELECT * FROM address') = adStateOpen
>       
>           !- loop through recordset and store field values into quueue
>           LOOP WHILE NOT ado.rs.EOF()
>             CLEAR(AddrQ)
>             Adr:CustomerID = ado.rs.Field('CustomerID')
>             Adr:Company    = ado.rs.Field('Company')
>             Adr:Street     = ado.rs.Field('Street')
>             ADD(AddrQ)
>       
>             !- next record
>             ado.rs.MoveNext()
>           END
>             
>           !- close recordset
>           ado.rs.Close()
>         ELSE
>           MESSAGE('Recordset.Open failed')
>         END
>           
>         !- close conection
>         ado.conn.Close()
>       ELSE
>         MESSAGE('Connection.Open failed')
>       END

#### Practical use of Active Scripting: [WMI](https://en.wikipedia.org/wiki/Windows_Management_Instrumentation)

Short list of WMI tasks:

- BIOS information
- Display configuration
- Video adapter configuration
- Serial ports information
- CPU information for each processor
- Disk drives information
- Physical memory information
- User accounts
- Daylight Saving Time
- Kill the specified program
- Logoff current user on any WMI enabled computer
- Ethernet adapters' link speed
- List printers with status and number of printjobs, or pause or resume printing on the specified printer(s), or flush all printjobs, or list all printers, their status and number of printjobs
- Windows Registry
- Reboot/Shut down any WMI enabled computer on the network
- Services
- Startup commands (Startup folder and registry Run)
- Synchronize your computer's system time with any webserver
- Uptime for any WMI enabled computer
- and many others

With Active Scripting any of above task is just few lines of code. For example, all disk drives information:

>             wmi.Connect()
>             wmi.ExecQuery('Select * from Win32_DiskDrive')
>             
>             LOOP i=1 TO wmi.items.Count()
>               CLEAR(DriveQ)
>               DriveQ:Caption           = wmi.items.GetProp(i, 'Caption')
>               DriveQ:Description       = wmi.items.GetProp(i, 'Description')
>               DriveQ:Manufacturer      = wmi.items.GetProp(i, 'Manufacturer')
>               DriveQ:Model             = wmi.items.GetProp(i, 'Model')
>               DriveQ:Name              = wmi.items.GetProp(i, 'Name')
>               DriveQ:Partitions        = wmi.items.GetProp(i, 'Partitions')
>               DriveQ:Size              = wmi.items.GetProp(i, 'Size')
>               DriveQ:Status            = wmi.items.GetProp(i, 'Status')
>               DriveQ:SystemName        = wmi.items.GetProp(i, 'SystemName')
>               DriveQ:TotalCylinders    = wmi.items.GetProp(i, 'TotalCylinders')
>               DriveQ:TotalHeads        = wmi.items.GetProp(i, 'TotalHeads')
>               DriveQ:TotalSectors      = wmi.items.GetProp(i, 'TotalSectors')
>               DriveQ:TotalTracks       = wmi.items.GetProp(i, 'TotalTracks')
>               DriveQ:TracksPerCylinder = wmi.items.GetProp(i, 'TracksPerCylinder')
>               ADD(DriveQ)
>             END

How to terminate running process:

>             wmi.Connect()
>             !- find all processes with specific caption
>             wmi.ExecQuery(printf('Select * from Win32_Process WHERE Caption=%S', pProcess))
>       
>             !- terminate each process
>             LOOP i=1 TO wmi.items.Count()
>               wmi.items.CallMethod(i, 'Terminate')
>             END

How to reboot specific computer:

>             !- required "Shutdown" privelege
>             wmi.Connect(pMachineName, 'Shutdown')
>       
>             !- find primary OS
>             wmi.ExecQuery('SELECT * FROM Win32_OperatingSystem WHERE Primary=True')
>       
>             !- reboot
>             wmi.items.CallMethod(1, 'Reboot()')
>             IF wmi.ErrNumber()
>               MESSAGE(wmi.ErrDescription())
>             END

Querying the amount of memory a particular process uses:

>             wmi.Connect()
>             wmi.ExecQuery(printf('SELECT * FROM Win32_Process WHERE Name = %S', pProcess))
>             LOOP i=1 TO wmi.items.Count()
>               CLEAR(MemQ)
>               MemQ:Amount = wmi.items.GetProp(1, 'WorkingSetSize')
>               ADD(MemQ)
>             END

#### Practical use of Active Scripting: [Windows Search](https://docs.microsoft.com/ru-ru/windows/win32/search/-search-3x-wds-overview?redirectedfrom=MSDN)

Windows Search is a desktop search platform that has instant search capabilities for most common file types and data types.

Let’s look at some examples.

**Search by file name**
To find all “setup.exe” files you define “System.FileName=‘setup.exe’” condition:

>               ws.Search('SELECT System.ItemPathDisplay FROM SystemIndex WHERE System.FileName=''setup.exe''')
>               LOOP UNTIL ws.EOF()
>                 CLEAR(FilesQ)
>                 FilesQ:Name = ws.rs.Field('System.ItemPathDisplay')
>                 ADD(FilesQ)
>                 ws.MoveNext()
>               END

To perform a search by mask (“setup.*”) use ‘LIKE’ operator:

>               ws.Search('SELECT System.ItemPathDisplay FROM SystemIndex WHERE System.FileName LIKE ''setup.*''')

**Search by file content**
You can search for words and phrases, below “FREETEXT” predicate is used:

>               ws.Search('SELECT System.ItemPathDisplay FROM SystemIndex WHERE FREETEXT(''Windows Search'')')

**Search by document type**
Next example finds all big pictures (image size is greater than 1024x768):

>               ws.Search('SELECT System.ItemPathDisplay FROM SystemIndex WHERE System.Kind = ''Picture'' AND System.Image.VerticalSize >= 768 AND System.Image.HorizontalSize >= 1024')

You can request not only System.ItemPathDisplay (which actually is full filename), but any type of information: size, modified date, author, image width, color depth and so on.

Windows Search allows requests against remote machine as well, all you need is machine name:
>               SELECT System.ItemPathDisplay FROM MACHINENAME.SystemIndex

You can add folders to Search Index using [this project](https://github.com/mikeduglas/Search-API).