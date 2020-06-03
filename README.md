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

VBScript examples.
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