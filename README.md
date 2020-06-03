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

## Practical use of Active Scripting: Regular Expressions.

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