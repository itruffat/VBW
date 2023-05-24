# VBW
#### Python Wrapper for Microsoft's VB (via a custom VBS interpreter). ![](logo.png)

Despite being somewhat deprecated, VisualBasic is still a very-powerful tool 
to manipulate Microsoft window in general, and programs in the office 
suite in particular. With that in mind, it'd be a waste not to 
look for ways to "seamlessly" integrate that into our Python workflow.

**VB Wrapper** proposes a different approach from the many libraries that 
already exists to interact with the office suite (***xlswings** -> Excel, 
**python-docx** -> Word, **python-pptx** -> Power-Point,etc.*), as they have 
their own abstractions and might lose some of VBS native features in 
translation. Instead of translating, the wrapper provides access to native 
VBS by calling a VBS program (a.k.a. the interpreter) that works as an 
interactive interpreter. By building on top of it, Python "wrappers" can 
access VBS functionalities without having to abstract anything on their own.

The underlying idea is to wrap up "VBS commands" inside a python instruction
so that they can be used interchangeably with Python code. Thus, this can 
really help beginners and people that don't want to learn the ropes behind 
VB, but might need to access some of its features. Just using a few 
snippets of VBS (possible obtained online) and keep working on Python.

Additionaly, The interpreter can also be provided with customizable 
*"error capturing code"* and *"exit commands"* to ensure an smooth user 
experience. By default it already captures errors so that they don't
stop the interpreter by receiving a bad parameter.

#### Background.

Trying to help an acquaintance automatize their Excel work-flow via Python,
we discovered that there were certain Spreadsheets we could not trivially 
edit with their regular tool-set. In particular, certain splicers would not
be preserved upon saving the document. Lurking online, it seems 
that most "easy to use excel-libraries" such as **openpyxl** do not take 
those features into consideration.

While it could have been possible to adapt some of those Python tools, 
instead of trying to fit a square peg in a round hole, I quickly 
cooked up a wrapper that allows direct interaction with Microsoft's services 
via the use of *"System32/CScript.exe"*. After a few iterations of code 
refining, this is the current project status. 

In case you are curious, *win32com.client* is also an interesting options, 
since that gives you access to windows api in general (and a VBS runner 
in particular). However, I consider win32com to already do too much, 
and thus, be difficult to customize. It might be interesting to
consider using win32 instead of subprocess in the future.
