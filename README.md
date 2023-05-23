# VBW
#### Python Wrapper for Microsoft's VB (via a custom VBS interpreter). ![](logo.png)

VisualBasic is a somewhat deprecated (and yet very-powerful) tool to 
manipulate Microsoft window in general, and programs in the office suite 
in particular. With that in mind, it's an interesting challenge to look 
for ways to "seamlessly" integrate it into Python.

This is different from the many libraries that already exists to interact with 
the office suite (***xlswings** -> Excel, **python-docx** -> Word, **python-pptx** 
-> Power-Point,etc.*), as they have their own abstractions and might lose some 
of VBS power. Instead, **VB Wrapper** provides access to native VBS, just calling
a VBS program that works as an interactive interpreter. (a.k.a. the wrapper) By 
building on top of it, Python "wrappers" can access VBS functionality without 
having to interpret or abstract anything of their own.

Focused on helping beginners and people that don't want to learn the ropes 
behind VB, but want to work and access some of its features. Just writing a few 
snippets of it (possible obtained online) and keep on working on Python.

#### Background.

Trying to help an acquaintance automatize their Excel work-flow via Python,
we discovered that there were certain Spreadsheets we could not trivially 
edit with their regular tool-set. The issue was caused by certain splicers 
that were not preserved upon saving the document. Lurking online, it seems 
that most "easy to use excel-libraries" such as **openpyxl** do not take 
those features into consideration.

While it could have been possible to adapt some of those Python tools, 
instead of trying to fit a square peg in a round hole, I quickly 
cooked up a wrapper that allows direct interaction with Microsoft's services 
via the use of *"System32/CScript.exe"*. After a few iterations of code 
refining, this is the current project status. 

The interpreter can now be loaded with both *"error capturing code"* and 
*"exit commands"* to ensure a smooth user experience. The underlying idea 
is to wrap up "VBS commands" inside python code such that it can be used 
interchangeably with Python flow-control. Other things like the excel 
modules are still at their infancy, so please refrain from using them.



In case you are curious, *win32com.client* is also an interesting options, 
since that gives you access to windows api in general (and a VBS runner 
in particular). However, I consider win32com to already do too much, 
and thus, be difficult to customize.
