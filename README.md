# VBW
#### Python Wrapper for Microsoft's VB (via a custom VBS interpreter). ![](logo.png)



Trying to help an acquaintance automatize their Excel work-flow via Python,
we discovered that there were certain Spreadsheets we could not trivially 
edit with their regular tool-set.[1] Probably **win32com.client** or 
**xlswings** could have solved the issue, but that implied going down a 
rabbit-hole I didn't want to expose my acquaintance to. Instead, I quickly 
cooked up a wrapper that allows direct interaction with Microsoft's services 
via the use of *"System32/CScript.exe"*. 

After a few iterations of code refining, this is the current project status.
The interpreter can now be loaded with both *"error capturing code"* and 
*"exit commands"* to ensure a smooth user experience. The underlying idea 
is to wrap up "VBS commands" inside python code such that it can be used 
interchangeably with Python flow-control. Other things like the excel 
modules are still at their infancy, so please refrain from using them.

One additional advantage of this solution is that it's not restricted to 
Excel, but can instead be used as a secondary terminal for VBS. Any command
can be given, which in turn will be run by CSCript.

[1] The issue was caused by certain splicers that were not preserved upon
saving the document. Lurking online, it seems that most "easy to use
libraries" like **openpyxl** do not take those features into consideration.