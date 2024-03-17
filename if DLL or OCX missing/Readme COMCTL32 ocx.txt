Copying the OCX file:

For Windows 64-bit systems, extract the OCX file to: C:\Windows\SysWOW64

For Windows 32-bit systems, extract the OCX file to: C:\Windows\System32

Register the OCX file:

Right-click Start, click Command Prompt (Admin)

If you're using Windows 32-bit, type the following command and press ENTER:

"regsvr32 COMCTL32.ocx"

If you're using Windows 64-bit, type the following command and press ENTER:

"C:\Windows\SysWOW64\regsvr32 C:\Windows\SysWOW64\COMCTL32.ocx"

Try launching the golf program now.