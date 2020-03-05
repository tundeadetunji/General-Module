# General-Module

Library that provides a variety of functions from customizing controls of System.Windows.Forms to returning list of IP addresses on the network the machine is connected to.

It is almost always required by the other modules, and inter-dependent on the other repositories (Web Module, NModule, Feedback, Language).

Uses:
.NET Framework 4.6

May use:
NewtonSoft.Json 10.0.0
MS SQL (with 2008 R2 in mind)

To Use:
For Windows Forms App, the main form that opens should have the tag value set to "main". Call FormatWindow from Form_Load (after copying the controls in WindowsApplication1/Main.frm and pasting unto the form). You may need to set Language and Theme. Follow the lead of Stock Manager Stores (Main.Form_Load).

Thanks!
