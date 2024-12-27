# SAP GUI Scripting

SAP GUI Scripting is an interface to the [SAP GUI for Windows](https://help.sap.com/docs/sap_gui_for_windows). It was introduced in 2002 with the release 6.20. It can be used to emulate user interaction via scripting to automate the SAP GUI. The interface is implemented in Component Object Model (COM) technology. So it can be used can be used from any programming language which is capable of invoking COM objects. SAP GUI Scripting is part of the SAP GUI installation.

This repository focuses on the use of SAP GUI Scripting with a diversity of programming languages. Unusual approaches and examples will be presented here, to demonstrate the integration capabilities.

## Scripting Tracker

[Scripting Tracker](https://tracker.stschnell.de/) is a tool for simplifying the use of SAP GUI scripting. It offers an SAP GUI analyzer, scripting recorder for different programming languages, SAP GUI Scripting API viewer, code composer, screen comparator and dump state viewer on SAP GUI Scripting base. The following programming languages are supported:
- PowerShell for .NET Framework and PowerShell .NET Core
- C# for .NET Core and C# .NET Framework
- VB.NET Framework
- Python with [PyWin32](https://pypi.org/project/pywin32/)
- Java via JShell with [JaCoB (Java COM Bridge)](https://github.com/freemansoft/jacob-project)
- [AutoIt](https://www.autoitscript.com/)
- Visual BASIC for Applications (VBA) / VBScript (VBS)
- JScript