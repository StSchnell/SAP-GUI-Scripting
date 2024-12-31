# SAP GUI Scripting

SAP GUI Scripting is an interface to the [SAP GUI for Windows](https://help.sap.com/docs/sap_gui_for_windows). It was introduced in 2002 with the release 6.20. It can be used to emulate user interaction via scripting to automate the SAP GUI. The interface is implemented in Component Object Model (COM) technology. So it can be used can be used from any programming language which is capable of invoking COM objects. SAP GUI Scripting is part of the SAP GUI installation.

This repository focuses on the use of SAP GUI Scripting with a diversity of programming languages. Unusual approaches and examples will be presented here, to demonstrate the integration capabilities.

## Scripting Tracker

[Scripting Tracker](https://tracker.stschnell.de/) is a tool for simplifying the use of SAP GUI scripting. It offers an SAP GUI analyzer, scripting recorder for different programming languages, SAP GUI Scripting API viewer, code composer, screen comparator and dump state viewer on SAP GUI Scripting base.

The following programming languages are supported:
- PowerShell for .NET Framework and [PowerShell .NET Core](https://github.com/PowerShell/PowerShell)
- C# for .NET Core and C# .NET Framework
- VB.NET Framework
- [Python](https://www.python.org/) with [PyWin32](https://pypi.org/project/pywin32/)
- Java via JShell with [JaCoB (Java COM Bridge)](https://github.com/freemansoft/jacob-project)
- [AutoIt](https://www.autoitscript.com/)
- Visual BASIC for Applications (VBA) / VBScript (VBS) via Windows Script Host (WSH)
- JScript via Windows Script Host (WSH)

## Component Object Model

The [Component Object Model](https://learn.microsoft.com/en-us/windows/win32/com/component-object-model--com--portal) (COM) is a system for creating binary software components. This system itself is platform-independent, but in reality it is only available on the Windows OS. This is important to know, because we are more or less talking about a native system here.

## Windows Script Host

The Windows Script Host (WSH) offers an environment in which users can execute scripts in various languages. The WSH cannot execute scripts itself, it uses script engines. In the standard the WSH contains the script engines for the programming languages VBScript and JScript. The standard SAP GUI Scripting Recorder records user activities in the VBScript programming language. But in the future VBScript will be an optional component that will be removed step by step from the Windows operating system. Even if many examples are available, it is no longer recommended to use VBScript for future automations. In general, the use of the WSH is no longer recommended.

PowerShell for .NET Framework is available in the standard of the Windows OS. It can be used very well as a replacement for this use case. It also offers many other possibilities and advantages that could never be realized with VBScript. Therefore the clear recommendation to use PowerShell for the automation of the SAP GUI for Windows with SAP GUI Scripting.

## Alternatives

### Python.NET

[Python.NET (pythonnet)](https://pypi.org/project/pythonnet/) offers the possibility to integrate .NET seamlessly into Python programming language. This makes it possible to use COM libraries with Python via .NET, and with that also the SAP GUI Scripting. Here an [example how to use SAP GUI Scripting via Python.NET](https://github.com/StSchnell/SAP-GUI-Scripting/blob/main/sapGuiScripting.py). This approach is a basic example of using COM libraries with Python.NET. It is a real alternative approach if as few Python packages as possible should to be used and other less necessary packages are excluded whose functionality is already available, such the COM interface as in this example.

### JavaScript

JavaScript is a widely used programming language. There are different variants and platforms with which it can be used. [JavaScript via Rhino Engine](https://github.com/mozilla/rhino) is one of them. It bases on Java and with JaCoB (Java COM Bridge) it offers the possibility to use COM seamlessly. Here a basic [example how to use SAP GUI Scripting via JavaScript with the Rhino engine and JaCoB](https://github.com/StSchnell/SAP-GUI-Scripting/blob/main/sapGuiScripting.js). On this way the functional variety of Java can be combined with COM and the simplicity of JavaScript.
