# Debugging Property Page

**Persistence**: `PST_USER_FILE` -> `*.vcxproj.user`

**"Debugger to launch" option name**:
- [Local Debugger](#local-windows-debugger)
- [Remote Debugger](#remote-windows-debugger)
- [Web Browser Debugger](#web-browser-debugger)
- [Web Service Debugger](#web-service-debugger)

---

# Local Windows Debugger

**Source**: C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Microsoft\VC\v170\1033\debugger_local_windows.xml

| **Property**                                                              | **Display**                              | **Type**     | **Description**                                                                                        |
|---------------------------------------------------------------------------|------------------------------------------|--------------|--------------------------------------------------------------------------------------------------------|
| [LocalDebuggerCommand](#localdebuggercommand)                             | Command                                  | String       | The debug command to execute.                                                                          |
| [LocalDebuggerCommandArguments](#localdebuggercommandarguments)           | Command Arguments                        | String       | The command line arguments to pass to the application.                                                 |
| [LocalDebuggerWorkingDirectory](#localdebuggerworkingdirectory)           | Working Directory                        | String       | The application's working directory. By default, the directory containing the project file.            |
| [LocalDebuggerAttach](#localdebuggerattach)                               | Attach                                   | Bool         | Specifies whether the debugger should attempt to attach to an existing process when debugging starts.  |
| [LocalDebuggerDebuggerType](#localdebuggerdebuggertype)                   | Debugger Type                            | Enum         | Specifies the debugger type to use. When set to Auto, the debugger type is selected based on the exe.  |
| [LocalDebuggerEnvironment](#localdebuggerenvironment)                     | Environment                              | String List  | Specifies the environment for the debugee, or variables to merge with the existing environment.        |
| [LocalGPUDebuggerTargetType](#localgpudebuggertargettype)                 | Debugging Accelerator Type               | Enum         | The debugging accelerator type to use for debugging GPU code (available when GPU debugger is active).  |
| [GPURefDebuggerBreakOnAllThreads](#gpurefdebuggerbreakonallthreads)       | GPU Default Breakpoint Behavior          | Enum         | Sets how often the GPU debugger breaks.                                                                |
| [LocalDebuggerMergeEnvironment](#localdebuggermergeenvironment)           | Merge Environment                        | Bool         | Merge specified environment variables with the existing environment.                                   |
| [LocalDebuggerSQLDebugging](#localdebuggersqldebugging)                   | SQL Debugging                            | Bool         | Attach the SQL debugger.                                                                               |
| [LocalDebuggerAmpDefaultAccelerator](#localdebuggerampdefaultaccelerator) | Amp Default Accelerator                  | Enum         | Override C++ AMP's default accelerator selection (not applicable for debugging managed code).          |
| [UseLegacyManagedDebugger](#uselegacymanageddebugger)                     | ???                                      | Bool         | ??? (Hidden)                                                                                           |

[Back to top](#debugging-property-page)

---

## LocalDebuggerCommand

**Display**: Command  
**Description**: The debug command to execute.  
**Type**: String  

| **Dropdown Option**               | **Display**                   |
|-----------------------------------|-------------------------------|
| DefaultFindFullPathPropertyEditor | `<regsvr32.exe>`              |
| DefaultStringPropertyEditor       | "Edit..."                     |
| DefaultFilePropertyEditor         | "Browse..."                   |

[Back to Local Windows Debugger](#local-windows-debugger)

---

## LocalDebuggerCommandArguments

**Display**: Command Arguments  
**Description**: The command line arguments to pass to the application.  
**Type**: String  

[Back to Local Windows Debugger](#local-windows-debugger)

---

## LocalDebuggerWorkingDirectory

**Display**: Working Directory  
**Description**: The application's working directory. By default, the directory containing the project file.  
**Type**: String  
**Subtype**: Folder  

[Back to Local Windows Debugger](#local-windows-debugger)

---

## LocalDebuggerAttach

**Display**: Attach  
**Description**: Specifies whether the debugger should attempt to attach to an existing process when debugging starts.  
**Type**: Bool  

[Back to Local Windows Debugger](#local-windows-debugger)

---

## LocalDebuggerDebuggerType

**Display**: Debugger Type  
**Description**: Specifies the debugger type to use. When set to Auto, the debugger type is selected based on the exe.  
**Type**: Enum  

| **Option**              | **Display**                   |
|-------------------------|-------------------------------|
| NativeOnly              | Native Only                   |
| ManagedOnly             | Managed Only (.NET Framework) |
| Mixed                   | Mixed (.NET Framework)        |
| ManagedCore             | Managed Only (.NET Core)      |
| NativeWithManagedCore   | Mixed (.NET Core)             |
| Auto                    | Auto                          |
| Script                  | Script                        |
| GPUOnly                 | GPU Only (C++ AMP)            |
| JavaScriptForWebView2   | JavaScript (WebView2)         |

[Back to Local Windows Debugger](#local-windows-debugger)

---

## LocalDebuggerEnvironment

**Display**: Environment  
**Description**: Specifies the environment for the debugee, or variables to merge with the existing environment.  
**Type**: String List  
**Separator**: `\n` (Newline)  

[Back to Local Windows Debugger](#local-windows-debugger)

---

## LocalGPUDebuggerTargetType

**Display**: Debugging Accelerator Type  
**Description**: The debugging accelerator type to use for debugging GPU code (available when GPU debugger is active).  
**Type**: Enum  
**Active if**: LocalDebuggerDebuggerType == GPUOnly  
**Category**: GPUOptions   
**Enum Provider**: GPUTargets **(TODO)**

[Back to Local Windows Debugger](#local-windows-debugger)

---

## GPURefDebuggerBreakOnAllThreads

**Display**: GPU Default Breakpoint Behavior  
**Description**: Sets how often the GPU debugger breaks.  
**Type**: Enum  

| **Option**              | **Display**                                |
|-------------------------|--------------------------------------------|
| GPURefBreakOncePerWarp  | Break once per warp                        |
| GPURefBreakForCPUThread | Break for every thread (like CPU behavior) |

[Back to Local Windows Debugger](#local-windows-debugger)

---

## LocalDebuggerMergeEnvironment

**Display**: Merge Environment  
**Description**: Merge specified environment variables with the existing environment.  
**Type**: Bool  

[Back to Local Windows Debugger](#local-windows-debugger)

---

## LocalDebuggerSQLDebugging

**Display**: SQL Debugging  
**Description**: Attach the SQL debugger.  
**Type**: Bool  

[Back to Local Windows Debugger](#local-windows-debugger)

---

## LocalDebuggerAmpDefaultAccelerator

**Display**: Amp Default Accelerator  
**Description**: Override C++ AMP's default accelerator selection (not applicable for debugging managed code).  
**Type**: Enum  

[Back to Local Windows Debugger](#local-windows-debugger)

---

## UseLegacyManagedDebugger

**Display**: ???  
**Description**: ???  
**Type**: Bool  
**Visible**: False  

[Back to Local Windows Debugger](#local-windows-debugger)

---

# Remote Windows Debugger

**Source**: C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Microsoft\VC\v170\1033\debugger_remote_windows.xml

| **Property**                                                                | **Display**                               | **Type**     | **Description**                                                                                        |
|-----------------------------------------------------------------------------|-------------------------------------------|--------------|--------------------------------------------------------------------------------------------------------|
| [RemoteDebuggerCommand](#remotedebuggercommand)                             | Remote Command                            | String       | The debug command to execute.                                                                          |
| [RemoteDebuggerCommandArguments](#remotedebuggercommandarguments)           | Remote Command Arguments                  | String       | The command line arguments to pass to the application.                                                 |
| [RemoteDebuggerWorkingDirectory](#remotedebuggerworkingdirectory)           | Working Directory                         | String       | The application's working directory. By default, the directory containing the project file.            |
| [RemoteDebuggerServerName](#remotedebuggerservername)                       | Remote Server Name                        | String       | Specifies a remote server name.                                                                        |
| [RemoteDebuggerConnection](#remotedebuggerconnection)                       | Connection                                | Enum         | Specifies the connection type.                                                                         |
| [RemoteDebuggerDebuggerType](#remotedebuggerdebuggertype)                   | Debugger Type                             | Enum         | Specifies the debugger type to use. When set to Auto, the debugger type is selected based on the exe.  |
| [RemoteDebuggerEnvironment](#remotedebuggerenvironment)                     | Environment                               | String List  | Specifies the environment for the debugee, or variables to merge with the existing environment.        |
| [RemoteGPUDebuggerTargetType](#remotegpudebuggertargettype)                 | Debugging Accelerator Type                | Enum         | The debugging accelerator type to use for debugging GPU code (available when GPU debugger is active).  |
| [GPURefDebuggerBreakOnAllThreads](#gpurefdebuggerbreakonallthreads)         | GPU Default Breakpoint Behavior           | Enum         | Sets how often the GPU debugger breaks.                                                                |
| [RemoteDebuggerAttach](#remotedebuggerattach)                               | Attach                                    | Bool         | Specifies whether the debugger should attempt to attach to an existing process when debugging starts.  |
| [RemoteDebuggerSQLDebugging](#remotedebuggersqldebugging)                   | SQL Debugging                             | Bool         | Attach the SQL debugger.                                                                               |
| [DeploymentDirectory](#deploymentdirectory)                                 | Deployment Directory                      | String       | Specifies the directory to deploy the project output to when debugging on a remote machine.            |
| [AdditionalFiles](#additionalfiles)                                         | Additional Files to Deploy                | String       | Specifies additional files or directories to deploy to the remote machine.                            |
| [RemoteDebuggerDeployDebugCppRuntime](#remotedebuggerdeploydebugcppruntime) | Deploy Visual C++ Debug Runtime Libraries | Bool         | Specifies whether to deploy the debug runtime libraries for the active platform.                       |
| [RemoteDebuggerDeployCppRuntime](#remotedebuggerdeploycppruntime)           | Deploy Visual C++ Runtime Libraries       | Bool         | Specifies whether to deploy the runtime libraries for the active platform.                             |
| [RemoteDebuggerAmpDefaultAccelerator](#remotedebuggerampdefaultaccelerator) | Amp Default Accelerator                   | Enum         | Override C++ AMP's default accelerator selection (not applicable for debugging managed code).          |
| [UseLegacyManagedDebugger](#uselegacymanageddebugger)                       | ???                                       | Bool         | ??? (Hidden)                                                                                           |

[Back to top](#debugging-property-page)

---

## RemoteDebuggerCommand

**Display**: Remote Command  
**Description**: The debug command to execute.  
**Type**: String  

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## RemoteDebuggerCommandArguments

**Display**: Remote Command Arguments  
**Description**: The command line arguments to pass to the application.  
**Type**: String  

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## RemoteDebuggerWorkingDirectory

**Display**: Working Directory  
**Description**: The application's working directory. By default, the directory containing the project file.  
**Type**: String  

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## RemoteDebuggerServerName

**Display**: Remote Server Name  
**Description**: Specifies a remote server name.  
**Type**: String  

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## RemoteDebuggerConnection

**Display**: Connection  
**Description**: Specifies the connection type.  
**Type**: Enum  

| **Option**                     | **Display**                           |
|--------------------------------|---------------------------------------|
| RemoteWithAuthentication       | Remote with Windows authentication    |
| RemoteWithoutAuthentication    | Remote with no authentication         |

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## RemoteDebuggerDebuggerType

**Display**: Debugger Type  
**Description**: Specifies the debugger type to use. When set to Auto, the debugger type is selected based on the exe.  
**Type**: Enum  

| **Option**              | **Display**                   |
|-------------------------|-------------------------------|
| NativeOnly              | Native Only                   |
| ManagedOnly             | Managed Only (.NET Framework) |
| Mixed                   | Mixed (.NET Framework)        |
| ManagedCore             | Managed Only (.NET Core)      |
| NativeWithManagedCore   | Mixed (.NET Core)             |
| Auto                    | Auto                          |
| Script                  | Script                        |
| GPUOnly                 | GPU Only (C++ AMP)            |
| JavaScriptForWebView2   | JavaScript (WebView2)         |

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## RemoteDebuggerEnvironment

**Display**: Environment  
**Description**: Specifies the environment for the debugee, or variables to merge with the existing environment.  
**Type**: String List  
**Separator**: `\n` (Newline)  

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## RemoteGPUDebuggerTargetType

**Display**: Debugging Accelerator Type  
**Description**: The debugging accelerator type to use for debugging GPU code (available when GPU debugger is active).  
**Type**: Enum  

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## GPURefDebuggerBreakOnAllThreads

**Display**: GPU Default Breakpoint Behavior  
**Description**: Sets how often the GPU debugger breaks.  
**Type**: Enum  

| **Option**              | **Display**                                |
|-------------------------|--------------------------------------------|
| GPURefBreakOncePerWarp  | Break once per warp                        |
| GPURefBreakForCPUThread | Break for every thread (like CPU behavior) |

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## RemoteDebuggerAttach

**Display**: Attach  
**Description**: Specifies whether the debugger should attempt to attach to an existing process when debugging starts.  
**Type**: Bool  

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## RemoteDebuggerSQLDebugging

**Display**: SQL Debugging  
**Description**: Attach the SQL debugger.  
**Type**: Bool  

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## DeploymentDirectory

**Display**: Deployment Directory  
**Description**: Specifies the directory to deploy the project output to when debugging on a remote machine.  
**Type**: String  

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## AdditionalFiles

**Display**: Additional Files to Deploy  
**Description**: Specifies additional files or directories to deploy to the remote machine.  
**Type**: String  

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## RemoteDebuggerDeployDebugCppRuntime

**Display**: Deploy Visual C++ Debug Runtime Libraries  
**Description**: Specifies whether to deploy the debug runtime libraries for the active platform.  
**Type**: Bool  

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## RemoteDebuggerDeployCppRuntime

**Display**: Deploy Visual C++ Runtime Libraries  
**Description**: Specifies whether to deploy the runtime libraries for the active platform.  
**Type**: Bool  

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## RemoteDebuggerAmpDefaultAccelerator

**Display**: Amp Default Accelerator  
**Description**: Override C++ AMP's default accelerator selection (not applicable for debugging managed code).  
**Type**: Enum  

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

## UseLegacyManagedDebugger

**Display**: ???  
**Description**: ???  
**Type**: Bool  
**Visible**: False  

[Back to Remote Windows Debugger](#remote-windows-debugger)

---

# Web Browser Debugger

**Source**: C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Microsoft\VC\v170\1033\debugger_web_browser.xml

| **Property**                                                      | **Display**      | **Type**     | **Description**                                                                                            |
|-------------------------------------------------------------------|------------------|--------------|------------------------------------------------------------------------------------------------------------|
| [WebBrowserDebuggerHttpUrl](#webbrowserdebuggerhttpurl)           | HTTP URL         | String       | Specifies the URL for the project.                                                                         |
| [WebBrowserDebuggerDebuggerType](#webbrowserdebuggerdebuggertype) | Debugger Type    | Enum         | Specifies the debugger type to use. When set to Auto, the debugger type is selected based on the exe file. |

[Back to top](#debugging-property-page)

---

## WebBrowserDebuggerHttpUrl

**Display**: HTTP URL  
**Description**: Specifies the URL for the project.  
**Type**: String  

[Back to Web Browser Debugger](#web-browser-debugger)

---

## WebBrowserDebuggerDebuggerType

**Display**: Debugger Type  
**Description**: Specifies the debugger type to use. When set to Auto, the debugger type is selected based on the exe file.  
**Type**: Enum  

| **Option**       | **Display**  |
|------------------|--------------|
| NativeOnly       | Native Only  |
| ManagedOnly      | Managed Only |
| Mixed            | Mixed        |
| Auto             | Auto         |
| Script           | Script       |

[Back to Web Browser Debugger](#web-browser-debugger)

---

# Web Service Debugger

**Source**: C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Microsoft\VC\v170\1033\debugger_web_service.xml

| **Property**                                                      | **Display**      | **Type**     | **Description**                                                                                            |
|-------------------------------------------------------------------|------------------|--------------|------------------------------------------------------------------------------------------------------------|
| [WebServiceDebuggerHttpUrl](#webservicedebuggerhttpurl)           | HTTP URL         | String       | Specifies the URL for the project.                                                                         |
| [WebServiceDebuggerDebuggerType](#webservicedebuggerdebuggertype) | Debugger Type    | Enum         | Specifies the debugger type to use. When set to Auto, the debugger type is selected based on the exe file. |
| [WebServiceDebuggerSQLDebugging](#webservicedebuggersqldebugging) | SQL Debugging    | Bool         | Attach the SQL debugger.                                                                                   |
| [UseLegacyManagedDebugger](#uselegacymanageddebugger)             | ???              | Bool         | ??? (Hidden)                                                                                               |

[Back to top](#debugging-property-page)

---

## WebServiceDebuggerHttpUrl

**Display**: HTTP URL  
**Description**: Specifies the URL for the project.  
**Type**: String  

[Back to Web Service Debugger](#web-service-debugger)

---

## WebServiceDebuggerDebuggerType

**Display**: Debugger Type  
**Description**: Specifies the debugger type to use. When set to Auto, the debugger type is selected based on the exe file.  
**Type**: Enum  

| **Option**       | **Display**  |
|------------------|--------------|
| NativeOnly       | Native Only  |
| ManagedOnly      | Managed Only |
| Mixed            | Mixed        |
| Auto             | Auto         |
| Script           | Script       |

[Back to Web Service Debugger](#web-service-debugger)

---

## WebServiceDebuggerSQLDebugging

**Display**: SQL Debugging  
**Description**: Attach the SQL debugger.  
**Type**: Bool  

[Back to Web Service Debugger](#web-service-debugger)

---

## UseLegacyManagedDebugger

**Display**: ???  
**Description**: ???  
**Type**: Bool  
**Visible**: False  

[Back to Web Service Debugger](#web-service-debugger)
