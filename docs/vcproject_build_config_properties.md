# Local debugger 

**Source**: C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Microsoft\VC\v170\1033\debugger_local_windows.xml

**Persistence**: `PST_USER_FILE` -> `*.vcxproj.user`

**"Debugger to launch" option name**: Local Windows Debugger

## LocalDebuggerCommand

**Display**: Command  
**Description**: The debug command to execute.  
**Type**: String  

| **Dropdown Option**               | **Display**                   |
|-----------------------------------|-------------------------------|
| DefaultFindFullPathPropertyEditor | <regsvr32.exe>                |
| DefaultStringPropertyEditor       | "Edit..."                     |
| DefaultFilePropertyEditor         | "Browse..."                   |

## LocalDebuggerCommandArguments

**Display**: Command Arguments  
**Description**: The command line arguments to pass to the application.  
**Type**: String

## LocalDebuggerWorkingDirectory

**Display**: Working Directory  
**Description**: The application's working directory. By default, the directory containing the project file.  
**Type**: String  
**Subtype**: Folder

## LocalDebuggerAttach

**Display**: Attach  
**Description**: Specifies whether the debugger should attempt to attach to an existing process when debugging starts.  
**Type**: Bool

## LocalDebuggerDebuggerType

**Display**: Debugger Type  
**Description**: Specifies the debugger type to use. When set to Auto, the debugger type will be selected based on contents of the exe file.  
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

## LocalDebuggerEnvironment

**Display**: Environment  
**Description**: Specifies the environment for the debugee, or variables to merge with existing environment.  
**Type**: String List  
**Separator**: `\n` (Newline)   

## LocalGPUDebuggerTargetType

**Display**: Debugging Accelerator Type  
**Description**: The debugging accelerator type to use for debugging the GPU code.  (Available when the GPU debugger is active.)  
**Type**: Enum  
**Active if**: LocalDebuggerDebuggerType == GPUOnly  
**Category**: GPUOptions   
**Enum Provider**: GPUTargets **(TODO)**

## GPURefDebuggerBreakOnAllThreads

**Display**: GPU Default Breakpoint Behavior   
**Description**: Sets how often the GPU debugger breaks.   
**Type**: Enum  
**Active if**: LocalDebuggerDebuggerType == GPUOnly  
**Category**: GPUOptions  

| **Option**              | **Display**                                |  
|-------------------------|--------------------------------------------| 
| GPURefBreakOncePerWarp  | Break once per warp                        | 
| GPURefBreakForCPUThread | Break for every thread (like CPU behavior) |

## LocalDebuggerMergeEnvironment

**Display**: Merge Environment      
**Description**: Merge specified environment variables with existing environment.  
**Type**: Bool  

## LocalDebuggerSQLDebugging

**Display**: SQL Debugging  
**Description**: Attach the SQL debugger.  
**Type**: Bool  

## LocalDebuggerAmpDefaultAccelerator

**Display**: Amp Default Accelerator  
**Description**: Override C++ AMP's default accelerator selection. Property does not apply when debugging managed code.  
**Type**: Enum  
**Enum Provider**: AmpAccelerators **(TODO)**

## UseLegacyManagedDebugger

**Display**: ???  
**Description**: ???  
**Type**: Bool  
**Visible**: False