#### Property persistence 

The persistence type (user or project) in the property sheet config XMLs is not a hard rule for what to specify with 
get/set property methods.  
A property may appear in both the user and project files at the same time. In that case, VS will pick up whichever
file matches the "canonical" persistence.

#### Property update

Changes in the project property sheet followed by "Apply/OK" will immediately be saved to disk, to the file 
corresponding to the project's canonical persistence ("user"/"project").  

Monitoring the files directly is a possible choice for monitoring changes to the project configuration. 
Risky, but the alternative (`IVs[...]Events`) interfaces make me not care right now.  
Alternatively, try DTE for update notifications, and IVs* for property getters/setters (which are broken for DTE).

#### Startup Project

Changing it doesn't affect either the project user/project files, or the solution file.  
Most likely the hidden `.vs` dir contains some flag somewhere with the Startup Project. 
Deleting `.vs` resets the Startup Project selection.

#### Threads

The package `SetSite()` method will always be called in the main UI thread ("VS Main").  
In the non-boilerplate code, this is the `pkg_on_startup()` function. 

The package `Exec()` method may or may not be called in the "VS Main" thread, throughout a session. 
In other words, whether it's called in "VS Main" or not is consistent during the entire session.  
In the non-boilerplate code, this is the `pkg_command_map()` function. 

Using `IVsTaskSchedulerService` does not guarantee running on the "VS Main" thread, regardless of the context flag
passed to `CreateTask()`. 
It will only run there if `Exec()` is called on "VS Main", and the context flag is one of the "UI thread" ones. 

#### Extension registration

The GUIDs specified in the .pkgdef don't seem to end up in the registry.  
Extensions are probably installed (.vsix is extracted) in `%localappdata%\Microsoft\VisualStudio\<version>[Exp]\Extensions`

#### What makes an extension

The DLL exports `DllGetClassObject`.  
VS calls this expecting an `IClassFactory` interface implementation in return.  
VS will call the factory's `CreateInstance`, expecting an object that implements both `IVsPackage` and `IOleCommandTarget`, which I will call "the package".

Therefore, the minimum code for a valid extension:
- DLL exports:
  - `DllGetClassObject` (for creating the factory object)
- COM interfaces implemented:
  - `IClassFactory` (for creating the package object)
  - `IVsPackage` (for integration with VS)
  - `IOleCommandTarget` (for handling VS commands)