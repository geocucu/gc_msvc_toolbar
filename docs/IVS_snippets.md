## Service examples

`IServiceProvider *service_provider` comes either from the VS Site cache or as the param to `SetSite` (later cached).

### IVsSolution

#### 1. Get the service  

```C++
IVsSolution *solution;
service_provider->QueryService(IID_IVsSolution, IID_IVsSolution, (void **)&solution);
```

#### 2. Enumerate loaded projects (`IVsHierarchy*`)  

```C++
IEnumHierarchies *enum_hierarchies;
solution->GetProjectEnum(__VSENUMPROJFLAGS::EPF_LOADEDINSOLUTION, GUID_NULL, &enum_hierarchies);

IVsHierarchy *hierarchy = 0;
ULONG fetched = 0;
while (enum_hierarchies->Next(1, &hierarchy, &fetched) == S_OK && fetched > 0) {
	// Do something with the hierarchy instance (project)
	CComVariant projname;
	hierarchy->GetProperty(VSITEMID_ROOT, VSHPROPID_Name, &projname);
	MessageBoxW(0, projname.bstrVal, L"Found project", MB_OK);

	hierarchy->Release();
}
```

### IVsSolutionBuildManager

#### 1. Get the service

```C++
IVsSolutionBuildManager *build_manager;
service_provider->QueryService(IID_IVsSolutionBuildManager, IID_IVsSolutionBuildManager, (void **)&build_manager);
```

#### 2. Query the startup project

```C++
IVsHierarchy *startup_proj;
build_manager->get_StartupProject(&startup_proj);
```

#### 3. Query a property of the project

Argument 2 uses the `VSHPROPID_[...]` constants, with self-explanatory names.

```C++
CComVariant proj_name;
startup_proj->GetProperty(VSITEMID_ROOT, VSHPROPID_Name, &proj_name);
```

#### 4. Query a configuration property. 

Neither `Microsoft.VisualStudio.VCProject.xml` nor `Microsoft.VisualStudio.VCProjectEngine.xml` contain all possible names. 
Have to pick from the generated `*.vcxproj` or `*.vcxproj.user`, or straight from the MSBuild config XMLs in the VS install dir.  

Argument 2 (`pszConfigName`) is actually `config_name|platform_name`, with the pipe separator and no spaces.  

Argument 3 (`storage`) is either `PST_PROJECT_FILE` or `PST_USER_FILE`.
Corresponds to either the `*.vcxproj` or `*.vcxproj.user` files.
It doesn't seem to matter (just passing `PST_PROJECT_FILE` is fine).
However, there are corresponding tags in the MSBuild config XMLs:

```XML
<Rule.DataSource> 
  <DataSource Persistence="UserFile" /> 
</Rule.DataSource>
```

Example:  

```C++
IVsBuildPropertyStorage *prop_storage;
startup_proj->QueryInterface(__uuidof(IVsBuildPropertyStorage), (void **)&prop_storage);

CComBSTR original_value;
prop_storage->GetPropertyValue(
	L"LocalDebuggerCommandArguments",
	L"Debug|x64",
	PST_PROJECT_FILE,
	&original_value
);
```

5. Marshaling 

```C++
// IStream *stream;
CoMarshalInterThreadInterfaceInStream(IID_IVsHierarchy, startup_proj, &stream);
```