#include "common.h"

// TODO: FIX LIFETIMES 

#include "vsix.h"
#include "..\res\guids.h"
#include "..\res\resource.h"
#include "debug_helper.h"

// ======== COM Utils ======== 
#pragma region

static IStream* com_duplicate_stream(IStream *src) {
	IStream *new_stream;
	HRESULT hr = S_OK;

	hr = CreateStreamOnHGlobal(0, TRUE, &new_stream);
	if (FAILED(hr)) {
		msgbox_hresult(hr, L"Failed to create an IStream.");
		return 0;
	}

	LARGE_INTEGER seek_start = {};
	ULARGE_INTEGER stream_size = {};
	src->Seek(seek_start, STREAM_SEEK_END, &stream_size);
	src->Seek(seek_start, STREAM_SEEK_SET, 0);

	hr = src->CopyTo(new_stream, stream_size, 0, 0);
	if (FAILED(hr)) {
		msgbox_hresult(hr, L"Failed to copy to duplicate stream");
		return 0;
	}

	src->Seek(seek_start, STREAM_SEEK_SET, 0);
	new_stream->Seek(seek_start, STREAM_SEEK_SET, 0);

	return new_stream;
}

#pragma endregion

// ================================ DATA ================================
struct {
	vsix::IServiceProvider *provider;

  vsix::IVsSolutionBuildManager *sln_build_manager;
  vsix::IVsHierarchy *startup_project;

  vsix::IVsUIShell *vs_ui_shell;
	HWND main_hwnd;
	WNDPROC main_wndproc;
} vs_services;

struct project_t {
	wchar_t *name;
	IStream *stream;
};

struct {
  project_t projects[64];
	int num_projects = 0;

	int startup_project_index = -1;
} sln_state;

// ================================ COMMANDLINE ARGS COMBO BOX ================================
#define WM_CMDLINE_ARGS (WM_USER + 1)

LRESULT __stdcall VS_Main_WndProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam) {
	// THREAD: VS Main

	switch (msg) {

	case WM_CMDLINE_ARGS: {
		VARIANT *in = reinterpret_cast<VARIANT *>(lparam);

		// Neither DTE, nor IVs* based get_StartupProject here will work, even with marshaling.
		// Can't reliably execute on the UI thread either...
		// So the project list is fixed at initialisation, along with the startup project. 

		vsix::IVsProjectCfg *active_project_cfg = 0;
		HRCALL(vs_services.sln_build_manager->FindActiveProjectCfg(0, 0, vs_services.startup_project, &active_project_cfg));

		BSTR active_project_cfg_name;
		active_project_cfg->get_DisplayName(&active_project_cfg_name); // Debug|x64

		// Do not release!
		vsix::IVsHierarchy *proj = 0;
		HRCALL(CoGetInterfaceAndReleaseStream(com_duplicate_stream(sln_state.projects[0].stream), vsix::IID_IVsHierarchy, (void **)&proj));

		vsix::IVsBuildPropertyStorage *props = 0;
		HRCALL(proj->QueryInterface(vsix::IID_IVsBuildPropertyStorage, (void **)&props));

		// == TODO: HARDCODED == 
		HRCALL(props->SetPropertyValue(
			L"LocalDebuggerCommandArguments",
			L"Release|x64",
			vsix::PST_USER_FILE,
			in->bstrVal));

		// Restore WNDPROC
		SetWindowLongPtrW(vs_services.main_hwnd, GWLP_WNDPROC, (LONG_PTR)vs_services.main_wndproc);

		delete in;
		return 0;
	}

	default:
		return DefWindowProc(hwnd, msg, wparam, lparam);
	}
}

static HRESULT on_cmdline_args(VARIANT *in) {
  // THREAD: VS Main OR <some other>

	vs_services.main_wndproc = (WNDPROC)SetWindowLongPtrW(vs_services.main_hwnd, GWLP_WNDPROC, (LONG_PTR)VS_Main_WndProc);

	VARIANT *var = new VARIANT;
	VariantInit(var);
	VariantCopy(var, in);
	if (!PostMessageW(vs_services.main_hwnd, WM_CMDLINE_ARGS, 0, reinterpret_cast<LPARAM>(var))) {
		delete var;
		return HRESULT_FROM_WIN32(GetLastError());
	}

  return S_OK;
}

// ================================ BOILERPLATE CONNECTION POINTS (START HERE) ================================
#pragma region

static HRESULT pkg_command_map(const GUID &cmdset_id, DWORD cmd_id, DWORD flags, VARIANT *in, VARIANT *out) {
	// THREAD: VS Main OR <some other> 
	//         ^^^ can only go wild here 

	if (InlineIsEqualGUID(cmdset_id, CLSID_cmdset) && cmd_id == CMDID_cmdline_args_control) {
		on_cmdline_args(in);
	}

	return OLECMDERR_E_NOTSUPPORTED;
}

// Init VSPackage
// https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.shell.interop.ivspackage.setsite?view=visualstudiosdk-2022
static HRESULT pkg_on_startup() {
	// THREAD: VS Main

	HRCALL(vs_services.provider->QueryService(&vs_services.vs_ui_shell));
	HRCALL(vs_services.vs_ui_shell->GetDialogOwnerHwnd(&vs_services.main_hwnd));

	vsix::IVsSolution *sln = 0;
	HRCALL(vs_services.provider->QueryService(&sln));

	// Save the startup project name for comparison within the enumeration later.
	// Any failure here is critical, since no updates to the solution are tracked.
	VARIANT startup_project_name;
  VariantInit(&startup_project_name);
	{
		vsix::IVsSolutionBuildManager *sln_build_mgr = 0;
		HRCALL(vs_services.provider->QueryService(&sln_build_mgr));
		
		vsix::IVsHierarchy *startup_project = 0;
    HRCALL(sln_build_mgr->get_StartupProject(&startup_project));

		VARIANT project_type;
		VariantInit(&project_type); 

		startup_project->GetProperty(VSITEMID_ROOT, vsix::VSHPROPID::VSHPROPID_Name, &startup_project_name);
		startup_project->GetProperty(VSITEMID_ROOT, vsix::VSHPROPID::VSHPROPID_ProjectType, &project_type);

    if (wcsstr(project_type.bstrVal, L"C++") == 0) {
      msgbox_hresult(E_FAIL, L"Startup project is not a C++ project.");
      return E_FAIL;
    }

		VariantClear(&project_type);
	}

	// Enumerate all projects in the sln, using IVs* interfaces
  // Store an index to the startup project by matching names with the previously found startup project 
	// Any updates to the solution's projects in this session will be ignored at the moment, including: add, rename, remove, change startup
	vsix::IEnumHierarchies *hierarchies = 0;
	HRCALL(sln->GetProjectEnum(vsix::VSENUMPROJFLAGS::EPF_ALLPROJECTS, GUID_NULL, &hierarchies));

	vsix::IVsHierarchy *proj = 0;
	ULONG fetched = 0;
	while (hierarchies->Next(1, &proj, &fetched) == S_OK && fetched > 0) {
		if (!proj) {
			continue;
		}

    // Fixed limit of 64 projects
    if (sln_state.num_projects >= 64) {
      MessageBoxW(0, L"Too many projects in the solution. Max supported: 64.", L"[Warning] " PROJECT_NAME, MB_ICONWARNING | MB_OK);
      break;
    }

		VARIANT name;
    VariantInit(&name);
		proj->GetProperty(VSITEMID_ROOT, vsix::VSHPROPID::VSHPROPID_ProjectName, &name);

		// Ignore non-C++ projects
		VARIANT type;
    VariantInit(&type);
		proj->GetProperty(VSITEMID_ROOT, vsix::VSHPROPID::VSHPROPID_ProjectType, &type);
		if (wcsstr(type.bstrVal, L"C++")) {
			// Serialise IVsHierarchy to IStream for GetPropertyValue/SetPropertyValue from non-UI thread calls later
			IStream *stream = 0;
			HRCALL(CoMarshalInterThreadInterfaceInStream(vsix::IID_IVsHierarchy, (IUnknown *)proj, &stream));

			sln_state.projects[sln_state.num_projects++] = {
				.name = name.bstrVal,
				.stream = stream
			};

			if (wcscmp(name.bstrVal, startup_project_name.bstrVal) == 0) {
				sln_state.startup_project_index = sln_state.num_projects - 1;
			}
    }
		else {
      VariantClear(&name);
		}

    VariantClear(&type);
		proj->Release();
  }

  if (sln_state.startup_project_index == -1) {
    msgbox_hresult(E_FAIL, L"Startup project not found in the solution.");
    return E_FAIL;
  }

  VariantClear(&startup_project_name);

	HRCALL(vs_services.provider->QueryService(&vs_services.sln_build_manager));
	HRCALL(vs_services.sln_build_manager->get_StartupProject(&vs_services.startup_project));

	return S_OK;
}

#pragma endregion 

// ================================ BOILERPLATE START (IGNORE BELOW THIS LINE) ================================
#pragma region 

struct pkg_t : vsix::IVsPackage, IOleCommandTarget {
	long refcount = 0;

	// ==== IUnknown ==== 
	HRESULT __stdcall QueryInterface(const IID &iid, void **obj) override {
		*obj = 0;
		if (InlineIsEqualGUID(iid, IID_IUnknown) || InlineIsEqualGUID(iid, vsix::IID_IVsPackage)) {
			*obj = static_cast<IVsPackage *>(this);
		}
		else if (InlineIsEqualGUID(iid, IID_IOleCommandTarget)) {
			*obj = static_cast<IOleCommandTarget *>(this);
		}
		else {
			return E_NOINTERFACE;
		}

		((IUnknown*)*obj)->AddRef();
		return S_OK;
	}

	ULONG __stdcall AddRef() override {
		return _InterlockedIncrement(&refcount);
	}

	ULONG __stdcall Release() override {
		long new_refcount = _InterlockedDecrement(&refcount);
		if (new_refcount == 0) {
			delete this;
		}
		return new_refcount;
	}

	// == IVsPackage == 
	HRESULT __stdcall SetSite(vsix::IServiceProvider *service_provider) override {
		vs_services.provider = service_provider;
		HRCALL(pkg_on_startup());

		return S_OK;
	}

	HRESULT __stdcall QueryClose(BOOL *pbCanClose) override {
		*pbCanClose = TRUE;
		return S_OK;
	}

	HRESULT __stdcall Close() override {
		return S_OK;
	}

	HRESULT __stdcall GetAutomationObject(const wchar_t *prop_name, vsix::IDispatch **dispatch) override { return E_NOTIMPL; }
	HRESULT __stdcall CreateTool(const GUID &guid_persistance_slot) override { return E_NOTIMPL; }
	HRESULT __stdcall ResetDefaults(vsix::VSPKGRESETFLAGS flags) override { return E_NOTIMPL; }
	HRESULT __stdcall GetPropertyPage(const GUID &guid_page, vsix::VSPROPSHEETPAGE *page) override { return E_NOTIMPL; }

	// == IOleCommandTarget == 
	HRESULT __stdcall QueryStatus(const GUID *pCmdGroupGuid, ULONG cCmds, OLECMD pCmds[], OLECMDTEXT *pCmdText) override { return E_NOTIMPL; }

	HRESULT __stdcall Exec(const GUID *pCmdGroupGuid, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *in, VARIANT *out) override {
		return pkg_command_map(*pCmdGroupGuid, nCmdID, nCmdexecopt, in, out);
	}

	// ==== Factory ====
	struct factory_t : IClassFactory {
		long refcount = 0;

		HRESULT __stdcall CreateInstance(IUnknown *agg_outer, const IID &iid, void **obj) override {
			if (agg_outer) {
				return CLASS_E_NOAGGREGATION;
			}

			pkg_t *p = new pkg_t;
			p->QueryInterface(iid, obj);
			return S_OK;
		}

		HRESULT __stdcall LockServer(BOOL) override { return S_OK; }

		HRESULT __stdcall QueryInterface(const IID &iid, void **obj) override { 
			if (InlineIsEqualGUID(iid, IID_IUnknown) || InlineIsEqualGUID(iid, IID_IClassFactory)) {
				*obj = static_cast<IClassFactory *>(this);
				AddRef();
				return S_OK;
			}

			*obj = 0;
			return E_NOINTERFACE;
		}
		
		ULONG __stdcall AddRef() override {
			return _InterlockedIncrement(&refcount);
		}

		ULONG __stdcall Release() override {
			long new_refcount = _InterlockedDecrement(&refcount);
			if (new_refcount == 0) {
				delete this;
			}
			return new_refcount;
		}
	};
};

_Check_return_
extern "C" HRESULT DllGetClassObject(_In_ const IID &, _In_ const IID &, _Outptr_ void **ppv) {
	pkg_t::factory_t *factory = new pkg_t::factory_t;
	factory->refcount = 1;
	*ppv = factory;
	return S_OK;
}

#pragma endregion