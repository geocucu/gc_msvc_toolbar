#include "common.h"

#include "vsix.h"
#include "..\res\guids.h"
#include "..\res\resource.h"
#include "debug_helper.h"

// ================================ UTILS ================================
struct command_args_t {
	VARIANT in;

	command_args_t() {
		VariantInit(&in);
	}

	command_args_t(VARIANT *raw_in) {
		VariantInit(&in);
		VariantCopy(&in, raw_in);
	}

	~command_args_t() {
		VariantClear(&in);
	}
};

static inline bool vsproj_is_cpp(vsix::IVsHierarchy *proj) {
  // THREAD: VS Main

  VARIANT type;
  VariantInit(&type);
  proj->GetProperty(VSITEMID_ROOT, vsix::VSHPROPID::VSHPROPID_ProjectType, &type);
  bool is_cpp = wcsstr(type.bstrVal, L"C++") != 0;
  VariantClear(&type);
  return is_cpp;
}

// ================================ DATA ================================
struct {
	vsix::IServiceProvider *provider;

  vsix::IVsUIShell *vs_ui_shell;
	HWND main_hwnd;
	WNDPROC main_wndproc;
} vs_services;

// ================================ COMMANDLINE ARGS COMBO BOX ================================
#define WM_CMDLINE_ARGS (WM_USER + 1)

HRESULT on_cmdline_args(command_args_t *args) {
  // THREAD: VS Main

	vsix::IVsSolution *sln = 0;
	HRCALL(vs_services.provider->QueryService(&sln));

	vsix::IVsSolutionBuildManager *sln_build_manager = 0;
	HRCALL(vs_services.provider->QueryService(&sln_build_manager));

	vsix::IVsHierarchy *startup_project = 0;
	HRCALL(sln_build_manager->get_StartupProject(&startup_project));

	VARIANT name;
	VariantInit(&name);
	HRCALL(startup_project->GetProperty(VSITEMID_ROOT, vsix::VSHPROPID::VSHPROPID_Name, &name));
	bool is_cpp = vsproj_is_cpp(startup_project);

	vsix::IVsProjectCfg *active_project_cfg = 0;
	HRCALL(sln_build_manager->FindActiveProjectCfg(0, 0, startup_project, &active_project_cfg));

	BSTR active_project_cfg_name;
	active_project_cfg->get_DisplayName(&active_project_cfg_name); // E.g. Debug|x64

	vsix::IVsBuildPropertyStorage *props = 0;
	HRCALL(startup_project->QueryInterface(vsix::IID_IVsBuildPropertyStorage, (void **)&props));

	HRCALL(props->SetPropertyValue(
		L"LocalDebuggerCommandArguments",
		active_project_cfg_name,
		vsix::PST_USER_FILE,
		args->in.bstrVal));

	VariantClear(&name);
	SysFreeString(active_project_cfg_name);
	active_project_cfg->Release();
	props->Release();
	startup_project->Release();
	sln_build_manager->Release();
	sln->Release();

  return S_OK;
}

HRESULT on_update_sln_state() {
  // THREAD: VS Main
  
  // Set the combo box to the Startup Project's commandline args.
  // If the project is not a C++ project, set the combo box to "N/A".
	
	vsix::IVsSolution *sln = 0;
  HRCALL(vs_services.provider->QueryService(&sln));
  vsix::IVsSolutionBuildManager *sln_build_manager = 0;
  HRCALL(vs_services.provider->QueryService(&sln_build_manager));
  vsix::IVsHierarchy *startup_project = 0;
  HRCALL(sln_build_manager->get_StartupProject(&startup_project));

  bool is_cpp = vsproj_is_cpp(startup_project);

	BSTR cmdline_args;
	if (is_cpp) {
		vsix::IVsProjectCfg *active_project_cfg = 0;
		HRCALL(sln_build_manager->FindActiveProjectCfg(0, 0, startup_project, &active_project_cfg));
		BSTR active_project_cfg_name;
		active_project_cfg->get_DisplayName(&active_project_cfg_name); // E.g. Debug|x64
		vsix::IVsBuildPropertyStorage *props = 0;
		HRCALL(startup_project->QueryInterface(vsix::IID_IVsBuildPropertyStorage, (void **)&props));

    HRCALL(props->GetPropertyValue(
      L"LocalDebuggerCommandArguments",
      active_project_cfg_name,
      vsix::PST_USER_FILE,
      &cmdline_args));

    SysFreeString(active_project_cfg_name);
    active_project_cfg->Release();
    props->Release();
	}
	else {
    cmdline_args = SysAllocString(L"N/A");
	}

	vs_services.vs_ui_shell->SetMRUComboTextW(&CLSID_cmdset, CMDID_cmdline_args_control, cmdline_args, true);

  SysFreeString(cmdline_args);
  startup_project->Release();
  sln_build_manager->Release();
  sln->Release();

  return S_OK;
}

// ================================ BOILERPLATE CONNECTION POINTS (START HERE) ================================
#pragma region

#define TIMER_UPDATE_SLN_STATE 0x1000000

LRESULT __stdcall VS_Main_WndProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam) {
	// THREAD: VS Main

	if (msg == WM_TIMER && wparam == TIMER_UPDATE_SLN_STATE) {
		on_update_sln_state();
		return 0;
	}

  if (msg > WM_USER) {
    command_args_t *args = (command_args_t *)lparam;

		bool handled = true;
    if (msg == WM_CMDLINE_ARGS) {
      on_cmdline_args(args);
    }
    // else if (msg == ...) {
    //   ...
    // }
		else {
			handled = false;
		}

    if (handled) {
      delete args;
			return 0;	
    }
		else {
      return CallWindowProcW(vs_services.main_wndproc, hwnd, msg, wparam, lparam);
		}
  }

  return CallWindowProcW(vs_services.main_wndproc, hwnd, msg, wparam, lparam);
}

static HRESULT pkg_command_map(const GUID &cmdset_id, DWORD cmd_id, DWORD flags, VARIANT *in, VARIANT *out) {
	// THREAD: VS Main OR <some other> 
	//         ^^^ can only go wild here 

	command_args_t *args = new command_args_t(in);
	bool post_ok = true;
	if (InlineIsEqualGUID(cmdset_id, CLSID_cmdset) && cmd_id == CMDID_cmdline_args_control) {
    post_ok = PostMessageW(vs_services.main_hwnd, WM_CMDLINE_ARGS, 0, (LPARAM)args);
	}
	// else if (IsInlineEqualGUID(cmdset_id, ...) && cmd_id == ...) {
	//   ...
	// }

  if (!post_ok) {
    delete args;
    return HRESULT_FROM_WIN32(GetLastError());
  }

	return OLECMDERR_E_NOTSUPPORTED;
}

// Init VSPackage
// https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.shell.interop.ivspackage.setsite?view=visualstudiosdk-2022
static HRESULT pkg_on_startup() {
	// THREAD: VS Main

  // Do non-standard init stuff here...

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

		HRCALL(service_provider->QueryService(&vs_services.vs_ui_shell));
		HRCALL(vs_services.vs_ui_shell->GetDialogOwnerHwnd(&vs_services.main_hwnd));
		// TODO: Check if it would be better to take over only temporarily:
		//       Restore original after processing, then take over again whenever needed.
		vs_services.main_wndproc = (WNDPROC)SetWindowLongPtrW(vs_services.main_hwnd, GWLP_WNDPROC, (LONG_PTR)VS_Main_WndProc);

		// The Advise...Event functions are a mess.
		// Just query the sln state at coarse regular intervals, say 5 seconds.
    SetTimer(vs_services.main_hwnd, TIMER_UPDATE_SLN_STATE, 5000, 0);

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