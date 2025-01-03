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

// Not meant to be a robust solution, just a quick and dirty way to manage COM/OLE (and related) objects in bulk, without individual RAII wrappers.
struct IUnknown_scope_t {
	IUnknown **objs;
	int count = 0;
	int capacity = 0;

	IUnknown_scope_t() {
    objs = (IUnknown **)calloc(32, sizeof(IUnknown *));
    capacity = 32;
	}

	~IUnknown_scope_t() {
		for (int i = count - 1; i >= 0; --i) {
			if (objs[i]) {
        objs[i]->Release();
			}
		}
	}

  IUnknown* next() {
    if (count == capacity) {
      capacity *= 2;
      objs = (IUnknown **)realloc(objs, capacity * sizeof(IUnknown *));
    }

    return objs[count++];
  }
};

struct VARIANT_RTTI_t : VARIANT {
  VARIANT_RTTI_t() { VariantInit(this); }
  ~VARIANT_RTTI_t() { VariantClear(this); }
};

struct BSTR_RTTI_t {
	wchar_t *val;
  BSTR_RTTI_t() { val = 0; }
	~BSTR_RTTI_t() { val ? SysFreeString(val) : void(0); }
};

static inline bool vsproj_is_cpp(vsix::IVsHierarchy *proj) {
  // THREAD: VS Main

  VARIANT_RTTI_t type;
  proj->GetProperty(VSITEMID_ROOT, vsix::VSHPROPID::VSHPROPID_ProjectType, &type);
  return wcsstr(type.bstrVal, L"C++") != 0;
}

// ================================ DATA ================================

// Safe to store for the entire session, unlike e.g. Solution/Project references.
struct {
	vsix::IServiceProvider *provider;

  vsix::IVsUIShell *vs_ui_shell;
	HWND main_hwnd;
	WNDPROC main_wndproc;
} vs_services;

// ================================ COMMANDLINE ARGS COMBO BOX ================================
// 0x0400 to 0x7FFF
#define WM_CMDLINE_ARGS (0x7000 + 1)

// First update seems okay; anything after that clears the combo box text.
wchar_t last_cmdline_args[256];

HRESULT on_cmdline_args(command_args_t *args) {
  // THREAD: VS Main

  IUnknown_scope_t com_objs;

  auto sln = (vsix::IVsSolution*)com_objs.next();
	auto sln_build_manager = (vsix::IVsSolutionBuildManager *)com_objs.next();
	auto startup_project = (vsix::IVsHierarchy *)com_objs.next();
	auto active_project_cfg = (vsix::IVsProjectCfg *)com_objs.next();
	auto props = (vsix::IVsBuildPropertyStorage *)com_objs.next();
  VARIANT_RTTI_t startup_project_name;
  BSTR_RTTI_t active_project_cfg_name;

	HRCALL(vs_services.provider->QueryService(&sln));
	HRCALL(vs_services.provider->QueryService(&sln_build_manager));
	HRCALL(sln_build_manager->get_StartupProject(&startup_project));

	HRCALL(startup_project->GetProperty(VSITEMID_ROOT, vsix::VSHPROPID::VSHPROPID_Name, &startup_project_name));
	if (!vsproj_is_cpp(startup_project)) {
    return S_OK; // No error, but nothing to do either (non-C++ project).
	}

  sln_build_manager->get_ActiveProjectCfg_name(startup_project, &active_project_cfg_name.val);

  HRCALL(startup_project->get_BuildPropertyStorage(&props));

	HRCALL(props->SetPropertyValue(
		L"LocalDebuggerCommandArguments",
		active_project_cfg_name.val,
		vsix::PST_USER_FILE,
		args->in.bstrVal));

  wcscpy_s(last_cmdline_args, args->in.bstrVal);

  return S_OK;
}

// TODO: Figure out how to do the initial "wake up".
HRESULT on_update_sln_state() {
  // THREAD: VS Main
  
  // Set the combo box to the Startup Project's commandline args.
  // If the project is not a C++ project, set the combo box to "N/A".
	
  IUnknown_scope_t com_objs;

  auto sln = (vsix::IVsSolution *)com_objs.next();
  auto sln_build_manager = (vsix::IVsSolutionBuildManager *)com_objs.next();
  auto startup_project = (vsix::IVsHierarchy *)com_objs.next();
  auto props = (vsix::IVsBuildPropertyStorage *)com_objs.next();
  BSTR_RTTI_t active_project_cfg_name;

  HRCALL(vs_services.provider->QueryService(&sln));
  HRCALL(vs_services.provider->QueryService(&sln_build_manager));
  HRCALL(sln_build_manager->get_StartupProject(&startup_project));

  bool is_cpp = vsproj_is_cpp(startup_project);

  //BSTR_RTTI_t cmdline_args;
  wchar_t *cmdline_args = 0; // Do not free this here (or wrap in the automatic com_objs block). Leave it to VS.
	if (is_cpp) {
    HRCALL(sln_build_manager->get_ActiveProjectCfg_name(startup_project, &active_project_cfg_name.val));
    HRCALL(startup_project->get_BuildPropertyStorage(&props));
    HRCALL(props->GetPropertyValue(
      L"LocalDebuggerCommandArguments",
      active_project_cfg_name.val,
      vsix::PST_USER_FILE,
      &cmdline_args));
	}
	else {
    cmdline_args = SysAllocString(L"N/A");
	}

  if (wcscmp(last_cmdline_args, cmdline_args) != 0) {
		vs_services.vs_ui_shell->SetMRUComboTextW(&CLSID_cmdset, CMDID_cmdline_args_control, cmdline_args, true);
    wcscpy_s(last_cmdline_args, cmdline_args);
  }

  return S_OK;
}

// ================================ BOILERPLATE CONNECTION POINTS (START HERE) ================================
#pragma region

#define TIMER_UPDATE_SLN_STATE 0x01000000

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
	bool post_ok = false;
	if (InlineIsEqualGUID(cmdset_id, CLSID_cmdset) && cmd_id == CMDID_cmdline_args_control) {
    post_ok = PostMessageW(vs_services.main_hwnd, WM_CMDLINE_ARGS, 0, (LPARAM)args);
	}
	// else if (IsInlineEqualGUID(cmdset_id, ...) && cmd_id == ...) {
	//   ...
	// }

  if (!post_ok) {
    delete args;
		return OLECMDERR_E_NOTSUPPORTED;
		//return HRESULT_FROM_WIN32(GetLastError());
  }

  return S_OK;
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
		HRCALL(pkg_command_map(*pCmdGroupGuid, nCmdID, nCmdexecopt, in, out));
    return S_OK;
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