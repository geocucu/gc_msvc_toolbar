#include "common.h"

// TODO: FIX LIFETIMES 

#include "vsix.h"
#include "..\res\guids.h"
#include "..\res\resource.h"
#include "debug_helper.h"

// ======== COM Utils ======== 
#pragma region

IStream* com_duplicate_stream(IStream *src) {
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

	// == DTE ==
	vsix::_DTE *dte;
	vsix::_Solution *sln;
	vsix::SolutionBuild *sln_build;
} vs_services;

struct {
	IStream *project_streams[64];
	int num_projects;
} sln_state;

// ================================ COMMANDLINE ARGS COMBO BOX ================================

HRESULT on_cmdline_args(VARIANT *in) {
	// TODO: Use DTE for picking out the Startup Project, since get_StartupProject here will 
	// break, even with marshaling.
	
	// Do not release!
	vsix::IVsHierarchy *proj = 0;
	HRCALL(CoGetInterfaceAndReleaseStream(com_duplicate_stream(sln_state.project_streams[0]), vsix::IID_IVsHierarchy, (void **)&proj));

	vsix::IVsBuildPropertyStorage *props = 0;
	HRCALL(proj->QueryInterface(vsix::IID_IVsBuildPropertyStorage, (void **)&props));

	// == TODO: HARDCODED == 
	HRCALL(props->SetPropertyValue(
			L"LocalDebuggerCommandArguments",
			L"Release|x64",
			vsix::PST_USER_FILE,
			in->bstrVal));

	return S_OK;
}

// ================================ BOILERPLATE CONNECTION POINTS (START HERE) ================================
#pragma region

// Init VSPackage
// https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.shell.interop.ivspackage.setsite?view=visualstudiosdk-2022
HRESULT pkg_on_startup() {
	// THREAD: VS Main

	// == IVs* interfaces ==
	vsix::IVsSolution *sln = 0;
	HRCALL(vs_services.provider->QueryService(&sln));

  // Query startup project by new IVsSolution
  vsix::IVsSolutionBuildManager *sln_build_mgr = 0;
  HRCALL(vs_services.provider->QueryService(vsix::IID_IVsSolutionBuildManager, vsix::IID_IVsSolutionBuildManager, (void**)&sln_build_mgr));

	{
		vsix::IVsHierarchy *startup_project = 0;
		HRCALL(sln_build_mgr->get_StartupProject(&startup_project));

    vsix::IVsProjectCfg *active_project_cfg = 0;
    sln_build_mgr->FindActiveProjectCfg(NULL, NULL, startup_project, &active_project_cfg);

		BSTR active_project_cfg_name;
    active_project_cfg->get_DisplayName(&active_project_cfg_name); // Debug|x64

    vsix::IVsProjectCfgProvider *project_cfg_provider = 0;
    vsix::IVsBuildableProjectCfg *buildable_project_cfg = 0;

    active_project_cfg->get_ProjectCfgProvider(&project_cfg_provider);
    active_project_cfg->get_BuildableProjectCfg(&buildable_project_cfg);

    VARIANT name, project_type;
    VariantInit(&name);
    VariantInit(&project_type); // C++

    startup_project->GetProperty(VSITEMID_ROOT, vsix::VSHPROPID::VSHPROPID_Name, &name);
    startup_project->GetProperty(VSITEMID_ROOT, vsix::VSHPROPID::VSHPROPID_ProjectType, &project_type);


		VariantClear(&name);
		VariantClear(&project_type); // C++
	}

	vsix::IEnumHierarchies *hierarchies = 0;
	HRCALL(sln->GetProjectEnum(vsix::VSENUMPROJFLAGS::EPF_ALLPROJECTS, GUID_NULL, &hierarchies));

	vsix::IVsHierarchy *hierarchy = 0;
	ULONG fetched = 0;
	while (hierarchies->Next(1, &hierarchy, &fetched) == S_OK && fetched > 0) {
		if (!hierarchy) {
			continue;
		}

		HRCALL(CoMarshalInterThreadInterfaceInStream(vsix::IID_IVsHierarchy, (IUnknown *)hierarchy, &sln_state.project_streams[sln_state.num_projects]));
		sln_state.num_projects++;

		hierarchy->Release();
	}

	// == Older _DTE ==
	HRCALL(vs_services.provider->QueryService(&vs_services.dte));
	HRCALL(vs_services.dte->get_Solution(&vs_services.sln));
	HRCALL(vs_services.sln->get_SolutionBuild(&vs_services.sln_build));

	VARIANT projects;
	VariantInit(&projects);
	vs_services.sln_build->get_StartupProjects(&projects);
	VariantClear(&projects);

	vsix::SolutionConfigurations *sln_configs = 0;
	vs_services.sln_build->get_SolutionConfigurations(&sln_configs);
	long config_count = 0;
	sln_configs->get_Count(&config_count);
	for (int i = 0; i < config_count; ++i) {
		vsix::SolutionConfiguration *cfg = 0;
		VARIANT idx;
		VariantInit(&idx);
		idx.vt = VT_I4;
		idx.lVal = i + 1;
		sln_configs->Item(idx, &cfg);
		BSTR name;
		cfg->get_Name(&name);
		VariantClear(&idx);

		vsix::SolutionContexts *sln_contexts;
		cfg->get_SolutionContexts(&sln_contexts);
		long sln_contexts_count;
		sln_contexts->get_Count(&sln_contexts_count);
    for (int j = 0; j < sln_contexts_count; ++j) {
      vsix::SolutionContext *ctx = 0;
      VARIANT idx;
      VariantInit(&idx);
      idx.vt = VT_I4;
      idx.lVal = j + 1;
      sln_contexts->Item(idx, &ctx);
			BSTR platform_name, cfg_name, proj_name;
      ctx->get_PlatformName(&platform_name); // x64
      ctx->get_ConfigurationName(&cfg_name); // Debug
      ctx->get_ProjectName(&proj_name);      // List of projects with x64|Debug
    }
	}

	return S_OK;
}

HRESULT pkg_command_map(const GUID &cmdset_id, DWORD cmd_id, DWORD flags, VARIANT *in, VARIANT *out) {
	// THREAD: VS Main OR <some other> 
	//         ^^^ can only go wild here 

	if (InlineIsEqualGUID(cmdset_id, CLSID_cmdset) && cmd_id == CMDID_cmdline_args_control) {
		on_cmdline_args(in);
	}

	return OLECMDERR_E_NOTSUPPORTED;
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