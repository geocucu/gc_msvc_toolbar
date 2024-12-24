#include "common.h"

// TODO: FIX LIFETIMES 

// ======== VSSDK ======== 
#include "vsix_ivs.h"
#include "vsix_dte.h"

// ======== GUIDs/IDs ======== 
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
	vsix::_DTE *dte;
} vs_services;

IStream *project_streams[64];
int num_projects;

// ================================ COMMANDLINE ARGS COMBO BOX ================================

HRESULT on_cmdline_args(VARIANT *in) {
	// TODO: Use DTE for picking out the Startup Project, since get_StartupProject here will 
	// break, even with marshaling.
	
	// Do not release!
	vsix::IVsHierarchy *proj;
	HRCALL(CoGetInterfaceAndReleaseStream(com_duplicate_stream(project_streams[0]), vsix::IID_IVsHierarchy, (void **)&proj));

	vsix::IVsBuildPropertyStorage *props;
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
	{
		vsix::IVsSolution *sln;
		HRCALL(vs_services.provider->QueryService(vsix::IID_IVsSolution, vsix::IID_IVsSolution, (void **)&sln));

		vsix::IEnumHierarchies *hierarchies;
		HRCALL(sln->GetProjectEnum(vsix::VSENUMPROJFLAGS::EPF_ALLPROJECTS, GUID_NULL, &hierarchies));

		vsix::IVsHierarchy *hierarchy;
		ULONG fetched = 0;
		while (hierarchies->Next(1, &hierarchy, &fetched) == S_OK && fetched > 0) {
			if (!hierarchy) {
				continue;
			}

			HRCALL(CoMarshalInterThreadInterfaceInStream(vsix::IID_IVsHierarchy, (IUnknown *)hierarchy, &project_streams[num_projects]));
			num_projects++;

			hierarchy->Release();
		}
	}

	// == Older _DTE ==
	{
		vsix::_DTE *dte;
		HRCALL(vs_services.provider->QueryService(vsix::IID__DTE, vsix::IID__DTE, (void**)&dte));

		IUnknown *sln_CLR;
		HRCALL(dte->get_Solution((vsix::Solution **)&sln_CLR));

		// Actual sln
		vsix::_Solution *sln;
		HRCALL(sln_CLR->QueryInterface(vsix::IID__Solution, (void **)&sln));

		vsix::SolutionBuild *sln_build;
		HRCALL(sln->get_SolutionBuild(&sln_build));
		
		VARIANT projects;
		VariantInit(&projects);
		sln_build->get_StartupProjects(&projects);
		VariantClear(&projects);
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