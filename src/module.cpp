#include "common.h"

// ======== ATL ======== 
#define _ATL_APARTMENT_THREADED
#define _ATL_REGISTER_PER_USER
#include <atlbase.h>
#include <atlcom.h>
#include <atlwin.h>
#include <atlstr.h>
#include <atlfile.h>
#include <atlsafe.h>

// ======== VSSDK ======== 
#include "vsix_ivs.h"
#include "vsix_dte.h"

// ======== GUIDs/IDs ======== 
#include "..\res\guids.h"
#include "..\res\resource.h"

#include "debug_helper.h"

// ======== COM Utils ======== 
#pragma region

CComPtr<IStream> com_duplicate_stream(IStream *src) {
	CComPtr<IStream> new_stream;
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
	IServiceProvider *provider;
	CComPtr<vsix::_DTE> dte;
} vs_services;

IStream *project_streams[64];
int num_projects;

// ================================ COMMANDLINE ARGS COMBO BOX ================================

HRESULT on_cmdline_args(VARIANT *in) {
	// TODO: Use DTE for picking out the Startup Project, since get_StartupProject here will 
	// break, even with marshaling.
	
	// Do not release!
	IVsHierarchy *proj;
	HRCALL(CoGetInterfaceAndReleaseStream(com_duplicate_stream(project_streams[0]), vsix::IID_IVsHierarchy, (void **)&proj));

	CComPtr<vsix::IVsBuildPropertyStorage> props;
	HRCALL(proj->QueryInterface(vsix::IID_IVsBuildPropertyStorage, (void **)&props));

	// == TODO: HARDCODED == 
	HRCALL(props->SetPropertyValue(
			L"LocalDebuggerCommandArguments",
			L"Release|x64",
			PST_USER_FILE,
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
		CComPtr<vsix::IVsSolution> sln;
		HRCALL(vs_services.provider->QueryService(vsix::IID_IVsSolution, vsix::IID_IVsSolution, (void **)&sln));

		CComPtr<IEnumHierarchies> hierarchies;
		HRCALL(sln->GetProjectEnum(__VSENUMPROJFLAGS::EPF_ALLPROJECTS, GUID_NULL, &hierarchies));

		CComPtr<IVsHierarchy> hierarchy;
		ULONG fetched = 0;
		while (hierarchies->Next(1, &hierarchy, &fetched) == S_OK && fetched > 0) {
			if (!hierarchy) {
				continue;
			}

			HRCALL(CoMarshalInterThreadInterfaceInStream(vsix::IID_IVsHierarchy, (IUnknown *)hierarchy, &project_streams[num_projects]));
			num_projects++;

			hierarchy.Release();
		}
	}

	// == Older _DTE ==
	{
		CComPtr<VxDTE::_DTE> dte;
		HRCALL(vs_services.provider->QueryService(__uuidof(VxDTE::_DTE), __uuidof(VxDTE::_DTE), (void**)&dte));

		CComPtr<IUnknown> sln_CLR;
		HRCALL(dte->get_Solution((VxDTE::Solution **)&sln_CLR));

		// Actual sln
		CComPtr<VxDTE::_Solution> sln;
		HRCALL(sln_CLR->QueryInterface(__uuidof(VxDTE::_Solution), (void **)&sln));

		CComPtr<VxDTE::SolutionBuild> sln_build;
		HRCALL(sln->get_SolutionBuild(&sln_build));

		CComVariant projects;
		sln_build->get_StartupProjects(&projects);
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

struct __declspec(novtable) pkg_t :
	// Non-thread safe COM object; partial implementation for IUnknown (the COM map below provides the rest).
	CComObjectRootEx<CComSingleThreadModel>,
	CComCoClass<pkg_t, &CLSID_pkg>,
	// Implement IVsPackage => make this COM object into a VS Package.
	VSL::IVsPackageImpl<pkg_t, &CLSID_pkg>,
	VSL::IOleCommandTargetImpl<pkg_t>,
	// Provides consumers of this object with the ability to determine which interfaces support extended error information.
	ATL::ISupportErrorInfoImpl<&IID_IVsPackage>
{
	// ==== Start of COM map ==== 
	// Provides a portion of the implementation of IUnknown, in particular the list of interfaces the pkg_t object will support via QueryInterface
	
	//BEGIN_COM_MAP(pkg_t)
	typedef pkg_t _ComMapClass; 
	static HRESULT __stdcall _Cache(void *pv, const IID &iid, void **ppvObject, u64 dw) {
		pkg_t *p = (pkg_t *)pv;
		p->Lock(); 
		HRESULT hr = E_FAIL; 
		__try {
			hr = ATL::CComObjectRootBase::_Cache(pv, iid, ppvObject, dw);
		}
		__finally {
			p->Unlock();
		} 
		return hr;
	} 
	
	IUnknown *_GetRawUnknown() {
		return (IUnknown *)((i64)this + _GetEntries()->dw);
	} 
	
	IUnknown *GetUnknown() {
		return (IUnknown *)((i64)this + _GetEntries()->dw);
	} 
	
	HRESULT _InternalQueryInterface(const IID &iid, void **ppvObject) {
		return this->InternalQueryInterface(this, _GetEntries(), iid, ppvObject);
	} 
	
	static const ATL::_ATL_INTMAP_ENTRY* _GetEntries() {
		static pkg_t *addr_8 = (pkg_t *)8; // Fake non-null addr for adjusting pointers by base interfaces (within the derived pkg_t)
		static u64 dw_IVsPackage        = u64((IVsPackage *)       addr_8) - 8; // Offset of vfptr for IVsPackage
		static u64 dw_IOleCommandTarget = u64((IOleCommandTarget *)addr_8) - 8; // Offset of vfptr for IOleCommandTarget
		static u64 dw_ISupportErrorInfo = u64((ISupportErrorInfo *)addr_8) - 8; // Offset of vfptr for ISupportErrorInfo
		static _ATL_CREATORARGFUNC *func = (ATL::_ATL_CREATORARGFUNC *)1; // Sentinel?

		static const ATL::_ATL_INTMAP_ENTRY entries[] = {
			//COM_INTERFACE_ENTRY(IVsPackage)
			{ &IID_IVsPackage, dw_IVsPackage, func },
		
			//COM_INTERFACE_ENTRY(IOleCommandTarget)
			{ &IID_IOleCommandTarget, dw_IOleCommandTarget, func },
		
			//COM_INTERFACE_ENTRY(ISupportErrorInfo)
			{ &IID_ISupportErrorInfo, dw_ISupportErrorInfo, func },

			//END_COM_MAP()
__if_exists(_GetAttrEntries) { 
			{ 0, (u64)_GetAttrEntries, _ChainAttr }, } 

			{ 0, 0, 0 } 
		}; 
		
		// Return &entries[1] instead, if the map contains { 0, (DWORD_PTR)L"pkg_t", 0 } as the first elem.
		return entries;
	} 
	
	virtual ULONG __stdcall AddRef() = 0; 
	virtual ULONG __stdcall Release() = 0; 
	virtual __declspec(nothrow) HRESULT __stdcall QueryInterface(const IID &, void **) = 0;

	// ==== End of COM map === 

	// Provide error information if it is not possible to load the UI dll. 
	static const LoadUILibrary::ExtendedErrorInfo &GetLoadUILibraryErrorInfo() {
		static LoadUILibrary::ExtendedErrorInfo info(L"The product is not installed properly. Please reinstall.");
		return info;
	}

	// DLL is registered with VS via a pkgdef file. Don't do anything if asked to self-register.
	static HRESULT UpdateRegistry(BOOL bRegister) { return S_OK; }
	static CommandHandler *GetCommand(const VSL::CommandId &target_id) { return 0; }

	HRESULT __stdcall Exec(const GUID *pCmdGroupGuid, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *in, VARIANT *out) override {
		return pkg_command_map(*pCmdGroupGuid, nCmdID, nCmdexecopt, in, out);
	}

	HRESULT __stdcall SetSite(IServiceProvider *service_provider) override {
		if (!vs_services.provider) {
			// Only on init. 
			vs_services.provider = service_provider;
			HRCALL(pkg_on_startup());
		}

		return VSL::IVsPackageImpl<pkg_t, &CLSID_pkg>::SetSite(service_provider);
	}
};

// This exposes pkg_t for instantiation via DllGetClassObject; however, an instance
// can not be created by CoCreateInstance, as pkg_t is specifically registered with
// VS, not the the system in general.
OBJECT_ENTRY_AUTO(CLSID_pkg, pkg_t)

// Should be instantiated before the point where DllGetClassObject(factory...) goes through 
struct dll_module_t : CAtlDllModuleT<dll_module_t> {} dll_module;

extern "C" BOOL DllMain(HINSTANCE, DWORD reason, void *) {
	if (reason == DLL_THREAD_ATTACH && CAtlBaseModule::m_bInitFailed) {
		return 0;
	}

	return 1;
}

extern "C" HRESULT DllCanUnloadNow() {
	if (dll_module.m_nLockCnt == 0) {
		return S_OK;
	}

	return S_FALSE;
}

// Returns a class factory to create an object of the requested type
extern "C" HRESULT DllGetClassObject(const IID &rclsid, const IID &riid, void **ppv) {
	if (!ppv) {
		return E_POINTER;
	}

	*ppv = 0;

	HRESULT hr = S_OK;

	for (_ATL_OBJMAP_ENTRY_EX **iter = _AtlComModule.m_ppAutoObjMapFirst; iter < _AtlComModule.m_ppAutoObjMapLast; iter++) {
		if (*iter == 0) {
			continue;
		}

		const _ATL_OBJMAP_ENTRY_EX *entry = *iter;

		if (!entry->pfnGetClassObject || !InlineIsEqualGUID(rclsid, *entry->pclsid)) {
			continue;
		}

		_ATL_OBJMAP_CACHE *cache = entry->pCache;

		if (!cache->pCF) {
			CComCritSecLock<CComCriticalSection> lock(_AtlComModule.m_csObjMap, false);
			hr = lock.Lock();
			if (FAILED(hr)) {
				break;
			}

			if (!cache->pCF) {
				IUnknown *factory = 0;
				hr = entry->pfnGetClassObject(entry->pfnCreateInstance, IID_IUnknown, (void**)&factory);
				if (SUCCEEDED(hr)) {
					cache->pCF = (IUnknown*)EncodePointer(factory);
				}
			}
		}

		if (cache->pCF) {
			// Decode factory pointer
			IUnknown *factory = (IUnknown*)DecodePointer(cache->pCF);
			hr = factory->QueryInterface(riid, ppv);
		}

		break;
	}

	if (*ppv == NULL && hr == S_OK) {
		hr = CLASS_E_CLASSNOTAVAILABLE;
	}

	return hr;
}

#pragma endregion