#pragma warning(push)
#pragma warning(disable:28251) // Inconsistent annotations 

// Bake GUIDs
#define INITGUID
#include <guiddef.h>
#include "..\res\guids.h"
#undef INITGUID
#include <VSShellInterfaces.h>

// ======== MIDL ======== 
#define _MIDL_USE_GUIDDEF_ 
#include "gc_msvc_toolbar.c" 
#include "gc_msvc_toolbar.h" 

// ======== ATL ======== 
#define _ATL_APARTMENT_THREADED
#define _ATL_REGISTER_PER_USER
#include <atlbase.h>
#include <atlcom.h>
#include <atlwin.h>
#include <atlstr.h>
#include <atlfile.h>
#include <atlsafe.h>

// ======== Visual Studio ======== 
#include <dte.h> // for extensibility
#include <objext.h> // for ILocalRegistry
#include <vshelp.h> // for Help
#include <uilocale.h> // for IUIHostLocale2
#include <IVsQueryEditQuerySave2.h> // for IVsQueryEditQuerySave2
#include <vbapkg.h> // for IVsMacroRecorder
#include <fpstfmt.h> // for IPersistFileFormat
#include <VSRegKeyNames.h>
#include <stdidcmd.h>
#include <stdiduie.h> // For status bar consts.
#include <textfind.h>
#include <textmgr.h>

// ======== VSL ======== 
#define VSLASSERT _ASSERTE
#define VSLASSERTEX(exp, szMsg) _ASSERT_BASE(exp, szMsg)
#define VSLTRACE ATLTRACE
#include <VSLPackage.h>
#include <VSLCommandTarget.h>
#include <VSLWindows.h>
#include <VSLFile.h>
#include <VSLContainers.h>
#include <VSLComparison.h>
#include <VSLAutomation.h>
#include <VSLFindAndReplace.h>
#include <VSLShortNameDefines.h>

// ======== GUIDs/IDs ======== 
#include "..\res\resource.h"

#include "vsix_interfaces.h"
#include "ivs_interfaces.h"


typedef unsigned long long u64;
typedef long long i64;

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

	static void msgbox_hresult(HRESULT hr, const wchar_t *format, ...) {
		wchar_t error_msg[512] = {};
		wchar_t system_message[256] = {};
		wchar_t custom_message[256] = {};

		FormatMessageW(
			FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
			0,
			hr,
			MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
			system_message,
			sizeof(system_message) / sizeof(wchar_t),
			0
		);

		va_list args;
		va_start(args, format);
		vswprintf_s(custom_message, sizeof(custom_message) / sizeof(wchar_t), format, args);
		va_end(args);

		swprintf_s(error_msg, sizeof(error_msg) / sizeof(wchar_t), L"%s\nHRESULT: 0x%08X\n%s", custom_message, hr, system_message);

		MessageBoxW(0, error_msg, L"[Error] GC MSVC Toolbar", MB_ICONERROR | MB_OK);
	}

	// ================================ Interesting non-boilerplate START ================================

	static CommandHandler *GetCommand(const VSL::CommandId &target_id) {
		// Command map specific to this package
		// Combined command id: cmdset GUID & ID within cmdset
		static VSL::CommandId id_cmdline_args_control(CLSID_cmdset, CMDID_cmdline_args_control);
		static CommandHandler cmdline_args_handler(id_cmdline_args_control, 0, &on_cmdline_args);

		if (target_id == id_cmdline_args_control) {
			return &cmdline_args_handler;
		}

		return 0;
	}

	void on_cmdline_args(CommandHandler *sender, DWORD flags, VARIANT *in, VARIANT *out) {
		if (!marshaled_startup_proj_stream) {
			msgbox_hresult(E_POINTER, L"No valid marshaled startup project available");
			return;
		}

		HRESULT hr = S_OK;

		// Duplicate the stream for reuse
		IStream *duplicate_stream = 0;
		hr = CreateStreamOnHGlobal(0, TRUE, &duplicate_stream);
		if (FAILED(hr)) {
			msgbox_hresult(hr, L"Failed to create duplicate stream");
			return;
		}

		// Copy the original stream to the duplicate
		LARGE_INTEGER seek_start = {};
		ULARGE_INTEGER stream_size = {};
		marshaled_startup_proj_stream->Seek(seek_start, STREAM_SEEK_END, &stream_size);
		marshaled_startup_proj_stream->Seek(seek_start, STREAM_SEEK_SET, 0);

		hr = marshaled_startup_proj_stream->CopyTo(duplicate_stream, stream_size, 0, 0);
		if (FAILED(hr)) {
			msgbox_hresult(hr, L"Failed to copy to duplicate stream");
			duplicate_stream->Release();
			return;
		}

		marshaled_startup_proj_stream->Seek(seek_start, STREAM_SEEK_SET, 0);
		duplicate_stream->Seek(seek_start, STREAM_SEEK_SET, 0);

		// Unmarshal the interface from the duplicate stream
		CComPtr<vsix::IVsHierarchy> startup_proj;
		hr = CoGetInterfaceAndReleaseStream(duplicate_stream, vsix::IID_IVsHierarchy, (void **)&startup_proj);
		if (FAILED(hr)) {
			msgbox_hresult(hr, L"Failed to unmarshal IVsHierarchy from duplicate stream");
			return;
		}

		// Interact with the unmarshaled IVsHierarchy
		CComPtr<vsix::IVsBuildPropertyStorage> prop_storage;
		hr = startup_proj->QueryInterface(vsix::IID_IVsBuildPropertyStorage, (void **)&prop_storage);
		if (FAILED(hr)) {
			msgbox_hresult(hr, L"Failed to query IVsBuildPropertyStorage from IVsHierarchy");
			return;
		}

		// == TODO: HARDCODED == 

		prop_storage->SetPropertyValue(
			L"LocalDebuggerCommandArguments",
			L"Debug|x64",
			PST_USER_FILE,
			in->bstrVal
		);
	}

	// Init VSPackage
	// https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.shell.interop.ivspackage.setsite?view=visualstudiosdk-2022
	HRESULT __stdcall SetSite(IServiceProvider *service_provider) override {
		// Also called on VsSiteBaseImpl::Close (not documented)
		if (!service_provider) {
			return S_OK;
		}

		HRESULT hr = S_OK;

		hr = service_provider->QueryService(vsix::IID_IVsUIShell, vsix::IID_IVsUIShell, (void **)&services.ui_shell);
		if (FAILED(hr)) {
			msgbox_hresult(hr, L"Failed retrieving service: VS UI shell");
			return hr;
		}

		hr = service_provider->QueryService(vsix::IID_IVsSolution, vsix::IID_IVsSolution, (void **)&services.solution);
		if (FAILED(hr)) {
			msgbox_hresult(hr, L"Failed retrieving service: Solution");
			return hr;
		}

		hr = service_provider->QueryService(vsix::IID_IVsSolutionBuildManager, vsix::IID_IVsSolutionBuildManager, (void **)&services.build_manager);
		if (FAILED(hr)) {
			msgbox_hresult(hr, L"Failed retrieving service: Build Manager");
			return hr;
		}

		vsix::IVsHierarchy *startup_proj;
		hr = services.build_manager->get_StartupProject(&startup_proj);
		if (FAILED(hr)) {
			return hr;
		}

		// Create a stream to marshal the interface
		hr = CoMarshalInterThreadInterfaceInStream(vsix::IID_IVsHierarchy, (IUnknown*)startup_proj, &marshaled_startup_proj_stream);
		if (FAILED(hr)) {
			msgbox_hresult(hr, L"Failed to marshal IVsHierarchy to a stream");
			return hr;
		}

		return VSL::IVsPackageImpl<pkg_t, &CLSID_pkg>::SetSite(service_provider);
	}

	// ================= DATA =================

	// Service cache 
	struct {
		CComPtr<vsix::IVsUIShell> ui_shell;
		CComPtr<vsix::IVsSolution> solution;
		CComPtr<vsix::IVsSolutionBuildManager> build_manager;
	} services;

	// Don't wrap! Stream will be released by COM internally. 
	IStream *marshaled_startup_proj_stream = 0;

	// ================================ Interesting non-boilerplate END ================================
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

#pragma warning(pop)
