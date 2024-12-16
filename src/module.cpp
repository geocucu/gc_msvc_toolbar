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
#include <functional>

typedef unsigned long long u64;
typedef long long i64;

#ifdef _DEBUG
#define DEBUGBREAK __debugbreak();
#else 
#define DEBUGBREAK 
#endif 

// Also wrap in the HRESULT var. Use the watch window for the code. 
#define HRCALL_NORET(x, msg) { \
  HRESULT hr = x; \
  if (FAILED(hr)) { msgbox_hresult(hr, msg); DEBUGBREAK } }

#define HRCALL(x, msg) { \
  HRESULT hr = x; \
  if (FAILED(hr)) { msgbox_hresult(hr, msg); DEBUGBREAK; return hr; } }

void msgbox_hresult(HRESULT hr, const wchar_t *format, ...) {
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

struct {
	IServiceProvider *provider;
	CComPtr<IVsTaskSchedulerService> task_scheduler;
} vs_services;

struct task_body_t : CComObjectRootEx<CComSingleThreadModel>, IVsTaskBody {
	std::function<HRESULT(VARIANT *pResult)> task_logic;

	BEGIN_COM_MAP(task_body_t)
		COM_INTERFACE_ENTRY(IVsTaskBody)
	END_COM_MAP()

	task_body_t() : task_logic([](VARIANT *) -> HRESULT { return S_OK; }) {}
	task_body_t(const std::function<HRESULT(VARIANT *pResult)> &logic) : task_logic(logic) {}

	HRESULT __stdcall DoWork(IVsTask *pTask, DWORD dwCount, IVsTask **pParentTasks, VARIANT *pResult) override {
		if (!pResult) {
			return E_POINTER; 
		}

		VariantInit(pResult); 
		if (task_logic) {
			return task_logic(pResult);
		}
		return S_OK; 
	}

	static CComPtr<IVsTaskBody> create(const std::function<HRESULT(VARIANT *pResult)> &logic) {
		CComPtr<IVsTaskBody> task_body;
		CComObject<task_body_t> *raw_body = nullptr;
		HRESULT hr = CComObject<task_body_t>::CreateInstance(&raw_body);
		if (SUCCEEDED(hr) && raw_body) {
			raw_body->AddRef();
			raw_body->task_logic = logic;
			task_body.Attach(raw_body);
		}
		return task_body;
	}
};

// ================================ COMMANDLINE ARGS COMBO BOX ================================

HRESULT on_cmdline_args_body(VARIANT *result) {
	CComPtr<vsix::IVsSolutionBuildManager> build_manager;
	HRCALL(
		vs_services.provider->QueryService(vsix::IID_IVsSolutionBuildManager, vsix::IID_IVsSolutionBuildManager, (void **)&build_manager),
		L"Failed to retrieve IVsSolutionBuildManager");

	CComPtr<vsix::IVsHierarchy> startup_proj;
	HRCALL(
		build_manager->get_StartupProject(&startup_proj),
		L"Failed to get startup project");

	CComPtr<vsix::IVsBuildPropertyStorage> prop_storage;
	HRCALL(
		startup_proj->QueryInterface(vsix::IID_IVsBuildPropertyStorage, (void **)&prop_storage),
		L"Failed to query IVsBuildPropertyStorage from IVsHierarchy");

	HRCALL(
		prop_storage->SetPropertyValue(
		L"LocalDebuggerCommandArguments",
		L"Debug|x64",
		PST_USER_FILE,
		L"DUMMY ARG"), 
		L"Failed IVsBuildPropertyStorage::SetPropertyValue()");

	return S_OK;
}

HRESULT on_cmdline_args(VARIANT *in) {
	// THREAD: VS Main OR <some other> 
	//         ^^^ can only go wild here 
	
	CComBSTR new_arg(in->bstrVal);
	CComPtr<IVsTaskBody> task_body = task_body_t::create(on_cmdline_args_body);
	if (!task_body) {
		MessageBoxW(0, L"Failed to create task body", L"Error", MB_ICONERROR);
		return E_FAIL;
	}

	// NONE of these are guaranteed
  //VSTC_BACKGROUNDTHREAD = 0,
  //VSTC_UITHREAD_SEND = 1,
  //VSTC_UITHREAD_BACKGROUND_PRIORITY = 2,
  //VSTC_UITHREAD_IDLE_PRIORITY = 3,
  //VSTC_CURRENTCONTEXT = 4,
  //VSTC_BACKGROUNDTHREAD_LOW_IO_PRIORITY = 5,
  //VSTC_UITHREAD_NORMAL_PRIORITY = 6

	// DoWork() called in threads (by debugger "names"): "combase.dll" chain of calls or "Worker Thread"
	//CComPtr<IVsTask> tasks[7];
	//for (int i = 0; i < 7; ++i) {
	//	vs_services.task_scheduler->CreateTask((VSTASKRUNCONTEXT)i, task_body, &tasks[i]);
	//}

	//for (int i = 0; i < 7; ++i) {
	//	tasks[i]->Start();
	//}

	// TODO: Test IVsUIShell::PostExecCommand

	CComPtr<IVsTask> task;
	vs_services.task_scheduler->CreateTask(VSTC_UITHREAD_SEND, task_body, &task);

	return S_OK;
}

// ================================ BOILERPLATE CONNECTION POINTS (START HERE) ================================
#pragma region

HRESULT pkg_on_startup() {
	// THREAD: VS Main
	
	HRCALL_NORET(
		vs_services.provider->QueryService(IID_SVsTaskSchedulerService, IID_IVsTaskSchedulerService, (void **)&vs_services.task_scheduler),
		L"Failed to retrieve an IVsTaskSchedulerService");

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

	// Init VSPackage
	// https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.shell.interop.ivspackage.setsite?view=visualstudiosdk-2022
	HRESULT __stdcall SetSite(IServiceProvider *service_provider) override {
		if (!vs_services.provider) {
			// Only on init. 
			vs_services.provider = service_provider;
			HRCALL(pkg_on_startup(), L"Failed initialising extension on startup.");
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

#pragma warning(pop)

#pragma endregion