#pragma once

//
// Older Visual Studio SDK ("DTE" based)
//

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

// DTE 
#include <dte80.h>
#include <dte90.h>
#include <dte90a.h>
#include <dte100.h>

// ==== TODO: DELETE? ====
// It's probably fine to stick to the IVs[...] interfaces, and not need DTE at all. 

namespace vsix {
	struct _DTE;
	struct Solution;
	struct _Solution;
	struct SolutionBuild;
	struct Projects;
	struct Project;
	struct Properties;
	struct Property;
	struct ConfigurationManager;
	struct Configurations;
	struct Configuration;

	// ================================ _DTE ================================
#pragma region
	const GUID IID__DTE = { 0x04a72314, 0x32e9, 0x48e2,{ 0x9b, 0x87, 0xa6, 0x36, 0x03, 0x45, 0x4f, 0x3e } };
	struct __declspec(novtable) _DTE {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **obj) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(const IID &riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, const IID &riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == _DTE == 
		virtual HRESULT __stdcall get_Name(BSTR *lpbstrReturn) = 0;
		virtual HRESULT __stdcall get_FileName(BSTR *lpbstrReturn) = 0;
		virtual HRESULT __stdcall get_Version(BSTR *lpbstrReturn) = 0;
		virtual HRESULT __stdcall get_CommandBars(IDispatch **ppcbs) = 0;
		virtual HRESULT __stdcall get_Windows(Windows **ppwnsVBWindows) = 0;
		virtual HRESULT __stdcall get_Events(Events **ppevtEvents) = 0;
		virtual HRESULT __stdcall get_AddIns(AddIns **lpppAddIns) = 0;
		virtual HRESULT __stdcall get_MainWindow(Window **ppwin) = 0;
		virtual HRESULT __stdcall get_ActiveWindow(Window **ppwinActive) = 0;
		virtual HRESULT __stdcall Quit() = 0;
		virtual HRESULT __stdcall get_DisplayMode(vsDisplay *lpDispMode) = 0;
		virtual HRESULT __stdcall put_DisplayMode(vsDisplay DispMode) = 0;
		virtual HRESULT __stdcall get_Solution(Solution **ppSolution) = 0;
		virtual HRESULT __stdcall get_Commands(Commands **ppCommands) = 0;
		virtual HRESULT __stdcall GetObject(BSTR Name, IDispatch **ppObject) = 0;
		virtual HRESULT __stdcall get_Properties(BSTR Category, BSTR Page, Properties **ppObject) = 0;
		virtual HRESULT __stdcall get_SelectedItems(SelectedItems **ppSelectedItems) = 0;
		virtual HRESULT __stdcall get_CommandLineArguments(BSTR *lpbstrReturn) = 0;
		virtual HRESULT __stdcall OpenFile(BSTR ViewKind, BSTR FileName, Window **ppwin) = 0;
		virtual HRESULT __stdcall get_IsOpenFile(BSTR ViewKind, BSTR FileName, VARIANT_BOOL *lpfReturn) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_LocaleID(long *lpReturn) = 0;
		virtual HRESULT __stdcall get_WindowConfigurations(WindowConfigurations **WindowConfigurationsObject) = 0;
		virtual HRESULT __stdcall get_Documents(Documents **ppDocuments) = 0;
		virtual HRESULT __stdcall get_ActiveDocument(Document **ppDocument) = 0;
		virtual HRESULT __stdcall ExecuteCommand(BSTR CommandName, BSTR CommandArgs) = 0;
		virtual HRESULT __stdcall get_Globals(Globals **ppGlobals) = 0;
		virtual HRESULT __stdcall get_StatusBar(StatusBar **ppStatusBar) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *lpbstrReturn) = 0;
		virtual HRESULT __stdcall get_UserControl(VARIANT_BOOL *UserControl) = 0;
		virtual HRESULT __stdcall put_UserControl(VARIANT_BOOL HasControl) = 0;
		virtual HRESULT __stdcall get_ObjectExtenders(ObjectExtenders **ppObjectExtenders) = 0;
		virtual HRESULT __stdcall get_Find(Find **ppFind) = 0;
		virtual HRESULT __stdcall get_Mode(vsIDEMode *pVal) = 0;
		virtual HRESULT __stdcall LaunchWizard(BSTR VSZFile, SAFEARRAY **ContextParams, wizardResult *pResult) = 0;
		virtual HRESULT __stdcall get_ItemOperations(ItemOperations **ppItemOperations) = 0;
		virtual HRESULT __stdcall get_UndoContext(UndoContext **ppUndoContext) = 0;
		virtual HRESULT __stdcall get_Macros(Macros **ppMacros) = 0;
		virtual HRESULT __stdcall get_ActiveSolutionProjects(VARIANT *pProjects) = 0;
		virtual HRESULT __stdcall get_MacrosIDE(DTE **pDTE) = 0;
		virtual HRESULT __stdcall get_RegistryRoot(BSTR *pVal) = 0;
		virtual HRESULT __stdcall get_Application(DTE **pVal) = 0;
		virtual HRESULT __stdcall get_ContextAttributes(ContextAttributes **ppVal) = 0;
		virtual HRESULT __stdcall get_SourceControl(SourceControl **ppVal) = 0;
		virtual HRESULT __stdcall get_SuppressUI(VARIANT_BOOL *pVal) = 0;
		virtual HRESULT __stdcall put_SuppressUI(VARIANT_BOOL Suppress) = 0;
		virtual HRESULT __stdcall get_Debugger(Debugger **ppDebugger) = 0;
		virtual HRESULT __stdcall SatelliteDLLPath(BSTR Path, BSTR Name, BSTR *pFullPath) = 0;
		virtual HRESULT __stdcall get_Edition(BSTR *ProductEdition) = 0;
	};
#pragma endregion

	// ================================ Solution ================================
#pragma region
	const GUID IID_Solution = { 0xb35caa8c, 0x77de, 0x4ab3, { 0x8e, 0x5a, 0xf0, 0x38, 0xe3, 0xfc, 0x60, 0x56 } };
	// Only forward declared. 
	// Retrieve _Solution from this with QueryInterface.
	struct Solution;
#pragma endregion

	// ================================ _Solution ================================
#pragma region
	const GUID IID__Solution = { 0x26f6cc4b, 0x7a48, 0x4e4d,{ 0x8a, 0xf5, 0x9e, 0x96, 0x02, 0x32, 0xe0, 0x5f } };
	
	struct __declspec(novtable) _Solution {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(const IID &riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, const IID &riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == _Solution methods == 
		virtual HRESULT __stdcall Item(VARIANT index, Project **lppcReturn) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **lppiuReturn) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Parent(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Count(long *lplReturn) = 0;
		virtual HRESULT __stdcall get_FileName(BSTR *lpbstrReturn) = 0;
		virtual HRESULT __stdcall SaveAs(BSTR FileName) = 0;
		virtual HRESULT __stdcall AddFromTemplate(BSTR FileName, BSTR Destination, BSTR ProjectName, VARIANT_BOOL Exclusive, Project **lppptReturn) = 0;
		virtual HRESULT __stdcall AddFromFile(BSTR FileName, VARIANT_BOOL Exclusive, Project **lppptReturn) = 0;
		virtual HRESULT __stdcall Open(BSTR FileName) = 0;
		virtual HRESULT __stdcall Close(VARIANT_BOOL SaveFirst) = 0;
		virtual HRESULT __stdcall get_Properties(Properties **ppObject) = 0;
		virtual HRESULT __stdcall get_IsDirty(VARIANT_BOOL *lpfReturn) = 0;
		virtual HRESULT __stdcall put_IsDirty(VARIANT_BOOL fDirty) = 0;
		virtual HRESULT __stdcall Remove(Project *proj) = 0;
		virtual HRESULT __stdcall get_TemplatePath(BSTR ProjectType, BSTR *lpbstrPath) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *lpbstrReturn) = 0;
		virtual HRESULT __stdcall get_Saved(VARIANT_BOOL *lpfReturn) = 0;
		virtual HRESULT __stdcall put_Saved(VARIANT_BOOL fDirty) = 0;
		virtual HRESULT __stdcall get_Globals(Globals **ppGlobals) = 0;
		virtual HRESULT __stdcall get_AddIns(AddIns **ppAddIns) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall get_IsOpen(VARIANT_BOOL *pVal) = 0;
		virtual HRESULT __stdcall get_SolutionBuild(SolutionBuild **ppSolutionBuild) = 0;
		virtual HRESULT __stdcall Create(BSTR Destination, BSTR Name) = 0;
		virtual HRESULT __stdcall get_Projects(Projects **ppProjects) = 0;
		virtual HRESULT __stdcall FindProjectItem(BSTR FileName, ProjectItem **ppProjectItem) = 0;
		virtual HRESULT __stdcall ProjectItemsTemplatePath(BSTR ProjectKind, BSTR *pFullPath) = 0;
	};
#pragma endregion

	// ================================ SolutionBuild ================================
#pragma region
	const GUID IID_SolutionBuild = { 0xa3c1c40c, 0x9218, 0x4d4c, { 0x9d, 0xaa, 0x07, 0x5f, 0x64, 0xf6, 0x92, 0x2c } };
	struct __declspec(novtable) SolutionBuild {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(const IID &riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, const IID &riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == SolutionBuild == 
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Parent(Solution **ppSolution) = 0;
		virtual HRESULT __stdcall get_ActiveConfiguration(SolutionConfiguration **ppSolutionConfiguration) = 0;
		virtual HRESULT __stdcall get_BuildDependencies(BuildDependencies **ppBuildDependencies) = 0;
		virtual HRESULT __stdcall get_BuildState(vsBuildState *pvsBuildState) = 0;
		virtual HRESULT __stdcall get_LastBuildInfo(long *pBuiltSuccessfully) = 0;
		virtual HRESULT __stdcall put_StartupProjects(VARIANT Projects) = 0;
		virtual HRESULT __stdcall get_StartupProjects(VARIANT *pProject) = 0;
		virtual HRESULT __stdcall Build(VARIANT_BOOL WaitForBuildToFinish = 0) = 0;
		virtual HRESULT __stdcall Debug() = 0;
		virtual HRESULT __stdcall Deploy(VARIANT_BOOL WaitForDeployToFinish = 0) = 0;
		virtual HRESULT __stdcall Clean(VARIANT_BOOL WaitForCleanToFinish = 0) = 0;
		virtual HRESULT __stdcall Run() = 0;
		virtual HRESULT __stdcall get_SolutionConfigurations(SolutionConfigurations **ppSolutionConfigurations) = 0;
		virtual HRESULT __stdcall BuildProject(BSTR SolutionConfiguration, BSTR ProjectUniqueName, VARIANT_BOOL WaitForBuildToFinish = 0) = 0;
	};
#pragma endregion

	// ================================ _Solution ================================
#pragma region
	const GUID IID_Projects = { 0xe3ec0add, 0x31b3, 0x461f, { 0x83, 0x03, 0x8a, 0x5e, 0x69, 0x31, 0x25, 0x7a } };
	struct __declspec(novtable) Projects {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(const IID &riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, const IID &riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Projects == 
		virtual HRESULT __stdcall Item(VARIANT index, Project **lppcReturn) = 0;
		virtual HRESULT __stdcall get_Parent(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Count(long *lplReturn) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **lppiuReturn) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Properties(Properties **ppObject) = 0;
		virtual HRESULT __stdcall get_Kind(BSTR *lpbstrReturn) = 0;
	};
#pragma endregion

	// ================================ Project ================================
#pragma region
	const GUID IID_Project = { 0x866311e6, 0xc887, 0x4143, { 0x98, 0x33, 0x64, 0x5f, 0x5b, 0x93, 0xf6, 0xf1 } };
	struct __declspec(novtable) Project {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(const IID &riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, const IID &riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// Project
		virtual HRESULT __stdcall get_Name(BSTR *lpbstrName) = 0;
		virtual HRESULT __stdcall put_Name(BSTR bstrName) = 0;
		virtual HRESULT __stdcall get_FileName(BSTR *lpbstrReturn) = 0;
		virtual HRESULT __stdcall get_IsDirty(VARIANT_BOOL *lpfReturn) = 0;
		virtual HRESULT __stdcall put_IsDirty(VARIANT_BOOL Dirty) = 0;
		virtual HRESULT __stdcall get_Collection(Projects **lppaReturn) = 0;
		virtual HRESULT __stdcall SaveAs(BSTR NewFileName) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Kind(BSTR *lpbstrFileName) = 0;
		virtual HRESULT __stdcall get_ProjectItems(ProjectItems **lppcReturn) = 0;
		virtual HRESULT __stdcall get_Properties(Properties **ppObject) = 0;
		virtual HRESULT __stdcall get_UniqueName(BSTR *lpbstrFileName) = 0;
		virtual HRESULT __stdcall get_Object(IDispatch **ProjectModel) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *lpbstrReturn) = 0;
		virtual HRESULT __stdcall get_Saved(VARIANT_BOOL *lpfReturn) = 0;
		virtual HRESULT __stdcall put_Saved(VARIANT_BOOL SavedFlag) = 0;
		virtual HRESULT __stdcall get_ConfigurationManager(ConfigurationManager **ppConfigurationManager) = 0;
		virtual HRESULT __stdcall get_Globals(Globals **ppGlobals) = 0;
		virtual HRESULT __stdcall Save(BSTR FileName = L"") = 0;
		virtual HRESULT __stdcall get_ParentProjectItem(ProjectItem **ppParentProjectItem) = 0;
		virtual HRESULT __stdcall get_CodeModel(CodeModel **ppCodeModel) = 0;
		virtual HRESULT __stdcall Delete() = 0;
	};
#pragma endregion

	// ================================ Properties ================================
#pragma region
	const GUID IID_Properties = { 0x4cc8ccf5, 0xa926, 0x4646, { 0xb1, 0x7f, 0xb4, 0x94, 0x0c, 0xae, 0xd4, 0x72 } };
	struct __declspec(novtable) Properties {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(const IID &riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, const IID &riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Properties == 
		virtual HRESULT __stdcall Item(VARIANT index, Property **lplppReturn) = 0;
		virtual HRESULT __stdcall get_Application(IDispatch **lppidReturn) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **lppidReturn) = 0;
		virtual HRESULT __stdcall get_Count(long *lplReturn) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **lppiuReturn) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
	};
#pragma endregion

	// ================================ Property ================================
#pragma region
	const GUID IID_Property = { 0x7b988e06, 0x2581, 0x485e, { 0x93, 0x22, 0x04, 0x88, 0x1e, 0x06, 0x00, 0xd0 } };
	struct __declspec(novtable) Property {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(const IID &riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, const IID &riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Property == 
		virtual HRESULT __stdcall get_Value(VARIANT *lppvReturn) = 0;
		virtual HRESULT __stdcall put_Value(VARIANT NewValue) = 0;
		virtual HRESULT __stdcall putref_Value(VARIANT NewValue) = 0;
		virtual HRESULT __stdcall get_IndexedValue(VARIANT Index1, VARIANT Index2, VARIANT Index3, VARIANT Index4, VARIANT *Val) = 0;
		virtual HRESULT __stdcall put_IndexedValue(VARIANT Index1, VARIANT Index2, VARIANT Index3, VARIANT Index4, VARIANT NewValue) = 0;
		virtual HRESULT __stdcall get_NumIndices(short *lpiRetVal) = 0;
		virtual HRESULT __stdcall get_Application(IDispatch **lppidReturn) = 0;
		virtual HRESULT __stdcall get_Parent(Properties **lpppReturn) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *lpbstrReturn) = 0;
		virtual HRESULT __stdcall get_Collection(Properties **lpppReturn) = 0;
		virtual HRESULT __stdcall get_Object(IDispatch **lppunk) = 0;
		virtual HRESULT __stdcall putref_Object(IUnknown *lpunk) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
	};
#pragma endregion

	// ================================ ConfigurationManager ================================
#pragma region
	const GUID IID_ConfigurationManager = { 0x9043fda1, 0x345b, 0x4364, { 0x90, 0x0f, 0xbc, 0x85, 0x98, 0xeb, 0x8e, 0x4f } };
	struct __declspec(novtable) ConfigurationManager {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(const IID &riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, const IID &riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == ConfigurationManager == 
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **ppParent) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **Enumerator) = 0;
		virtual HRESULT __stdcall Item(VARIANT Index, BSTR Platform, Configuration **ppOut) = 0;
		virtual HRESULT __stdcall get_Count(long *pCount) = 0;
		virtual HRESULT __stdcall ConfigurationRow(BSTR Name, Configurations **ppOut) = 0;
		virtual HRESULT __stdcall AddConfigurationRow(BSTR NewName, BSTR ExistingName, VARIANT_BOOL Propagate, Configurations **ppOut) = 0;
		virtual HRESULT __stdcall DeleteConfigurationRow(BSTR Name) = 0;
		virtual HRESULT __stdcall get_ConfigurationRowNames(VARIANT *pNames) = 0;
		virtual HRESULT __stdcall Platform(BSTR Name, Configurations **ppOut) = 0;
		virtual HRESULT __stdcall AddPlatform(BSTR NewName, BSTR ExistingName, VARIANT_BOOL Propagate, Configurations **ppOut) = 0;
		virtual HRESULT __stdcall DeletePlatform(BSTR Name) = 0;
		virtual HRESULT __stdcall get_PlatformNames(VARIANT *pNames) = 0;
		virtual HRESULT __stdcall get_SupportedPlatforms(VARIANT *pPlatforms) = 0;
		virtual HRESULT __stdcall get_ActiveConfiguration(Configuration **ppOut) = 0;
	};
#pragma endregion

	// ================================ Configurations ================================
	const GUID IID_Configurations = { 0xb6b4c8d6, 0x4d27, 0x43b9, { 0xb4, 0x5c, 0x52, 0xbd, 0x16, 0xb6, 0xba, 0x38 } };
	struct __declspec(novtable) Configurations {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(const IID &riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, const IID &riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Configurations ==
		virtual HRESULT __stdcall get_DTE(DTE * *ppDTE) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **ppParent) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **Enumerator) = 0;
		virtual HRESULT __stdcall Item(VARIANT Index, Configuration **ppOut) = 0;
		virtual HRESULT __stdcall get_Count(long *pCount) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pName) = 0;
		virtual HRESULT __stdcall get_Type(vsConfigurationType *pType) = 0;
	};

	// ================================ Configuration ================================
#pragma region
	const GUID IID_Configuration = { 0x90813589, 0xfe21, 0x4aa4, { 0xa2, 0xe5, 0x05, 0x3f, 0xd2, 0x74, 0xe9, 0x80 } };
	struct __declspec(novtable) Configuration {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch == 
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(const IID &riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, const IID &riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Configuration ==

		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Collection(ConfigurationManager **ppConfigurationManager) = 0;
		virtual HRESULT __stdcall get_ConfigurationName(BSTR *pName) = 0;
		virtual HRESULT __stdcall get_PlatformName(BSTR *pName) = 0;
		virtual HRESULT __stdcall get_Type(vsConfigurationType *pType) = 0;
		virtual HRESULT __stdcall get_Owner(IDispatch **ppOwner) = 0;

		// Will return 0 for MSVC Projects!
		virtual HRESULT __stdcall get_Properties(Properties **ppProperties) = 0;

		virtual HRESULT __stdcall get_IsBuildable(VARIANT_BOOL *pBuildable) = 0;
		virtual HRESULT __stdcall get_IsRunable(VARIANT_BOOL *pRunable) = 0;
		virtual HRESULT __stdcall get_IsDeployable(VARIANT_BOOL *pDeployable) = 0;
		virtual HRESULT __stdcall get_Object(IDispatch **ppDisp) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall get_OutputGroups(OutputGroups **ppOutputGroups) = 0;
	};
#pragma endregion
}