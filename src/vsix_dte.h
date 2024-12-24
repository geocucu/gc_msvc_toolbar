#pragma once

//
// Older Visual Studio SDK ("DTE" based)
//

#include <guiddef.h>


namespace vsix {
	// Only forward declared
	struct DTE;
	struct CommandEvents;
	struct SelectionEvents;
	struct SolutionEvents;
	struct BuildEvents;
	struct WindowEvents;
	struct OutputWindowEvents;
	struct FindEvents;
	struct TaskListEvents;
	struct DTEEvents;
	struct DocumentEvents;
	struct ProjectItemsEvents;
	struct DebuggerEvents;
	struct TextEditorEvents;

	// Full definitions
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
	struct IDispatch;
	struct Windows;
	struct Window;
	struct LinkedWindows;
	struct ProjectItem;
	struct ProjectItems;
	struct FileCodeModel;
	struct CodeElement;
	struct CodeElements;
	struct TextPoint;
	struct TextDocument;
	struct Document;
	struct Documents;
	struct TextSelection;
	struct VirtualPoint;
	struct EditPoint;
	struct TextRange;
	struct TextRanges;
	struct TextPane;
	struct TextPanes;
	struct TextWindow;
	struct CodeNamespace;
	struct CodeClass;
	struct CodeAttribute;
	struct CodeInterface;
	struct CodeFunction;
	struct CodeTypeRef;
	struct CodeType;
	struct CodeParameter;
	struct CodeProperty;
	struct CodeVariable;
	struct CodeStruct;
	struct CodeEnum;
	struct CodeDelegate;
	struct ContextAttribute;
	struct ContextAttributes;
	struct AddIn;
	struct AddIns;
	struct Events;
	struct Command;
	struct Commands;
	struct SelectedItem;
	struct SelectedItems;
	struct SelectionContainer;
	struct WindowConfigurations;
	struct Globals;
	struct StatusBar;
	struct ObjectExtenders;
	struct IExtenderProvider;
	struct IExtenderSite;
	struct IExtenderProviderUnk;
	struct ItemOperations;
	struct UndoContext;
	struct Macros;
	struct SourceControl;
	struct Debugger;
	struct Expression;
	struct Expressions;
	struct Breakpoint;
	struct Breakpoints;
	struct Program;
	struct Programs;
	struct Process;
	struct Processes;
	struct Thread;
	struct Threads;
	struct StackFrame;
	struct StackFrames;
	struct Language;
	struct Languages;
	struct Find;
	struct SolutionConfiguration;
	struct SolutionConfigurations;
	struct SolutionContext;
	struct SolutionContexts;
	struct BuildDependency;
	struct BuildDependencies;
	struct CodeModel;
	struct OutputGroup;
	struct OutputGroups;

	// ================================ OutputGroup ================================
#pragma region
	const GUID IID_OutputGroup = { 0xa3a80783, 0x875f, 0x435b, { 0x96, 0x39, 0xe5, 0xce, 0x88, 0x8d, 0xf7, 0x37 } };

	struct __declspec(novtable) OutputGroup {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == OutputGroup ==
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Collection(OutputGroups **ppOutputGroups) = 0;
		virtual HRESULT __stdcall get_FileNames(VARIANT *pNames) = 0;
		virtual HRESULT __stdcall get_FileCount(long *pCountNames) = 0;
		virtual HRESULT __stdcall get_DisplayName(BSTR *pName) = 0;
		virtual HRESULT __stdcall get_CanonicalName(BSTR *pName) = 0;
		virtual HRESULT __stdcall get_FileURLs(VARIANT *pURLs) = 0;
		virtual HRESULT __stdcall get_Description(BSTR *pDesc) = 0;
	};
#pragma endregion

	// ================================ OutputGroups ================================
#pragma region
	const GUID IID_OutputGroups = { 0xf9fa748e, 0xe302, 0x44cf, { 0x89, 0x1b, 0xe2, 0x63, 0x18, 0x9d, 0x58, 0x5e } };

	struct __declspec(novtable) OutputGroups {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == OutputGroups ==
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Parent(Configuration **ppConfiguration) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **Enumerator) = 0;
		virtual HRESULT __stdcall Item(VARIANT Index, OutputGroup **ppOut) = 0;
		virtual HRESULT __stdcall get_Count(long *pCount) = 0;
	};
#pragma endregion

	// ================================ CodeModel ================================
#pragma region
	const GUID IID_CodeModel = { 0x0CFBC2B4, 0x0D4E, 0x11D3, { 0x89, 0x97, 0x00, 0xC0, 0x4F, 0x68, 0x8D, 0xDE } };

	enum vsCMAccess {
		vsCMAccessPublic = 1,
		vsCMAccessPrivate = 2,
		vsCMAccessProject = 4,
		vsCMAccessProtected = 8,
		vsCMAccessDefault = 32,
		vsCMAccessAssemblyOrFamily = 64,
		vsCMAccessWithEvents = 128,
		vsCMAccessProjectOrProtected = 12
	};

	enum vsCMFunction {
		vsCMFunctionOther = 0,
		vsCMFunctionConstructor = 1,
		vsCMFunctionPropertyGet = 2,
		vsCMFunctionPropertyLet = 4,
		vsCMFunctionPropertySet = 8,
		vsCMFunctionPutRef = 16,
		vsCMFunctionPropertyAssign = 32,
		vsCMFunctionSub = 64,
		vsCMFunctionFunction = 128,
		vsCMFunctionTopLevel = 256,
		vsCMFunctionDestructor = 512,
		vsCMFunctionOperator = 1024,
		vsCMFunctionVirtual = 2048,
		vsCMFunctionPure = 4096,
		vsCMFunctionConstant = 8192,
		vsCMFunctionShared = 16384,
		vsCMFunctionInline = 32768,
		vsCMFunctionComMethod = 65536
	};

	struct __declspec(novtable) CodeModel {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeModel ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Parent(Project **pProj) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall get_CodeElements(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall CodeTypeFromFullName(BSTR Name, CodeType **ppCodeType) = 0;
		virtual HRESULT __stdcall AddNamespace(BSTR Name, VARIANT Location, VARIANT Position, CodeNamespace **ppCodeNamespace) = 0;
		virtual HRESULT __stdcall AddClass(BSTR Name, VARIANT Location, VARIANT Position, VARIANT Bases, VARIANT ImplementedInterfaces, vsCMAccess Access, CodeClass **ppCodeClass) = 0;
		virtual HRESULT __stdcall AddInterface(BSTR Name, VARIANT Location, VARIANT Position, VARIANT Bases, vsCMAccess Access, CodeInterface **ppCodeInterface) = 0;
		virtual HRESULT __stdcall AddFunction(BSTR Name, VARIANT Location, vsCMFunction Kind, VARIANT Type, VARIANT Position, vsCMAccess Access, CodeFunction **ppCodeFunction) = 0;
		virtual HRESULT __stdcall AddVariable(BSTR Name, VARIANT Location, VARIANT Type, VARIANT Position, vsCMAccess Access, CodeVariable **ppCodeVariable) = 0;
		virtual HRESULT __stdcall AddStruct(BSTR Name, VARIANT Location, VARIANT Position, VARIANT Bases, VARIANT ImplementedInterfaces, vsCMAccess Access, CodeStruct **ppCodeStruct) = 0;
		virtual HRESULT __stdcall AddEnum(BSTR Name, VARIANT Location, VARIANT Position, VARIANT Bases, vsCMAccess Access, CodeEnum **ppCodeEnum) = 0;
		virtual HRESULT __stdcall AddDelegate(BSTR Name, VARIANT Location, VARIANT Type, VARIANT Position, vsCMAccess Access, CodeDelegate **ppCodeDelegate) = 0;
		virtual HRESULT __stdcall AddAttribute(BSTR Name, VARIANT Location, BSTR Value, VARIANT Position, CodeAttribute **ppCodeAttribute) = 0;
		virtual HRESULT __stdcall Remove(VARIANT Element) = 0;
		virtual HRESULT __stdcall IsValidID(BSTR Name, VARIANT_BOOL *pValid) = 0;
		virtual HRESULT __stdcall get_IsCaseSensitive(VARIANT_BOOL *pSensitive) = 0;
		virtual HRESULT __stdcall CreateCodeTypeRef(VARIANT Type, CodeTypeRef **ppCodeTypeRef) = 0;
	};
#pragma endregion

	// ================================ BuildDependency ================================
#pragma region
	const GUID IID_BuildDependency = { 0x9c5ceaac, 0x062f, 0x4434, { 0xa2, 0xed, 0x78, 0xab, 0x4d, 0x61, 0x34, 0xfe } };

	struct __declspec(novtable) BuildDependency {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == BuildDependency ==
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Collection(BuildDependencies **ppBuildDependencies) = 0;
		virtual HRESULT __stdcall get_Project(Project **ppProject) = 0;
		virtual HRESULT __stdcall get_RequiredProjects(VARIANT *pProjects) = 0;
		virtual HRESULT __stdcall AddProject(BSTR ProjectUniqueName) = 0;
		virtual HRESULT __stdcall RemoveProject(BSTR ProjectUniqueName) = 0;
		virtual HRESULT __stdcall RemoveAllProjects() = 0;
	};
#pragma endregion

	// ================================ BuildDependencies ================================
#pragma region
	const GUID IID_BuildDependencies = { 0xead260eb, 0x1e5b, 0x450a, { 0xb6, 0x28, 0x4c, 0xfa, 0xda, 0x11, 0xb4, 0xa1 } };

	struct __declspec(novtable) BuildDependencies {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == BuildDependencies ==
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Parent(SolutionBuild **ppSolutionBuild) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **Enumerator) = 0;
		virtual HRESULT __stdcall Item(VARIANT Index, BuildDependency **ppOut) = 0;
		virtual HRESULT __stdcall get_Count(long *pCount) = 0;
	};
#pragma endregion

	// ================================ SolutionContext ================================
#pragma region
	const GUID IID_SolutionContext = { 0xfc6a1a82, 0x9c8a, 0x47bb, { 0xa0, 0x46, 0x6e, 0x96, 0x5d, 0xf5, 0xa9, 0x9b } };

	struct __declspec(novtable) SolutionContext {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == SolutionContext ==
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Collection(SolutionContexts **ppSolutionContexts) = 0;
		virtual HRESULT __stdcall get_ProjectName(BSTR *pName) = 0;
		virtual HRESULT __stdcall get_ConfigurationName(BSTR *pConfigurationName) = 0;
		virtual HRESULT __stdcall put_ConfigurationName(BSTR Name) = 0;
		virtual HRESULT __stdcall get_PlatformName(BSTR *pPlatformName) = 0;
		virtual HRESULT __stdcall get_ShouldBuild(VARIANT_BOOL *pPlatformName) = 0;
		virtual HRESULT __stdcall put_ShouldBuild(VARIANT_BOOL IsBuilt) = 0;
		virtual HRESULT __stdcall get_ShouldDeploy(VARIANT_BOOL *pPlatformName) = 0;
		virtual HRESULT __stdcall put_ShouldDeploy(VARIANT_BOOL IsDeployed) = 0;
	};
#pragma endregion

	// ================================ SolutionContexts ================================
#pragma region
	const GUID IID_SolutionContexts = { 0x0685b546, 0xfb84, 0x4917, { 0xab, 0x98, 0x98, 0xd4, 0x0f, 0x89, 0x2d, 0x61 } };

	struct __declspec(novtable) SolutionContexts {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == SolutionContexts ==
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Parent(SolutionConfiguration **ppSolutionConfiguration) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **Enumerator) = 0;
		virtual HRESULT __stdcall Item(VARIANT Index, SolutionContext **ppOut) = 0;
		virtual HRESULT __stdcall get_Count(long *pCount) = 0;
	};
#pragma endregion

	// ================================ SolutionConfigurations ================================
#pragma region
	const GUID IID_SolutionConfigurations = { 0x23e78ed7, 0xc9e1, 0x462d, { 0x8b, 0xc6, 0x36, 0x60, 0x03, 0x48, 0x6e, 0xd9 } };

	struct __declspec(novtable) SolutionConfigurations {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == SolutionConfigurations ==
		virtual HRESULT __stdcall _NewEnum(IUnknown **Enumerator) = 0;
		virtual HRESULT __stdcall Item(VARIANT Index, SolutionConfiguration **ppOut) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Parent(SolutionBuild **ppBuild) = 0;
		virtual HRESULT __stdcall Add(BSTR NewName, BSTR ExistingName, VARIANT_BOOL Propagate, SolutionConfiguration **ppSolutionConfiguration) = 0;
		virtual HRESULT __stdcall get_Count(long *pCount) = 0;
	};
#pragma endregion

	// ================================ SolutionConfiguration ================================
#pragma region
	const GUID IID_SolutionConfiguration = { 0x60aaad75, 0xcb8d, 0x4c62, { 0x99, 0x59, 0x24, 0xd6, 0xa6, 0xa5, 0x0d, 0xe7 } };

	struct __declspec(novtable) SolutionConfiguration {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == SolutionConfiguration ==
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Collection(SolutionConfigurations **ppSolutionConfigurations) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pVal) = 0;
		virtual HRESULT __stdcall get_SolutionContexts(SolutionContexts **ppOut) = 0;
		virtual HRESULT __stdcall Delete() = 0;
		virtual HRESULT __stdcall Activate() = 0;
	};
#pragma endregion

	// ================================ Find ================================
#pragma region
	const GUID IID_Find = { 0x40d4b9b6, 0x739b, 0x4965, { 0x8d, 0x65, 0x69, 0x2a, 0xec, 0x69, 0x22, 0x66 } };

	enum vsFindAction {
		vsFindActionFind = 1,
		vsFindActionFindAll = 2,
		vsFindActionReplace = 3,
		vsFindActionReplaceAll = 4,
		vsFindActionBookmarkAll = 5
	};

	enum vsFindPatternSyntax {
		vsFindPatternSyntaxLiteral = 0,
		vsFindPatternSyntaxRegExpr = 1,
		vsFindPatternSyntaxWildcards = 2
	};

	enum vsFindTarget {
		vsFindTargetCurrentDocument = 1,
		vsFindTargetCurrentDocumentSelection = 2,
		vsFindTargetCurrentDocumentFunction = 3,
		vsFindTargetOpenDocuments = 4,
		vsFindTargetCurrentProject = 5,
		vsFindTargetSolution = 6,
		vsFindTargetFiles = 7
	};

	enum vsFindResultsLocation {
		vsFindResultsNone = 0,
		vsFindResults1 = 1,
		vsFindResults2 = 2
	};

	enum vsFindResult {
		vsFindResultNotFound = 0,
		vsFindResultFound = 1,
		vsFindResultReplaceAndNotFound = 2,
		vsFindResultReplaceAndFound = 3,
		vsFindResultReplaced = 4,
		vsFindResultPending = 5,
		vsFindResultError = 6
	};

	struct __declspec(novtable) Find {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Find ==
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Parent(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Action(vsFindAction *pAction) = 0;
		virtual HRESULT __stdcall put_Action(vsFindAction Action) = 0;
		virtual HRESULT __stdcall get_FindWhat(BSTR *pFindWhat) = 0;
		virtual HRESULT __stdcall put_FindWhat(BSTR FindWhat) = 0;
		virtual HRESULT __stdcall get_MatchCase(VARIANT_BOOL *pMatchCase) = 0;
		virtual HRESULT __stdcall put_MatchCase(VARIANT_BOOL MatchCase) = 0;
		virtual HRESULT __stdcall get_MatchWholeWord(VARIANT_BOOL *pMatchWholeWord) = 0;
		virtual HRESULT __stdcall put_MatchWholeWord(VARIANT_BOOL MatchWholeWord) = 0;
		virtual HRESULT __stdcall get_MatchInHiddenText(VARIANT_BOOL *pMatchInHiddenText) = 0;
		virtual HRESULT __stdcall put_MatchInHiddenText(VARIANT_BOOL MatchInHiddenText) = 0;
		virtual HRESULT __stdcall get_Backwards(VARIANT_BOOL *pBackwards) = 0;
		virtual HRESULT __stdcall put_Backwards(VARIANT_BOOL Backwards) = 0;
		virtual HRESULT __stdcall get_SearchSubfolders(VARIANT_BOOL *pSearchSubfolders) = 0;
		virtual HRESULT __stdcall put_SearchSubfolders(VARIANT_BOOL SearchSubfolders) = 0;
		virtual HRESULT __stdcall get_KeepModifiedDocumentsOpen(VARIANT_BOOL *pKeepModifiedDocumentsOpen) = 0;
		virtual HRESULT __stdcall put_KeepModifiedDocumentsOpen(VARIANT_BOOL KeepModifiedDocumentsOpen) = 0;
		virtual HRESULT __stdcall get_PatternSyntax(vsFindPatternSyntax *pPatternSyntax) = 0;
		virtual HRESULT __stdcall put_PatternSyntax(vsFindPatternSyntax PatternSyntax) = 0;
		virtual HRESULT __stdcall get_ReplaceWith(BSTR *pReplaceWith) = 0;
		virtual HRESULT __stdcall put_ReplaceWith(BSTR ReplaceWith) = 0;
		virtual HRESULT __stdcall get_Target(vsFindTarget *pTarget) = 0;
		virtual HRESULT __stdcall put_Target(vsFindTarget Target) = 0;
		virtual HRESULT __stdcall get_SearchPath(BSTR *pSearchPath) = 0;
		virtual HRESULT __stdcall put_SearchPath(BSTR SearchPath) = 0;
		virtual HRESULT __stdcall get_FilesOfType(BSTR *pFilesOfType) = 0;
		virtual HRESULT __stdcall put_FilesOfType(BSTR FilesOfType) = 0;
		virtual HRESULT __stdcall get_ResultsLocation(vsFindResultsLocation *pResultsLocation) = 0;
		virtual HRESULT __stdcall put_ResultsLocation(vsFindResultsLocation ResultsLocation) = 0;
		virtual HRESULT __stdcall Execute(vsFindResult *pResult) = 0;
		virtual HRESULT __stdcall FindReplace(vsFindAction Action, BSTR FindWhat, long vsFindOptionsValue, BSTR ReplaceWith, vsFindTarget Target, BSTR SearchPath, BSTR FilesOfType, vsFindResultsLocation ResultsLocation, vsFindResult *pResult) = 0;
	};
#pragma endregion

	// ================================ Language ================================
#pragma region
	const GUID IID_Language = { 0xB3CCFA68, 0xC145, 0x11d2, { 0x8A, 0xD1, 0x00, 0xC0, 0x4F, 0x79, 0xE4, 0x79 } };

	struct __declspec(novtable) Language {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Language ==
		virtual HRESULT __stdcall get_Name(BSTR *Name) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTE) = 0;
		virtual HRESULT __stdcall get_Parent(Debugger **Debugger) = 0;
		virtual HRESULT __stdcall get_Collection(Languages **Languages) = 0;
	};
#pragma endregion

	// ================================ Languages ================================
#pragma region
	const GUID IID_Languages = { 0xa4f4246c, 0xc131, 0x11d2, { 0x8a, 0xd1, 0x00, 0xc0, 0x4f, 0x79, 0xe4, 0x79 } };

	struct __declspec(novtable) Languages {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Languages ==
		virtual HRESULT __stdcall Item(VARIANT Index, Language **Language) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **Enumerator) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Parent(Debugger **Debugger) = 0;
		virtual HRESULT __stdcall get_Count(long *Count) = 0;
	};
#pragma endregion

	// ================================ StackFrame ================================
#pragma region
	const GUID IID_StackFrame = { 0x1342D0D8, 0xBBA3, 0x11d2, { 0x8A, 0xD1, 0x00, 0xC0, 0x4F, 0x79, 0xE4, 0x79 } };

	struct __declspec(novtable) StackFrame {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == StackFrame ==
		virtual HRESULT __stdcall get_Language(BSTR *Language) = 0;
		virtual HRESULT __stdcall get_FunctionName(BSTR *FunctionName) = 0;
		virtual HRESULT __stdcall get_ReturnType(BSTR *ReturnType) = 0;
		virtual HRESULT __stdcall get_Locals(Expressions **Expressions) = 0;
		virtual HRESULT __stdcall get_Arguments(Expressions **Expressions) = 0;
		virtual HRESULT __stdcall get_Module(BSTR *Module) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTE) = 0;
		virtual HRESULT __stdcall get_Parent(Thread **Thread) = 0;
		virtual HRESULT __stdcall get_Collection(StackFrames **StackFrames) = 0;
	};
#pragma endregion

	// ================================ StackFrames ================================
#pragma region
	const GUID IID_StackFrames = { 0x4ed85664, 0xbba2, 0x11d2, { 0x8a, 0xd1, 0x00, 0xc0, 0x4f, 0x79, 0xe4, 0x79 } };

	struct __declspec(novtable) StackFrames {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == StackFrames ==
		virtual HRESULT __stdcall Item(VARIANT Index, StackFrame **StackFrame) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **Enumerator) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Parent(Debugger **Debugger) = 0;
		virtual HRESULT __stdcall get_Count(long *Count) = 0;
	};
#pragma endregion

	// ================================ Thread ================================
#pragma region
	const GUID IID_Thread = { 0x9407F466, 0xBBA1, 0x11d2, { 0x8A, 0xD1, 0x00, 0xC0, 0x4F, 0x79, 0xE4, 0x79 } };

	struct __declspec(novtable) Thread {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Thread ==
		virtual HRESULT __stdcall Freeze() = 0;
		virtual HRESULT __stdcall Thaw() = 0;
		virtual HRESULT __stdcall get_Name(BSTR *Name) = 0;
		virtual HRESULT __stdcall get_SuspendCount(long *Count) = 0;
		virtual HRESULT __stdcall get_ID(long *ID) = 0;
		virtual HRESULT __stdcall get_StackFrames(StackFrames **StackFrames) = 0;
		virtual HRESULT __stdcall get_Program(Program **Program) = 0;
		virtual HRESULT __stdcall get_IsAlive(VARIANT_BOOL *IsAlive) = 0;
		virtual HRESULT __stdcall get_Priority(BSTR *Priority) = 0;
		virtual HRESULT __stdcall get_Location(BSTR *Location) = 0;
		virtual HRESULT __stdcall get_IsFrozen(VARIANT_BOOL *IsFrozen) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTE) = 0;
		virtual HRESULT __stdcall get_Parent(Debugger **Debugger) = 0;
		virtual HRESULT __stdcall get_Collection(Threads **Threads) = 0;
	};
#pragma endregion

	// ================================ Threads ================================
#pragma region
	const GUID IID_Threads = { 0x6AA23FB4, 0xBBA1, 0x11d2, { 0x8A, 0xD1, 0x00, 0xC0, 0x4F, 0x79, 0xE4, 0x79 } };

	struct __declspec(novtable) Threads {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Threads ==
		virtual HRESULT __stdcall Item(VARIANT Index, Thread **Thread) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **Enumerator) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Parent(Debugger **Debugger) = 0;
		virtual HRESULT __stdcall get_Count(long *Count) = 0;
	};
#pragma endregion


	// ================================ Processes ================================
#pragma region
	const GUID IID_Processes = { 0x9F379969, 0x5EAC, 0x4a54, { 0xB2, 0xBC, 0x69, 0x46, 0xCF, 0xFB, 0x56, 0xEF } };

	struct __declspec(novtable) Processes {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Processes ==
		virtual HRESULT __stdcall Item(VARIANT Index, Process **Process) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **Enumerator) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Parent(Debugger **Debugger) = 0;
		virtual HRESULT __stdcall get_Count(long *Count) = 0;
	};
#pragma endregion


	// ================================ Programs ================================
#pragma region
	const GUID IID_Programs = { 0xDC6A118A, 0xBBAB, 0x11d2, { 0x8A, 0xD1, 0x00, 0xC0, 0x4F, 0x79, 0xE4, 0x79 } };

	struct __declspec(novtable) Programs {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Programs ==
		virtual HRESULT __stdcall Item(VARIANT Index, Program **Program) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **Enumerator) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Parent(Debugger **Debugger) = 0;
		virtual HRESULT __stdcall get_Count(long *Count) = 0;
	};
#pragma endregion

	// ================================ Process ================================
#pragma region
	const GUID IID_Process = { 0x5C5A0070, 0xF396, 0x4E37, { 0xA8, 0x2A, 0x1B, 0x76, 0x7E, 0x27, 0x2D, 0xF9 } };

	struct __declspec(novtable) Process {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Process ==
		virtual HRESULT __stdcall Attach() = 0;
		virtual HRESULT __stdcall Detach(VARIANT_BOOL WaitForBreakOrEnd) = 0;
		virtual HRESULT __stdcall Break(VARIANT_BOOL WaitForBreakMode) = 0;
		virtual HRESULT __stdcall Terminate(VARIANT_BOOL WaitForBreakOrEnd) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *Name) = 0;
		virtual HRESULT __stdcall get_ProcessID(long *ID) = 0;
		virtual HRESULT __stdcall get_Programs(Programs **Programs) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTE) = 0;
		virtual HRESULT __stdcall get_Parent(Debugger **Debugger) = 0;
		virtual HRESULT __stdcall get_Collection(Processes **Processes) = 0;
	};
#pragma endregion

	// ================================ Program ================================
#pragma region
	const GUID IID_Program = { 0x6a38d87c, 0xbba0, 0x11d2, { 0x8a, 0xd1, 0x00, 0xc0, 0x4f, 0x79, 0xe4, 0x79 } };

	struct __declspec(novtable) Program {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Program ==
		virtual HRESULT __stdcall get_Name(BSTR *Name) = 0;
		virtual HRESULT __stdcall get_Process(Process **Process) = 0;
		virtual HRESULT __stdcall get_Threads(Threads **Threads) = 0;
		virtual HRESULT __stdcall get_IsBeingDebugged(VARIANT_BOOL *IsBeingDebugged) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTE) = 0;
		virtual HRESULT __stdcall get_Parent(Debugger **Debugger) = 0;
		virtual HRESULT __stdcall get_Collection(Programs **Programs) = 0;
	};
#pragma endregion

	// ================================ Breakpoint ================================
#pragma region
	const GUID IID_Breakpoint = { 0x11C5114C, 0xBB00, 0x11d2, { 0x8A, 0xD1, 0x00, 0xC0, 0x4F, 0x79, 0xE4, 0x79 } };

	enum dbgBreakpointType {
		dbgBreakpointTypePending = 1,
		dbgBreakpointTypeBound = (dbgBreakpointTypePending + 1)
	};

	enum dbgBreakpointLocationType {
		dbgBreakpointLocationTypeNone = 1,
		dbgBreakpointLocationTypeFunction = (dbgBreakpointLocationTypeNone + 1),
		dbgBreakpointLocationTypeFile = (dbgBreakpointLocationTypeFunction + 1),
		dbgBreakpointLocationTypeData = (dbgBreakpointLocationTypeFile + 1),
		dbgBreakpointLocationTypeAddress = (dbgBreakpointLocationTypeData + 1)
	};

	enum dbgBreakpointConditionType {
		dbgBreakpointConditionTypeWhenTrue = 1,
		dbgBreakpointConditionTypeWhenChanged = (dbgBreakpointConditionTypeWhenTrue + 1)
	};

	enum dbgHitCountType {
		dbgHitCountTypeNone = 1,
		dbgHitCountTypeEqual = (dbgHitCountTypeNone + 1),
		dbgHitCountTypeGreaterOrEqual = (dbgHitCountTypeEqual + 1),
		dbgHitCountTypeMultiple = (dbgHitCountTypeGreaterOrEqual + 1)
	};

	struct __declspec(novtable) Breakpoint {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Breakpoint ==
		virtual HRESULT __stdcall Delete() = 0;
		virtual HRESULT __stdcall get_Type(dbgBreakpointType *Type) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *Name) = 0;
		virtual HRESULT __stdcall put_Name(BSTR Name) = 0;
		virtual HRESULT __stdcall get_LocationType(dbgBreakpointLocationType *LocationType) = 0;
		virtual HRESULT __stdcall get_FunctionName(BSTR *FunctionName) = 0;
		virtual HRESULT __stdcall get_FunctionLineOffset(long *Offset) = 0;
		virtual HRESULT __stdcall get_FunctionColumnOffset(long *Offset) = 0;
		virtual HRESULT __stdcall get_File(BSTR *File) = 0;
		virtual HRESULT __stdcall get_FileLine(long *Line) = 0;
		virtual HRESULT __stdcall get_FileColumn(long *Column) = 0;
		virtual HRESULT __stdcall get_ConditionType(dbgBreakpointConditionType *ConditionType) = 0;
		virtual HRESULT __stdcall get_Condition(BSTR *Condition) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *Language) = 0;
		virtual HRESULT __stdcall get_HitCountType(dbgHitCountType *HitCountType) = 0;
		virtual HRESULT __stdcall get_HitCountTarget(long *HitCountTarget) = 0;
		virtual HRESULT __stdcall get_Enabled(VARIANT_BOOL *Enabled) = 0;
		virtual HRESULT __stdcall put_Enabled(VARIANT_BOOL Enable) = 0;
		virtual HRESULT __stdcall get_Tag(BSTR *Tag) = 0;
		virtual HRESULT __stdcall put_Tag(BSTR Tag) = 0;
		virtual HRESULT __stdcall get_CurrentHits(long *CurHitCount) = 0;
		virtual HRESULT __stdcall get_Program(Program **Program) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTE) = 0;
		virtual HRESULT __stdcall get_Parent(Breakpoint **Breakpoint) = 0;
		virtual HRESULT __stdcall get_Collection(Breakpoints **Breakpoints) = 0;
		virtual HRESULT __stdcall get_Children(Breakpoints **Breakpoints) = 0;
		virtual HRESULT __stdcall ResetHitCount() = 0;
	};
#pragma endregion

	// ================================ Breakpoints ================================
#pragma region
	const GUID IID_Breakpoints = { 0x25968106, 0xBAFB, 0x11d2, { 0x8A, 0xD1, 0x00, 0xC0, 0x4F, 0x79, 0xE4, 0x79 } };

	struct __declspec(novtable) Breakpoints {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Breakpoints ==
		virtual HRESULT __stdcall Item(VARIANT Index, Breakpoint **Breakpoint) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **Enumerator) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Parent(Debugger **Debugger) = 0;
		virtual HRESULT __stdcall get_Count(long *Count) = 0;
		virtual HRESULT __stdcall Add(BSTR Function, BSTR File, long Line, long Column, BSTR Condition, dbgBreakpointConditionType ConditionType, BSTR Language, BSTR Data, long DataCount, BSTR Address, long HitCount, dbgHitCountType HitCountType, Breakpoints **Breakpoints) = 0;
	};
#pragma endregion

	// ================================ Expressions ================================
#pragma region
	const GUID IID_Expressions = { 0x2685337A, 0xBB9E, 0x11d2, { 0x8A, 0xD1, 0x00, 0xC0, 0x4F, 0x79, 0xE4, 0x79 } };

	struct __declspec(novtable) Expressions {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Expressions ==
		virtual HRESULT __stdcall Item(VARIANT Index, Expression **Expression) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **Enumerator) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Parent(Debugger **Debugger) = 0;
		virtual HRESULT __stdcall get_Count(long *Count) = 0;
	};
#pragma endregion

	// ================================ Expression ================================
#pragma region
	const GUID IID_Expression = { 0x27ADC812, 0xBB07, 0x11d2, { 0x8A, 0xD1, 0x00, 0xC0, 0x4F, 0x79, 0xE4, 0x79 } };

	struct __declspec(novtable) Expression {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Expression ==
		virtual HRESULT __stdcall get_Name(BSTR *Name) = 0;
		virtual HRESULT __stdcall get_Type(BSTR *Type) = 0;
		virtual HRESULT __stdcall get_DataMembers(Expressions **Expressions) = 0;
		virtual HRESULT __stdcall get_Value(BSTR *Value) = 0;
		virtual HRESULT __stdcall put_Value(BSTR NewValue) = 0;
		virtual HRESULT __stdcall get_IsValidValue(VARIANT_BOOL *ValidValue) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTE) = 0;
		virtual HRESULT __stdcall get_Parent(Debugger **Debugger) = 0;
		virtual HRESULT __stdcall get_Collection(Expressions **Expressions) = 0;
	};
#pragma endregion

	// ================================ Debugger ================================
#pragma region
	const GUID IID_Debugger = { 0x338FB9A0, 0xBAE5, 0x11d2, { 0x8A, 0xd1, 0x00, 0xC0, 0x4F, 0x79, 0xE4, 0x79 } };

	enum dbgDebugMode {
		dbgDesignMode = 1,
		dbgBreakMode = (dbgDesignMode + 1),
		dbgRunMode = (dbgBreakMode + 1)
	};

	enum dbgEventReason {
		dbgEventReasonNone = 1,
		dbgEventReasonGo = (dbgEventReasonNone + 1),
		dbgEventReasonAttachProgram = (dbgEventReasonGo + 1),
		dbgEventReasonDetachProgram = (dbgEventReasonAttachProgram + 1),
		dbgEventReasonLaunchProgram = (dbgEventReasonDetachProgram + 1),
		dbgEventReasonEndProgram = (dbgEventReasonLaunchProgram + 1),
		dbgEventReasonStopDebugging = (dbgEventReasonEndProgram + 1),
		dbgEventReasonStep = (dbgEventReasonStopDebugging + 1),
		dbgEventReasonBreakpoint = (dbgEventReasonStep + 1),
		dbgEventReasonExceptionThrown = (dbgEventReasonBreakpoint + 1),
		dbgEventReasonExceptionNotHandled = (dbgEventReasonExceptionThrown + 1),
		dbgEventReasonUserBreak = (dbgEventReasonExceptionNotHandled + 1),
		dbgEventReasonContextSwitch = (dbgEventReasonUserBreak + 1)
	};

	struct __declspec(novtable) Debugger {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Debugger ==
		virtual HRESULT __stdcall GetExpression(BSTR ExpressionText, VARIANT_BOOL UseAutoExpandRules, long Timeout, Expression **Expression) = 0;
		virtual HRESULT __stdcall DetachAll() = 0;
		virtual HRESULT __stdcall StepInto(VARIANT_BOOL WaitForBreakOrEnd) = 0;
		virtual HRESULT __stdcall StepOver(VARIANT_BOOL WaitForBreakOrEnd) = 0;
		virtual HRESULT __stdcall StepOut(VARIANT_BOOL WaitForBreakOrEnd) = 0;
		virtual HRESULT __stdcall Go(VARIANT_BOOL WaitForBreakOrEnd) = 0;
		virtual HRESULT __stdcall Break(VARIANT_BOOL WaitForBreakMode) = 0;
		virtual HRESULT __stdcall Stop(VARIANT_BOOL WaitForDesignMode) = 0;
		virtual HRESULT __stdcall SetNextStatement() = 0;
		virtual HRESULT __stdcall RunToCursor(VARIANT_BOOL WaitForBreakOrEnd) = 0;
		virtual HRESULT __stdcall ExecuteStatement(BSTR Statement, long Timeout, VARIANT_BOOL TreatAsExpression) = 0;
		virtual HRESULT __stdcall get_Breakpoints(Breakpoints **Breakpoints) = 0;
		virtual HRESULT __stdcall get_Languages(Languages **Languages) = 0;
		virtual HRESULT __stdcall get_CurrentMode(dbgDebugMode *Mode) = 0;
		virtual HRESULT __stdcall get_CurrentProcess(Process **Process) = 0;
		virtual HRESULT __stdcall put_CurrentProcess(Process *Process) = 0;
		virtual HRESULT __stdcall get_CurrentProgram(Program **Program) = 0;
		virtual HRESULT __stdcall put_CurrentProgram(Program *Program) = 0;
		virtual HRESULT __stdcall get_CurrentThread(Thread **Thread) = 0;
		virtual HRESULT __stdcall put_CurrentThread(Thread *Thread) = 0;
		virtual HRESULT __stdcall get_CurrentStackFrame(StackFrame **StackFrame) = 0;
		virtual HRESULT __stdcall put_CurrentStackFrame(StackFrame *StackFrame) = 0;
		virtual HRESULT __stdcall get_HexDisplayMode(VARIANT_BOOL *HexModeOn) = 0;
		virtual HRESULT __stdcall put_HexDisplayMode(VARIANT_BOOL HexModeOn) = 0;
		virtual HRESULT __stdcall get_HexInputMode(VARIANT_BOOL *HexModeOn) = 0;
		virtual HRESULT __stdcall put_HexInputMode(VARIANT_BOOL HexModeOn) = 0;
		virtual HRESULT __stdcall get_LastBreakReason(dbgEventReason *Reason) = 0;
		virtual HRESULT __stdcall get_BreakpointLastHit(Breakpoint **Breakpoint) = 0;
		virtual HRESULT __stdcall get_AllBreakpointsLastHit(Breakpoints **Breakpoints) = 0;
		virtual HRESULT __stdcall get_DebuggedProcesses(Processes **Processes) = 0;
		virtual HRESULT __stdcall get_LocalProcesses(Processes **Processes) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTE) = 0;
		virtual HRESULT __stdcall get_Parent(DTE **DTE) = 0;
		virtual HRESULT __stdcall TerminateAll() = 0;
	};
#pragma endregion

	// ================================ SourceControl ================================
#pragma region
	const GUID IID_SourceControl = { 0xf1ddc2c2, 0xdf76, 0x4ebb, { 0x9d, 0xe8, 0x48, 0xad, 0x25, 0x57, 0x06, 0x2c } };

	struct __declspec(novtable) SourceControl {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == SourceControl ==
		virtual HRESULT __stdcall get_DTE(DTE **ppDTEObject) = 0;
		virtual HRESULT __stdcall get_Parent(DTE **ppDTEObject) = 0;
		virtual HRESULT __stdcall IsItemUnderSCC(BSTR ItemName, VARIANT_BOOL *pfControlled) = 0;
		virtual HRESULT __stdcall IsItemCheckedOut(BSTR ItemName, VARIANT_BOOL *pfCheckedOut) = 0;
		virtual HRESULT __stdcall CheckOutItem(BSTR ItemName, VARIANT_BOOL *pfCheckedOut) = 0;
		virtual HRESULT __stdcall CheckOutItems(SAFEARRAY **ItemNames, VARIANT_BOOL *pfCheckedOut) = 0;
		virtual HRESULT __stdcall ExcludeItem(BSTR ProjectFile, BSTR ItemName) = 0;
		virtual HRESULT __stdcall ExcludeItems(BSTR ProjectFile, SAFEARRAY **ItemNames) = 0;
	};
#pragma endregion

	// ================================ Macros ================================
#pragma region
	const GUID IID_Macros = { 0xf9f99155, 0x6d4d, 0x49b1, { 0xad, 0x63, 0xc7, 0x8c, 0x3e, 0x8a, 0x59, 0x16 } };

	struct __declspec(novtable) Macros {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Macros ==
		virtual HRESULT __stdcall get_DTE(DTE **pDTE) = 0;
		virtual HRESULT __stdcall get_Parent(DTE **pDTE) = 0;
		virtual HRESULT __stdcall get_IsRecording(VARIANT_BOOL *vbIsRecording) = 0;
		virtual HRESULT __stdcall EmitMacroCode(BSTR Code) = 0;
		virtual HRESULT __stdcall Pause() = 0;
		virtual HRESULT __stdcall Resume() = 0;
	};
#pragma endregion


	// ================================ UndoContext ================================
#pragma region
	const GUID IID_UndoContext = { 0xd8dec44d, 0xcaf2, 0x4b39, { 0xa5, 0x39, 0xb9, 0x1a, 0xe9, 0x21, 0xba, 0x92 } };

	struct __declspec(novtable) UndoContext {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == UndoContext ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Parent(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall Open(BSTR Name, VARIANT_BOOL Strict) = 0;
		virtual HRESULT __stdcall Close() = 0;
		virtual HRESULT __stdcall SetAborted() = 0;
		virtual HRESULT __stdcall get_IsStrict(VARIANT_BOOL *pIsStrict) = 0;
		virtual HRESULT __stdcall get_IsAborted(VARIANT_BOOL *pIsAborted) = 0;
		virtual HRESULT __stdcall get_IsOpen(VARIANT_BOOL *pIsOpen) = 0;
	};
#pragma endregion

	// ================================ ItemOperations ================================
#pragma region
	const GUID IID_ItemOperations = { 0xd5dbe57b, 0xc074, 0x4e95, { 0xb0, 0x15, 0xab, 0xee, 0xaa, 0x39, 0x16, 0x93 } };

	enum vsNavigateOptions {
		vsNavigateOptionsDefault = 0,
		vsNavigateOptionsNewWindow = 0x1
	};

	enum vsPromptResult {
		vsPromptResultYes = 1,
		vsPromptResultNo = 2,
		vsPromptResultCancelled = 3
	};

	struct __declspec(novtable) ItemOperations {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == ItemOperations ==
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Parent(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall OpenFile(BSTR FileName, BSTR ViewKind, Window **Window) = 0;
		virtual HRESULT __stdcall NewFile(BSTR Item, BSTR Name, BSTR ViewKind, Window **Window) = 0;
		virtual HRESULT __stdcall IsFileOpen(BSTR FileName, BSTR ViewKind, VARIANT_BOOL *pfRetval) = 0;
		virtual HRESULT __stdcall AddExistingItem(BSTR FileName, ProjectItem **ProjectItem) = 0;
		virtual HRESULT __stdcall AddNewItem(BSTR Item, BSTR Name, ProjectItem **ProjectItem) = 0;
		virtual HRESULT __stdcall Navigate(BSTR URL, vsNavigateOptions Options, Window **Window) = 0;
		virtual HRESULT __stdcall get_PromptToSave(vsPromptResult *Saved) = 0;
	};
#pragma endregion

	// ================================== IExtenderProviderUnk ==================================
#pragma region
	const GUID IID_IExtenderProviderUnk = { 0xf69b64a3, 0x9017, 0x4e48, { 0x97, 0x84, 0xe1, 0x52, 0xb5, 0x1a, 0xa7, 0x22 } };

	struct __declspec(novtable) IExtenderProviderUnk {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == IExtenderProviderUnk ==
		virtual HRESULT __stdcall GetExtender(BSTR ExtenderCATID, BSTR ExtenderName, IUnknown *ExtendeeObject, IExtenderSite *ExtenderSite, long Cookie, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall CanExtend(BSTR ExtenderCATID, BSTR ExtenderName, IUnknown *ExtendeeObject, VARIANT_BOOL *fRetval) = 0;
	};
#pragma endregion

	// ================================== IExtenderSite ==================================
#pragma region
	const GUID IID_IExtenderSite = { 0xe57c510b, 0x968b, 0x4a3c, { 0xa4, 0x67, 0xee, 0x40, 0x13, 0x15, 0x7d, 0xc9 } };

	struct __declspec(novtable) IExtenderSite {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == IExtenderSite ==
		virtual HRESULT __stdcall NotifyDelete(long Cookie) = 0;
		virtual HRESULT __stdcall GetObject(BSTR Name, IDispatch **ppObject) = 0;
	};
#pragma endregion


	// ================================== IExtenderProvider ==================================
#pragma region
	const GUID IID_IExtenderProvider = { 0x4db06329, 0x23f4, 0x443b, { 0x9a, 0xbd, 0x9c, 0xf6, 0x11, 0xe8, 0xae, 0x07 } };

	struct __declspec(novtable) IExtenderProvider {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == IExtenderProvider ==
		virtual HRESULT __stdcall GetExtender(BSTR ExtenderCATID, BSTR ExtenderName, IDispatch *ExtendeeObject, IExtenderSite *ExtenderSite, long Cookie, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall CanExtend(BSTR ExtenderCATID, BSTR ExtenderName, IDispatch *ExtendeeObject, VARIANT_BOOL *fRetval) = 0;
	};
#pragma endregion

	// ================================== ObjectExtenders ==================================
#pragma region
	const GUID IID_ObjectExtenders = { 0x8d0aa9cc, 0x8465, 0x42f3, { 0xad, 0x6e, 0xdf, 0xde, 0x28, 0xcc, 0xc7, 0x5d } };

	struct __declspec(novtable) ObjectExtenders {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == ObjectExtenders ==
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Parent(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall RegisterExtenderProvider(BSTR ExtenderCATID, BSTR ExtenderName, IExtenderProvider *ExtenderProvider, BSTR LocalizedName, long *Cookie) = 0;
		virtual HRESULT __stdcall UnregisterExtenderProvider(long Cookie) = 0;
		virtual HRESULT __stdcall GetExtender(BSTR ExtenderCATID, BSTR ExtenderName, IUnknown *ExtendeeObject, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall GetExtenderNames(BSTR ExtenderCATID, IUnknown *ExtendeeObject, VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall GetContextualExtenderCATIDs(VARIANT *ExtenderCATIDs) = 0;
		virtual HRESULT __stdcall GetLocalizedExtenderName(BSTR ExtenderCATID, BSTR ExtenderName, BSTR *pLocalizedName) = 0;
		virtual HRESULT __stdcall RegisterExtenderProviderUnk(BSTR ExtenderCATID, BSTR ExtenderName, IExtenderProviderUnk *ExtenderProvider, BSTR LocalizedName, long *Cookie) = 0;
	};
#pragma endregion

	// ================================== StatusBar ==================================
#pragma region
	const GUID IID_StatusBar = { 0xc34301a1, 0x3ef1, 0x41d8, { 0x93, 0x2a, 0xfe, 0xa4, 0xa8, 0xa8, 0xce, 0x0c } };

	struct __declspec(novtable) StatusBar {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == StatusBar ==
		virtual HRESULT __stdcall get_DTE(DTE **pDTE) = 0;
		virtual HRESULT __stdcall get_Parent(DTE **pDTE) = 0;
		virtual HRESULT __stdcall Clear() = 0;
		virtual HRESULT __stdcall Animate(VARIANT_BOOL On, VARIANT AnimationType) = 0;
		virtual HRESULT __stdcall Progress(VARIANT_BOOL InProgress, BSTR Label = L"", long AmountCompleted = 0, long Total = 0) = 0;
		virtual HRESULT __stdcall SetXYWidthHeight(long X, long Y, long Width, long Height) = 0;
		virtual HRESULT __stdcall SetLineColumnCharacter(long Line, long Column, long Character) = 0;
		virtual HRESULT __stdcall put_Text(BSTR Text) = 0;
		virtual HRESULT __stdcall get_Text(BSTR *pText) = 0;
		virtual HRESULT __stdcall Highlight(VARIANT_BOOL Highlight) = 0;
		virtual HRESULT __stdcall ShowTextUpdates(VARIANT_BOOL TextUpdates, VARIANT_BOOL *WillShowUpdates) = 0;
	};
#pragma endregion

	// ================================== Globals ==================================
#pragma region
	const GUID IID_Globals = { 0xe68a3e0e, 0xb435, 0x4dde, { 0x86, 0xb7, 0xf5, 0xad, 0xef, 0xc1, 0x9d, 0xf2 } };

	struct __declspec(novtable) Globals {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Globals ==
		virtual HRESULT __stdcall get_DTE(DTE **pDTE) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **pDTE) = 0;
		virtual HRESULT __stdcall get_VariableValue(BSTR VariableName, VARIANT *pVal) = 0;
		virtual HRESULT __stdcall put_VariableValue(BSTR VariableName, VARIANT newVal) = 0;
		virtual HRESULT __stdcall put_VariablePersists(BSTR VariableName, VARIANT_BOOL pVal) = 0;
		virtual HRESULT __stdcall get_VariablePersists(BSTR VariableName, VARIANT_BOOL *pVal) = 0;
		virtual HRESULT __stdcall get_VariableExists(BSTR Name, VARIANT_BOOL *pVal) = 0;
		virtual HRESULT __stdcall get_VariableNames(VARIANT *Names) = 0;
	};
#pragma endregion

	// ================================== WindowConfiguration ==================================
#pragma region
	const GUID IID_WindowConfiguration = { 0x41d02413, 0x8a67, 0x4c28, { 0xa9, 0x80, 0xad, 0x75, 0x39, 0xed, 0x41, 0x5d } };

	struct __declspec(novtable) WindowConfiguration {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == WindowConfiguration ==
		virtual HRESULT __stdcall get_Name(BSTR *pName) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **pDTE) = 0;
		virtual HRESULT __stdcall get_Collection(WindowConfigurations **pWindowConfigurations) = 0;
		virtual HRESULT __stdcall Apply(VARIANT_BOOL FromCustomViews = -1) = 0;
		virtual HRESULT __stdcall Delete() = 0;
		virtual HRESULT __stdcall Update() = 0;
	};
#pragma endregion

	// ================================== WindowConfigurations ==================================
#pragma region
	const GUID IID_WindowConfigurations = { 0xe577442a, 0x98e1, 0x46c5, { 0xbd, 0x2e, 0xd2, 0x58, 0x07, 0xec, 0x81, 0xce } };

	struct __declspec(novtable) WindowConfigurations {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == WindowConfigurations ==
		virtual HRESULT __stdcall _NewEnum(IUnknown **ppEnum) = 0;
		virtual HRESULT __stdcall Item(VARIANT Index, WindowConfiguration **pWindowConfiguration) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **pDTE) = 0;
		virtual HRESULT __stdcall get_Parent(DTE **pParent) = 0;
		virtual HRESULT __stdcall get_Count(long *pCount) = 0;
		virtual HRESULT __stdcall Add(BSTR Name, WindowConfiguration **pWindowConfiguration) = 0;
		virtual HRESULT __stdcall get_ActiveConfigurationName(BSTR *pbstrName) = 0;
	};
#pragma endregion

	// ================================== SelectionContainer ==================================
#pragma region
	const GUID IID_SelectionContainer = { 0x02273422, 0x8dd4, 0x4a9f, { 0x8a, 0x8b, 0xd7, 0x04, 0x43, 0xd5, 0x10, 0xf4 } };

	struct __declspec(novtable) SelectionContainer {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == SelectionContainer ==
		virtual HRESULT __stdcall Item(VARIANT index, IDispatch **lplppReturn) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **lppiuReturn) = 0;
		virtual HRESULT __stdcall get_Count(long *lplReturn) = 0;
		virtual HRESULT __stdcall get_Parent(SelectedItems **lppReturn) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppReturn) = 0;
	};
#pragma endregion

	// ================================== SelectedItem ==================================
#pragma region
	const GUID IID_SelectedItem = { 0x049d2cdf, 0x3731, 0x4cb6, { 0xa2, 0x33, 0xbe, 0x97, 0xbc, 0xe9, 0x22, 0xd3 } };

	struct __declspec(novtable) SelectedItem {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == SelectedItem ==
		virtual HRESULT __stdcall get_Collection(SelectedItems **lppReturn) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppReturn) = 0;
		virtual HRESULT __stdcall get_Project(Project **lppReturn) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **lppReturn) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *lpbstrReturn) = 0;
		virtual HRESULT __stdcall get_InfoCount(short *lpnCount) = 0;
		virtual HRESULT __stdcall get_Info(short Index, VARIANT *lpbstrReturn) = 0;
	};
#pragma endregion

	// ================================== SelectedItems ==================================
#pragma region
	const GUID IID_SelectedItems = { 0x6caa67cf, 0x43ae, 0x4184, { 0xaa, 0xab, 0x02, 0x00, 0xdd, 0xf6, 0xb2, 0x40 } };

	struct __declspec(novtable) SelectedItems {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == SelectedItems ==
		virtual HRESULT __stdcall Item(VARIANT index, SelectedItem **lplppReturn) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **lppiuReturn) = 0;
		virtual HRESULT __stdcall get_Count(long *lplReturn) = 0;
		virtual HRESULT __stdcall get_Parent(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_MultiSelect(VARIANT_BOOL *pfMultiSelect) = 0;
		virtual HRESULT __stdcall get_SelectionContainer(SelectionContainer **lppdispSelContainer) = 0;
	};
#pragma endregion

	// ================================== Command ==================================
#pragma region
	const GUID IID_Command = { 0x5fe10fb0, 0x91a1, 0x4e55, { 0xba, 0xaa, 0xec, 0xca, 0xe5, 0xcc, 0xeb, 0x94 } };

	struct __declspec(novtable) Command {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Command ==
		virtual HRESULT __stdcall get_Name(BSTR *lpbstr) = 0;
		virtual HRESULT __stdcall get_Collection(Commands **lppcReturn) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_GUID(BSTR *lpbstr) = 0;
		virtual HRESULT __stdcall get_ID(long *lReturn) = 0;
		virtual HRESULT __stdcall get_IsAvailable(VARIANT_BOOL *pAvail) = 0;
		virtual HRESULT __stdcall AddControl(IDispatch *Owner, long Position, IDispatch **pCommandBarControl) = 0;
		virtual HRESULT __stdcall Delete() = 0;
		virtual HRESULT __stdcall get_Bindings(VARIANT *pVar) = 0;
		virtual HRESULT __stdcall put_Bindings(VARIANT Bindings) = 0;
		virtual HRESULT __stdcall get_LocalizedName(BSTR *lpbstr) = 0;
	};
#pragma endregion

	// ================================== Commands ==================================
#pragma region
	const GUID IID_Commands = { 0xe6b96cac, 0xb8c7, 0x40ae, { 0xb7, 0x05, 0x5c, 0x81, 0x87, 0x8c, 0x4a, 0x9e } };

	enum vsCommandBarType {
		vsCommandBarTypePopup = 10,
		vsCommandBarTypeToolbar = 23,
		vsCommandBarTypeMenu = 24
	};

	struct __declspec(novtable) Commands {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Commands ==
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Parent(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall Add(BSTR Guid, long ID, VARIANT *Control) = 0;
		virtual HRESULT __stdcall Raise(BSTR Guid, long ID, VARIANT *CustomIn, VARIANT *CustomOut) = 0;
		virtual HRESULT __stdcall CommandInfo(IDispatch *CommandBarControl, BSTR *Guid, long *ID) = 0;
		virtual HRESULT __stdcall get_Count(long *lplReturn) = 0;
		virtual HRESULT __stdcall Item(VARIANT Index, long ID, Command **lppcReturn) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **lppiuReturn) = 0;
		virtual HRESULT __stdcall AddNamedCommand(AddIn *AddInInstance, BSTR Name, BSTR ButtonText, BSTR Tooltip, VARIANT_BOOL MSOButton, long Bitmap, SAFEARRAY **ContextUIGUIDs, long vsCommandDisabledFlagsValue, Command **pVal) = 0;
		virtual HRESULT __stdcall AddCommandBar(BSTR Name, vsCommandBarType Type, IDispatch *CommandBarParent, long Position, IDispatch **pVal) = 0;
		virtual HRESULT __stdcall RemoveCommandBar(IDispatch *CommandBar) = 0;
	};
#pragma endregion

	// =================================== Events ===================================
#pragma region
	const GUID IID_Events = { 0x134170F8, 0x93B1, 0x42DD, { 0x9F, 0x89, 0xA2, 0xAC, 0x70, 0x10, 0xBA, 0x07 } };

	struct __declspec(novtable) Events {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Events ==
		virtual HRESULT __stdcall get_CommandBarEvents(IDispatch *CommandBarControl, IDispatch **prceNew) = 0;
		virtual HRESULT __stdcall get_CommandEvents(BSTR Guid, long ID, CommandEvents **ppceNew) = 0;
		virtual HRESULT __stdcall get_SelectionEvents(SelectionEvents **ppceNew) = 0;
		virtual HRESULT __stdcall get_SolutionEvents(SolutionEvents **ppceNew) = 0;
		virtual HRESULT __stdcall get_BuildEvents(BuildEvents **ppceNew) = 0;
		virtual HRESULT __stdcall get_WindowEvents(Window *WindowFilter, WindowEvents **ppceNew) = 0;
		virtual HRESULT __stdcall get_OutputWindowEvents(BSTR Pane, OutputWindowEvents **ppceNew) = 0;
		virtual HRESULT __stdcall get_FindEvents(FindEvents **ppFindEvents) = 0;
		virtual HRESULT __stdcall get_TaskListEvents(BSTR Filter, TaskListEvents **ppceNew) = 0;
		virtual HRESULT __stdcall get_DTEEvents(DTEEvents **ppceNew) = 0;
		virtual HRESULT __stdcall get_DocumentEvents(Document *Document, DocumentEvents **ppceNew) = 0;
		virtual HRESULT __stdcall get_SolutionItemsEvents(ProjectItemsEvents **ppeNew) = 0;
		virtual HRESULT __stdcall get_MiscFilesEvents(ProjectItemsEvents **ppeNew) = 0;
		virtual HRESULT __stdcall get_DebuggerEvents(DebuggerEvents **ppeNew) = 0;
		virtual HRESULT __stdcall get_TextEditorEvents(TextDocument *TextDocumentFilter, TextEditorEvents **ppeNew) = 0;
		virtual HRESULT __stdcall GetObject(BSTR Name, IDispatch **ppObject) = 0;
	};
#pragma endregion

	// =================================== AddIns ===================================
#pragma region
	const GUID IID_AddIns = { 0x50590801, 0xd13e, 0x4404, { 0x80, 0xc2, 0x5c, 0xa3, 0x0a, 0x4d, 0x0e, 0xe8 } };

	struct __declspec(novtable) AddIns {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == AddIns ==
		virtual HRESULT __stdcall Item(VARIANT Index, AddIn **lppaddin) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Count(long *lplReturn) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **lppiuReturn) = 0;
		virtual HRESULT __stdcall Update() = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall Add(BSTR ProgID, BSTR Description, BSTR Name, VARIANT_BOOL Connected, AddIn **__MIDL__AddIns0000) = 0;
	};
#pragma endregion

	// =================================== AddIn ===================================
#pragma region
	const GUID IID_AddIn = { 0x53a87fa1, 0xce93, 0x48bf, { 0x95, 0x8b, 0xc6, 0xda, 0x79, 0x3c, 0x50, 0x28 } };

	struct __declspec(novtable) AddIn {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == AddIn ==
		virtual HRESULT __stdcall get_Description(BSTR *lpbstr) = 0;
		virtual HRESULT __stdcall put_Description(BSTR bstr) = 0;
		virtual HRESULT __stdcall get_Collection(AddIns **lppaddins) = 0;
		virtual HRESULT __stdcall get_ProgID(BSTR *lpbstr) = 0;
		virtual HRESULT __stdcall get_Guid(BSTR *lpbstr) = 0;
		virtual HRESULT __stdcall get_Connected(VARIANT_BOOL *lpfConnect) = 0;
		virtual HRESULT __stdcall put_Connected(VARIANT_BOOL fConnect) = 0;
		virtual HRESULT __stdcall get_Object(IDispatch **lppdisp) = 0;
		virtual HRESULT __stdcall put_Object(IDispatch *_lpdispObject) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *lpbstr) = 0;
		virtual HRESULT __stdcall Remove() = 0;
		virtual HRESULT __stdcall get_SatelliteDllPath(BSTR *pbstrPath) = 0;
	};
#pragma endregion

	// =================================== ContextAttribute ===================================
#pragma region
	const GUID IID_ContextAttribute = { 0x1a6e2cb3, 0xb897, 0x42eb, { 0x96, 0xbe, 0xff, 0x0f, 0xdb, 0x65, 0xdb, 0x2f } };

	struct __declspec(novtable) ContextAttribute {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == ContextAttribute ==
		virtual HRESULT __stdcall get_Name(BSTR *pName) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Collection(ContextAttributes **ppCollection) = 0;
		virtual HRESULT __stdcall get_Values(VARIANT *pVal) = 0;
		virtual HRESULT __stdcall Remove() = 0;
	};
#pragma endregion

	// =================================== ContextAttributes ===================================
#pragma region
	const GUID IID_ContextAttributes = { 0x33c5ebb8, 0x244e, 0x449d, { 0x9c, 0xee, 0xfa, 0xd7, 0x0a, 0x77, 0x4e, 0x59 } };

	enum vsContextAttributeType {
		vsContextAttributeFilter = 1,
		vsContextAttributeLookup = 2,
		vsContextAttributeLookupF1 = 3
	};

	enum vsContextAttributes {
		vsContextAttributesGlobal = 1,
		vsContextAttributesWindow = 2,
		vsContextAttributesHighPriority = 3
	};

	struct __declspec(novtable) ContextAttributes {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == ContextAttributes ==
		virtual HRESULT __stdcall _NewEnum(IUnknown **lppiuReturn) = 0;
		virtual HRESULT __stdcall Item(VARIANT Index, ContextAttribute **ppVal) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Count(long *Count) = 0;
		virtual HRESULT __stdcall Add(BSTR AttributeName, BSTR AttributeValue, vsContextAttributeType Type, ContextAttribute **ppVal) = 0;
		virtual HRESULT __stdcall get_Type(vsContextAttributes *Type) = 0;
		virtual HRESULT __stdcall get_HighPriorityAttributes(ContextAttributes **ppVal) = 0;
		virtual HRESULT __stdcall Refresh() = 0;
	};
#pragma endregion

	// =================================== CodeDelegate ===================================
#pragma region
	const GUID IID_CodeDelegate = { 0xB1F42513, 0x91CD, 0x4D3A, { 0x8B, 0x25, 0xA3, 0x17, 0xD8, 0x03, 0x2B, 0x24 } };

	enum vsCMElement {
		vsCMElementOther = 0,
		vsCMElementClass = 1,
		vsCMElementFunction = 2,
		vsCMElementVariable = 3,
		vsCMElementProperty = 4,
		vsCMElementNamespace = 5,
		vsCMElementParameter = 6,
		vsCMElementAttribute = 7,
		vsCMElementInterface = 8,
		vsCMElementDelegate = 9,
		vsCMElementEnum = 10,
		vsCMElementStruct = 11,
		vsCMElementUnion = 12,
		vsCMElementLocalDeclStmt = 13,
		vsCMElementFunctionInvokeStmt = 14,
		vsCMElementPropertySetStmt = 15,
		vsCMElementAssignmentStmt = 16,
		vsCMElementInheritsStmt = 17,
		vsCMElementImplementsStmt = 18,
		vsCMElementOptionStmt = 19,
		vsCMElementVBAttributeStmt = 20,
		vsCMElementVBAttributeGroup = 21,
		vsCMElementEventsDeclaration = 22,
		vsCMElementUDTDecl = 23,
		vsCMElementDeclareDecl = 24,
		vsCMElementDefineStmt = 25,
		vsCMElementTypeDef = 26,
		vsCMElementIncludeStmt = 27,
		vsCMElementUsingStmt = 28,
		vsCMElementMacro = 29,
		vsCMElementMap = 30,
		vsCMElementIDLImport = 31,
		vsCMElementIDLImportLib = 32,
		vsCMElementIDLCoClass = 33,
		vsCMElementIDLLibrary = 34,
		vsCMElementImportStmt = 35,
		vsCMElementMapEntry = 36,
		vsCMElementVCBase = 37,
		vsCMElementEvent = 38,
		vsCMElementModule = 39
	};

	enum vsCMInfoLocation {
		vsCMInfoLocationProject = 1,
		vsCMInfoLocationExternal = 2,
		vsCMInfoLocationNone = 4,
		vsCMInfoLocationVirtual = 8
	};
	
	enum vsCMPart {
		vsCMPartName = 1,
		vsCMPartAttributes = 2,
		vsCMPartHeader = 4,
		vsCMPartWhole = 8,
		vsCMPartBody = 16,
		vsCMPartNavigate = 32,
		vsCMPartAttributesWithDelimiter = 68,
		vsCMPartBodyWithDelimiter = 80,
		vsCMPartHeaderWithAttributes = 6,
		vsCMPartWholeWithAttributes = 10
	};

	struct __declspec(novtable) CodeDelegate {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeDelegate ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Collection(CodeElements **ppCollection) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pVal) = 0;
		virtual HRESULT __stdcall put_Name(BSTR NewName) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *pVal) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **pProjItem) = 0;
		virtual HRESULT __stdcall get_Kind(vsCMElement *pCodeEltKind) = 0;
		virtual HRESULT __stdcall get_IsCodeType(VARIANT_BOOL *pIsCodeType) = 0;
		virtual HRESULT __stdcall get_InfoLocation(vsCMInfoLocation *pInfoLocation) = 0;
		virtual HRESULT __stdcall get_Children(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall get_StartPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_EndPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall GetStartPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall GetEndPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **ParentObject) = 0;
		virtual HRESULT __stdcall get_Namespace(CodeNamespace **ppCodeNamespace) = 0;
		virtual HRESULT __stdcall get_Bases(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Members(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall put_Access(vsCMAccess Access) = 0;
		virtual HRESULT __stdcall get_Access(vsCMAccess *pAccess) = 0;
		virtual HRESULT __stdcall get_Attributes(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_DocComment(BSTR *pDocComment) = 0;
		virtual HRESULT __stdcall put_DocComment(BSTR DocComment) = 0;
		virtual HRESULT __stdcall get_Comment(BSTR *pComment) = 0;
		virtual HRESULT __stdcall put_Comment(BSTR Comment) = 0;
		virtual HRESULT __stdcall AddBase(VARIANT Base, VARIANT Position, CodeElement **ppOut) = 0;
		virtual HRESULT __stdcall AddAttribute(BSTR Name, BSTR Value, VARIANT Position, CodeAttribute **ppCodeAttribute) = 0;
		virtual HRESULT __stdcall RemoveBase(VARIANT Element) = 0;
		virtual HRESULT __stdcall RemoveMember(VARIANT Element) = 0;
		virtual HRESULT __stdcall get_IsDerivedFrom(BSTR FullName, VARIANT_BOOL *pVal) = 0;
		virtual HRESULT __stdcall get_DerivedTypes(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_BaseClass(CodeClass **ppCodeClass) = 0;
		virtual HRESULT __stdcall get_Prototype(long Flags, BSTR *pPrototype) = 0;
		virtual HRESULT __stdcall get_Type(CodeTypeRef **pCodeTypeRef) = 0;
		virtual HRESULT __stdcall put_Type(CodeTypeRef *Type) = 0;
		virtual HRESULT __stdcall get_Parameters(CodeElements **ppParameters) = 0;
		virtual HRESULT __stdcall AddParameter(BSTR Name, VARIANT Type, VARIANT Position, CodeParameter **ppCodeParameter) = 0;
		virtual HRESULT __stdcall RemoveParameter(VARIANT Element) = 0;
	};
#pragma endregion

	// =================================== CodeEnum ===================================
#pragma region
	const GUID IID_CodeEnum = { 0xB1F42512, 0x91CD, 0x4D3A, { 0x8B, 0x25, 0xA3, 0x17, 0xD8, 0x03, 0x2B, 0x24 } };

	struct __declspec(novtable) CodeEnum {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeEnum ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Collection(CodeElements **ppCollection) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pVal) = 0;
		virtual HRESULT __stdcall put_Name(BSTR NewName) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *pVal) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **pProjItem) = 0;
		virtual HRESULT __stdcall get_Kind(vsCMElement *pCodeEltKind) = 0;
		virtual HRESULT __stdcall get_IsCodeType(VARIANT_BOOL *pIsCodeType) = 0;
		virtual HRESULT __stdcall get_InfoLocation(vsCMInfoLocation *pInfoLocation) = 0;
		virtual HRESULT __stdcall get_Children(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall get_StartPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_EndPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall GetStartPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall GetEndPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **ParentObject) = 0;
		virtual HRESULT __stdcall get_Namespace(CodeNamespace **ppCodeNamespace) = 0;
		virtual HRESULT __stdcall get_Bases(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Members(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall put_Access(vsCMAccess Access) = 0;
		virtual HRESULT __stdcall get_Access(vsCMAccess *pAccess) = 0;
		virtual HRESULT __stdcall get_Attributes(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_DocComment(BSTR *pDocComment) = 0;
		virtual HRESULT __stdcall put_DocComment(BSTR DocComment) = 0;
		virtual HRESULT __stdcall get_Comment(BSTR *pComment) = 0;
		virtual HRESULT __stdcall put_Comment(BSTR Comment) = 0;
		virtual HRESULT __stdcall AddBase(VARIANT Base, VARIANT Position, CodeElement **ppOut) = 0;
		virtual HRESULT __stdcall AddAttribute(BSTR Name, BSTR Value, VARIANT Position, CodeAttribute **ppCodeAttribute) = 0;
		virtual HRESULT __stdcall RemoveBase(VARIANT Element) = 0;
		virtual HRESULT __stdcall RemoveMember(VARIANT Element) = 0;
		virtual HRESULT __stdcall get_IsDerivedFrom(BSTR FullName, VARIANT_BOOL *pVal) = 0;
		virtual HRESULT __stdcall get_DerivedTypes(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall AddMember(BSTR Name, VARIANT Value, VARIANT Position, CodeVariable **ppCodeElements) = 0;
	};
#pragma endregion

	// =================================== CodeStruct ===================================
#pragma region
	const GUID IID_CodeStruct = { 0xB1F42511, 0x91CD, 0x4D3A, { 0x8B, 0x25, 0xA3, 0x17, 0xD8, 0x03, 0x2B, 0x24 } };

	struct __declspec(novtable) CodeStruct {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeStruct ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Collection(CodeElements **ppCollection) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pVal) = 0;
		virtual HRESULT __stdcall put_Name(BSTR NewName) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *pVal) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **pProjItem) = 0;
		virtual HRESULT __stdcall get_Kind(vsCMElement *pCodeEltKind) = 0;
		virtual HRESULT __stdcall get_IsCodeType(VARIANT_BOOL *pIsCodeType) = 0;
		virtual HRESULT __stdcall get_InfoLocation(vsCMInfoLocation *pInfoLocation) = 0;
		virtual HRESULT __stdcall get_Children(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall get_StartPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_EndPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall GetStartPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall GetEndPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **ParentObject) = 0;
		virtual HRESULT __stdcall get_Namespace(CodeNamespace **ppCodeNamespace) = 0;
		virtual HRESULT __stdcall get_Bases(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Members(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall put_Access(vsCMAccess Access) = 0;
		virtual HRESULT __stdcall get_Access(vsCMAccess *pAccess) = 0;
		virtual HRESULT __stdcall get_Attributes(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_DocComment(BSTR *pDocComment) = 0;
		virtual HRESULT __stdcall put_DocComment(BSTR DocComment) = 0;
		virtual HRESULT __stdcall get_Comment(BSTR *pComment) = 0;
		virtual HRESULT __stdcall put_Comment(BSTR Comment) = 0;
		virtual HRESULT __stdcall AddBase(VARIANT Base, VARIANT Position, CodeElement **ppOut) = 0;
		virtual HRESULT __stdcall AddAttribute(BSTR Name, BSTR Value, VARIANT Position, CodeAttribute **ppCodeAttribute) = 0;
		virtual HRESULT __stdcall RemoveBase(VARIANT Element) = 0;
		virtual HRESULT __stdcall RemoveMember(VARIANT Element) = 0;
		virtual HRESULT __stdcall get_IsDerivedFrom(BSTR FullName, VARIANT_BOOL *pVal) = 0;
		virtual HRESULT __stdcall get_DerivedTypes(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_ImplementedInterfaces(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_IsAbstract(VARIANT_BOOL *pIsAbstract) = 0;
		virtual HRESULT __stdcall put_IsAbstract(VARIANT_BOOL Abstract) = 0;
		virtual HRESULT __stdcall AddImplementedInterface(VARIANT Base, VARIANT Position, CodeInterface **ppCodeInterface) = 0;
		virtual HRESULT __stdcall AddFunction(BSTR Name, vsCMFunction Kind, VARIANT Type, VARIANT Position, vsCMAccess Access, VARIANT Location, CodeFunction **ppCodeFunction) = 0;
		virtual HRESULT __stdcall AddVariable(BSTR Name, VARIANT Type, VARIANT Position, vsCMAccess Access, VARIANT Location, CodeVariable **ppCodeVariable) = 0;
		virtual HRESULT __stdcall AddProperty(BSTR GetterName, BSTR PutterName, VARIANT Type, VARIANT Position, vsCMAccess Access, VARIANT Location, CodeProperty **ppCodeProperty) = 0;
		virtual HRESULT __stdcall AddClass(BSTR Name, VARIANT Position, VARIANT Bases, VARIANT ImplementedInterfaces, vsCMAccess Access, CodeClass **ppCodeClass) = 0;
		virtual HRESULT __stdcall AddStruct(BSTR Name, VARIANT Position, VARIANT Bases, VARIANT ImplementedInterfaces, vsCMAccess Access, CodeStruct **ppCodeStruct) = 0;
		virtual HRESULT __stdcall AddEnum(BSTR Name, VARIANT Position, VARIANT Bases, vsCMAccess Access, CodeEnum **ppCodeEnum) = 0;
		virtual HRESULT __stdcall AddDelegate(BSTR Name, VARIANT Type, VARIANT Position, vsCMAccess Access, CodeDelegate **ppCodeDelegate) = 0;
		virtual HRESULT __stdcall RemoveInterface(VARIANT Element) = 0;
	};
#pragma endregion

	// =================================== CodeVariable ===================================
#pragma region
	const GUID IID_CodeVariable = { 0x0CFBC2BA, 0x0D4E, 0x11D3, { 0x89, 0x97, 0x00, 0xC0, 0x4F, 0x68, 0x8D, 0xDE } };

	struct __declspec(novtable) CodeVariable {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeVariable ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Collection(CodeElements **ppCollection) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pVal) = 0;
		virtual HRESULT __stdcall put_Name(BSTR NewName) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *pVal) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **pProjItem) = 0;
		virtual HRESULT __stdcall get_Kind(vsCMElement *pCodeEltKind) = 0;
		virtual HRESULT __stdcall get_IsCodeType(VARIANT_BOOL *pIsCodeType) = 0;
		virtual HRESULT __stdcall get_InfoLocation(vsCMInfoLocation *pInfoLocation) = 0;
		virtual HRESULT __stdcall get_Children(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall get_StartPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_EndPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall GetStartPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall GetEndPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **ParentObject) = 0;
		virtual HRESULT __stdcall get_InitExpression(VARIANT *pExpr) = 0;
		virtual HRESULT __stdcall put_InitExpression(VARIANT Expr) = 0;
		virtual HRESULT __stdcall get_Prototype(long Flags, BSTR *pVal) = 0;
		virtual HRESULT __stdcall put_Type(CodeTypeRef *pCodeTypeRef) = 0;
		virtual HRESULT __stdcall get_Type(CodeTypeRef **pCodeTypeRef) = 0;
		virtual HRESULT __stdcall put_Access(vsCMAccess Access) = 0;
		virtual HRESULT __stdcall get_Access(vsCMAccess *Access) = 0;
		virtual HRESULT __stdcall get_IsConstant(VARIANT_BOOL *pIsConstant) = 0;
		virtual HRESULT __stdcall put_IsConstant(VARIANT_BOOL IsConstant) = 0;
		virtual HRESULT __stdcall get_Attributes(CodeElements **ppMembers) = 0;
		virtual HRESULT __stdcall get_DocComment(BSTR *pDocComment) = 0;
		virtual HRESULT __stdcall put_DocComment(BSTR DocComment) = 0;
		virtual HRESULT __stdcall get_Comment(BSTR *pComment) = 0;
		virtual HRESULT __stdcall put_Comment(BSTR Comment) = 0;
		virtual HRESULT __stdcall AddAttribute(BSTR Name, BSTR Value, VARIANT Position, CodeAttribute **ppCodeAttribute) = 0;
		virtual HRESULT __stdcall get_IsShared(VARIANT_BOOL *pShared) = 0;
		virtual HRESULT __stdcall put_IsShared(VARIANT_BOOL Shared) = 0;
	};
#pragma endregion

	// =================================== CodeProperty ===================================
#pragma region
	const GUID IID_CodeProperty = { 0x0CFBC2BB, 0x0D4E, 0x11D3, { 0x89, 0x97, 0x00, 0xC0, 0x4F, 0x68, 0x8D, 0xDE } };

	struct __declspec(novtable) CodeProperty {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeProperty ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Collection(CodeElements **ppCollection) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pVal) = 0;
		virtual HRESULT __stdcall put_Name(BSTR NewName) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *pVal) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **pProjItem) = 0;
		virtual HRESULT __stdcall get_Kind(vsCMElement *pCodeEltKind) = 0;
		virtual HRESULT __stdcall get_IsCodeType(VARIANT_BOOL *pIsCodeType) = 0;
		virtual HRESULT __stdcall get_InfoLocation(vsCMInfoLocation *pInfoLocation) = 0;
		virtual HRESULT __stdcall get_Children(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall get_StartPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_EndPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall GetStartPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall GetEndPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_Parent(CodeClass **ParentObject) = 0;
		virtual HRESULT __stdcall get_Prototype(long Flags, BSTR *pFullName) = 0;
		virtual HRESULT __stdcall put_Type(CodeTypeRef *pCodeTypeRef) = 0;
		virtual HRESULT __stdcall get_Type(CodeTypeRef **pCodeTypeRef) = 0;
		virtual HRESULT __stdcall get_Getter(CodeFunction **ppCodeFunction) = 0;
		virtual HRESULT __stdcall put_Getter(CodeFunction *pCodeFunction) = 0;
		virtual HRESULT __stdcall get_Setter(CodeFunction **ppCodeFunction) = 0;
		virtual HRESULT __stdcall put_Setter(CodeFunction *pCodeFunction) = 0;
		virtual HRESULT __stdcall put_Access(vsCMAccess Access) = 0;
		virtual HRESULT __stdcall get_Access(vsCMAccess *Access) = 0;
		virtual HRESULT __stdcall get_Attributes(CodeElements **ppMembers) = 0;
		virtual HRESULT __stdcall get_DocComment(BSTR *pDocComment) = 0;
		virtual HRESULT __stdcall put_DocComment(BSTR DocComment) = 0;
		virtual HRESULT __stdcall get_Comment(BSTR *pComment) = 0;
		virtual HRESULT __stdcall put_Comment(BSTR Comment) = 0;
		virtual HRESULT __stdcall AddAttribute(BSTR Name, BSTR Value, VARIANT Position, CodeAttribute **ppCodeAttribute) = 0;
	};
#pragma endregion

	// =================================== CodeParameter ===================================
#pragma region
	const GUID IID_CodeParameter = { 0x0CFBC2BD, 0x0D4E, 0x11D3, { 0x89, 0x97, 0x00, 0xC0, 0x4F, 0x68, 0x8D, 0xDE } };

	struct __declspec(novtable) CodeParameter {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeParameter ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Collection(CodeElements **ppCollection) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pVal) = 0;
		virtual HRESULT __stdcall put_Name(BSTR NewName) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *pVal) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **pProjItem) = 0;
		virtual HRESULT __stdcall get_Kind(vsCMElement *pCodeEltKind) = 0;
		virtual HRESULT __stdcall get_IsCodeType(VARIANT_BOOL *pIsCodeType) = 0;
		virtual HRESULT __stdcall get_InfoLocation(vsCMInfoLocation *pInfoLocation) = 0;
		virtual HRESULT __stdcall get_Children(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall get_StartPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_EndPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall GetStartPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall GetEndPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_Parent(CodeElement **ppCodeElement) = 0;
		virtual HRESULT __stdcall put_Type(CodeTypeRef *Type) = 0;
		virtual HRESULT __stdcall get_Type(CodeTypeRef **pCodeTypeRef) = 0;
		virtual HRESULT __stdcall get_Attributes(CodeElements **ppMembers) = 0;
		virtual HRESULT __stdcall get_DocComment(BSTR *pDocComment) = 0;
		virtual HRESULT __stdcall put_DocComment(BSTR DocComment) = 0;
		virtual HRESULT __stdcall AddAttribute(BSTR Name, BSTR Value, VARIANT Position, CodeAttribute **ppCodeAttribute) = 0;
	};
#pragma endregion

	// =================================== CodeType ===================================
#pragma region
	const GUID IID_CodeType = { 0x0CFBC2B7, 0x0D4E, 0x11D3, { 0x89, 0x97, 0x00, 0xC0, 0x4F, 0x68, 0x8D, 0xDE } };

	struct __declspec(novtable) CodeType {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeType ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Collection(CodeElements **ppCollection) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pVal) = 0;
		virtual HRESULT __stdcall put_Name(BSTR NewName) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *pVal) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **pProjItem) = 0;
		virtual HRESULT __stdcall get_Kind(vsCMElement *pCodeEltKind) = 0;
		virtual HRESULT __stdcall get_IsCodeType(VARIANT_BOOL *pIsCodeType) = 0;
		virtual HRESULT __stdcall get_InfoLocation(vsCMInfoLocation *pInfoLocation) = 0;
		virtual HRESULT __stdcall get_Children(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall get_StartPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_EndPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall GetStartPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall GetEndPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **ParentObject) = 0;
		virtual HRESULT __stdcall get_Namespace(CodeNamespace **ppCodeNamespace) = 0;
		virtual HRESULT __stdcall get_Bases(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Members(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall put_Access(vsCMAccess Access) = 0;
		virtual HRESULT __stdcall get_Access(vsCMAccess *pAccess) = 0;
		virtual HRESULT __stdcall get_Attributes(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_DocComment(BSTR *pDocComment) = 0;
		virtual HRESULT __stdcall put_DocComment(BSTR DocComment) = 0;
		virtual HRESULT __stdcall get_Comment(BSTR *pComment) = 0;
		virtual HRESULT __stdcall put_Comment(BSTR Comment) = 0;
		virtual HRESULT __stdcall AddBase(VARIANT Base, VARIANT Position, CodeElement **ppOut) = 0;
		virtual HRESULT __stdcall AddAttribute(BSTR Name, BSTR Value, VARIANT Position, CodeAttribute **ppCodeAttribute) = 0;
		virtual HRESULT __stdcall RemoveBase(VARIANT Element) = 0;
		virtual HRESULT __stdcall RemoveMember(VARIANT Element) = 0;
		virtual HRESULT __stdcall get_IsDerivedFrom(BSTR FullName, VARIANT_BOOL *pVal) = 0;
		virtual HRESULT __stdcall get_DerivedTypes(CodeElements **ppCodeElements) = 0;
	};
#pragma endregion


	// =================================== CodeTypeRef ===================================
#pragma region
	const GUID IID_CodeTypeRef = { 0x0CFBC2BC, 0x0D4E, 0x11D3, { 0x89, 0x97, 0x00, 0xC0, 0x4F, 0x68, 0x8D, 0xDE } };

	enum vsCMTypeRef {
		vsCMTypeRefOther = 0,
		vsCMTypeRefCodeType = 1,
		vsCMTypeRefArray = 2,
		vsCMTypeRefVoid = 3,
		vsCMTypeRefPointer = 4,
		vsCMTypeRefString = 5,
		vsCMTypeRefObject = 6,
		vsCMTypeRefByte = 7,
		vsCMTypeRefChar = 8,
		vsCMTypeRefShort = 9,
		vsCMTypeRefInt = 10,
		vsCMTypeRefLong = 11,
		vsCMTypeRefFloat = 12,
		vsCMTypeRefDouble = 13,
		vsCMTypeRefDecimal = 14,
		vsCMTypeRefBool = 15,
		vsCMTypeRefVariant = 16
	};

	struct __declspec(novtable) CodeTypeRef {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeTypeRef ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **ParentObject) = 0;
		virtual HRESULT __stdcall get_TypeKind(vsCMTypeRef *pType) = 0;
		virtual HRESULT __stdcall get_CodeType(CodeType **ppCodeType) = 0;
		virtual HRESULT __stdcall put_CodeType(CodeType *Type) = 0;
		virtual HRESULT __stdcall get_ElementType(CodeTypeRef **ppCodeTypeRef) = 0;
		virtual HRESULT __stdcall put_ElementType(CodeTypeRef *Type) = 0;
		virtual HRESULT __stdcall get_AsString(BSTR *pAsString) = 0;
		virtual HRESULT __stdcall get_AsFullName(BSTR *pAsFullName) = 0;
		virtual HRESULT __stdcall get_Rank(long *pRank) = 0;
		virtual HRESULT __stdcall put_Rank(long Rank) = 0;
		virtual HRESULT __stdcall CreateArrayType(long Rank, CodeTypeRef **ppTypeRef) = 0;
	};
#pragma endregion

	// =================================== CodeFunction ===================================
#pragma region
	const GUID IID_CodeFunction = { 0x0CFBC2B9, 0x0D4E, 0x11D3, { 0x89, 0x97, 0x00, 0xC0, 0x4F, 0x68, 0x8D, 0xDE } };

	struct __declspec(novtable) CodeFunction {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeFunction ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Collection(CodeElements **ppCollection) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pVal) = 0;
		virtual HRESULT __stdcall put_Name(BSTR NewName) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *pVal) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **pProjItem) = 0;
		virtual HRESULT __stdcall get_Kind(vsCMElement *pCodeEltKind) = 0;
		virtual HRESULT __stdcall get_IsCodeType(VARIANT_BOOL *pIsCodeType) = 0;
		virtual HRESULT __stdcall get_InfoLocation(vsCMInfoLocation *pInfoLocation) = 0;
		virtual HRESULT __stdcall get_Children(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall get_StartPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_EndPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall GetStartPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall GetEndPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **ParentObject) = 0;
		virtual HRESULT __stdcall get_FunctionKind(vsCMFunction *ppFunctionKind) = 0;
		virtual HRESULT __stdcall get_Prototype(long Flags, BSTR *pFullName) = 0;
		virtual HRESULT __stdcall get_Type(CodeTypeRef **ppCodeTypeRef) = 0;
		virtual HRESULT __stdcall put_Type(CodeTypeRef *pCodeTypeRef) = 0;
		virtual HRESULT __stdcall get_Parameters(CodeElements **ppMembers) = 0;
		virtual HRESULT __stdcall put_Access(vsCMAccess Access) = 0;
		virtual HRESULT __stdcall get_Access(vsCMAccess *Access) = 0;
		virtual HRESULT __stdcall get_IsOverloaded(VARIANT_BOOL *pvbOverloaded) = 0;
		virtual HRESULT __stdcall get_IsShared(VARIANT_BOOL *Shared) = 0;
		virtual HRESULT __stdcall put_IsShared(VARIANT_BOOL Shared) = 0;
		virtual HRESULT __stdcall get_MustImplement(VARIANT_BOOL *MustImplement) = 0;
		virtual HRESULT __stdcall put_MustImplement(VARIANT_BOOL MustImplement) = 0;
		virtual HRESULT __stdcall get_Overloads(CodeElements **ppMembers) = 0;
		virtual HRESULT __stdcall get_Attributes(CodeElements **ppMembers) = 0;
		virtual HRESULT __stdcall get_DocComment(BSTR *pDocComment) = 0;
		virtual HRESULT __stdcall put_DocComment(BSTR DocComment) = 0;
		virtual HRESULT __stdcall get_Comment(BSTR *pComment) = 0;
		virtual HRESULT __stdcall put_Comment(BSTR Comment) = 0;
		virtual HRESULT __stdcall AddParameter(BSTR Name, VARIANT Type, VARIANT Position, CodeParameter **ppCodeParameter) = 0;
		virtual HRESULT __stdcall AddAttribute(BSTR Name, BSTR Value, VARIANT Position, CodeAttribute **ppCodeAttribute) = 0;
		virtual HRESULT __stdcall RemoveParameter(VARIANT Element) = 0;
		virtual HRESULT __stdcall get_CanOverride(VARIANT_BOOL *pCanOverride) = 0;
		virtual HRESULT __stdcall put_CanOverride(VARIANT_BOOL CanOverride) = 0;
	};
#pragma endregion

	// ============================== CodeInterface ==============================
#pragma region
	const GUID IID_CodeInterface = { 0xB1F42510, 0x91CD, 0x4D3A, { 0x8B, 0x25, 0xA3, 0x17, 0xD8, 0x03, 0x2B, 0x24 } };

	struct __declspec(novtable) CodeInterface {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeInterface ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Collection(CodeElements **ppCollection) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pVal) = 0;
		virtual HRESULT __stdcall put_Name(BSTR NewName) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *pVal) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **pProjItem) = 0;
		virtual HRESULT __stdcall get_Kind(vsCMElement *pCodeEltKind) = 0;
		virtual HRESULT __stdcall get_IsCodeType(VARIANT_BOOL *pIsCodeType) = 0;
		virtual HRESULT __stdcall get_InfoLocation(vsCMInfoLocation *pInfoLocation) = 0;
		virtual HRESULT __stdcall get_Children(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall get_StartPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_EndPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall GetStartPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall GetEndPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **ParentObject) = 0;
		virtual HRESULT __stdcall get_Namespace(CodeNamespace **ppCodeNamespace) = 0;
		virtual HRESULT __stdcall get_Bases(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Members(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall put_Access(vsCMAccess Access) = 0;
		virtual HRESULT __stdcall get_Access(vsCMAccess *pAccess) = 0;
		virtual HRESULT __stdcall get_Attributes(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_DocComment(BSTR *pDocComment) = 0;
		virtual HRESULT __stdcall put_DocComment(BSTR DocComment) = 0;
		virtual HRESULT __stdcall get_Comment(BSTR *pComment) = 0;
		virtual HRESULT __stdcall put_Comment(BSTR Comment) = 0;
		virtual HRESULT __stdcall AddBase(VARIANT Base, VARIANT Position, CodeElement **ppOut) = 0;
		virtual HRESULT __stdcall AddAttribute(BSTR Name, BSTR Value, VARIANT Position, CodeAttribute **ppCodeAttribute) = 0;
		virtual HRESULT __stdcall RemoveBase(VARIANT Element) = 0;
		virtual HRESULT __stdcall RemoveMember(VARIANT Element) = 0;
		virtual HRESULT __stdcall get_IsDerivedFrom(BSTR FullName, VARIANT_BOOL *pVal) = 0;
		virtual HRESULT __stdcall get_DerivedTypes(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall AddFunction(BSTR Name, vsCMFunction Kind, VARIANT Type, VARIANT Position, vsCMAccess Access, CodeFunction **ppCodeFunction) = 0;
		virtual HRESULT __stdcall AddProperty(BSTR GetterName, BSTR PutterName, VARIANT Type, VARIANT Position, vsCMAccess Access, VARIANT Location, CodeProperty **ppCodeProperty) = 0;
	};
#pragma endregion

	// ============================== CodeAttribute ==============================
#pragma region
	const GUID IID_CodeAttribute = { 0x0CFBC2BE, 0x0D4E, 0x11D3, { 0x89, 0x97, 0x00, 0xC0, 0x4F, 0x68, 0x8D, 0xDE } };

	struct __declspec(novtable) CodeAttribute {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeAttribute ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Collection(CodeElements **ppCollection) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pVal) = 0;
		virtual HRESULT __stdcall put_Name(BSTR NewName) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *pVal) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **pProjItem) = 0;
		virtual HRESULT __stdcall get_Kind(vsCMElement *pCodeEltKind) = 0;
		virtual HRESULT __stdcall get_IsCodeType(VARIANT_BOOL *pIsCodeType) = 0;
		virtual HRESULT __stdcall get_InfoLocation(vsCMInfoLocation *pInfoLocation) = 0;
		virtual HRESULT __stdcall get_Children(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall get_StartPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_EndPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall GetStartPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall GetEndPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **ppParent) = 0;
		virtual HRESULT __stdcall get_Value(BSTR *pValue) = 0;
		virtual HRESULT __stdcall put_Value(BSTR Value) = 0;
		virtual HRESULT __stdcall Delete() = 0;
	};
#pragma endregion

	// ================================ CodeClass ================================
#pragma region
	const GUID IID_CodeClass = { 0xB1F42514, 0x91CD, 0x4D3A, { 0x8B, 0x25, 0xA3, 0x17, 0xD8, 0x03, 0x2B, 0x24 } };

	struct __declspec(novtable) CodeClass {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeClass ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Collection(CodeElements **ppCollection) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pVal) = 0;
		virtual HRESULT __stdcall put_Name(BSTR NewName) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *pVal) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **pProjItem) = 0;
		virtual HRESULT __stdcall get_Kind(vsCMElement *pCodeEltKind) = 0;
		virtual HRESULT __stdcall get_IsCodeType(VARIANT_BOOL *pIsCodeType) = 0;
		virtual HRESULT __stdcall get_InfoLocation(vsCMInfoLocation *pInfoLocation) = 0;
		virtual HRESULT __stdcall get_Children(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall get_StartPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_EndPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall GetStartPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall GetEndPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **ParentObject) = 0;
		virtual HRESULT __stdcall get_Namespace(CodeNamespace **ppCodeNamespace) = 0;
		virtual HRESULT __stdcall get_Bases(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Members(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall put_Access(vsCMAccess Access) = 0;
		virtual HRESULT __stdcall get_Access(vsCMAccess *pAccess) = 0;
		virtual HRESULT __stdcall get_Attributes(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_DocComment(BSTR *pDocComment) = 0;
		virtual HRESULT __stdcall put_DocComment(BSTR DocComment) = 0;
		virtual HRESULT __stdcall get_Comment(BSTR *pComment) = 0;
		virtual HRESULT __stdcall put_Comment(BSTR Comment) = 0;
		virtual HRESULT __stdcall AddBase(VARIANT Base, VARIANT Position, CodeElement **ppOut) = 0;
		virtual HRESULT __stdcall AddAttribute(BSTR Name, BSTR Value, VARIANT Position, CodeAttribute **ppCodeAttribute) = 0;
		virtual HRESULT __stdcall RemoveBase(VARIANT Element) = 0;
		virtual HRESULT __stdcall RemoveMember(VARIANT Element) = 0;
		virtual HRESULT __stdcall get_IsDerivedFrom(BSTR FullName, VARIANT_BOOL *pVal) = 0;
		virtual HRESULT __stdcall get_DerivedTypes(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_ImplementedInterfaces(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_IsAbstract(VARIANT_BOOL *pIsAbstract) = 0;
		virtual HRESULT __stdcall put_IsAbstract(VARIANT_BOOL Abstract) = 0;
		virtual HRESULT __stdcall AddImplementedInterface(VARIANT Base, VARIANT Position, CodeInterface **ppCodeInterface) = 0;
		virtual HRESULT __stdcall AddFunction(BSTR Name, vsCMFunction Kind, VARIANT Type, VARIANT Position, vsCMAccess Access, VARIANT Location, CodeFunction **ppCodeFunction) = 0;
		virtual HRESULT __stdcall AddVariable(BSTR Name, VARIANT Type, VARIANT Position, vsCMAccess Access, VARIANT Location, CodeVariable **ppCodeVariable) = 0;
		virtual HRESULT __stdcall AddProperty(BSTR GetterName, BSTR PutterName, VARIANT Type, VARIANT Position, vsCMAccess Access, VARIANT Location, CodeProperty **ppCodeProperty) = 0;
		virtual HRESULT __stdcall AddClass(BSTR Name, VARIANT Position, VARIANT Bases, VARIANT ImplementedInterfaces, vsCMAccess Access, CodeClass **ppCodeClass) = 0;
		virtual HRESULT __stdcall AddStruct(BSTR Name, VARIANT Position, VARIANT Bases, VARIANT ImplementedInterfaces, vsCMAccess Access, CodeStruct **ppCodeStruct) = 0;
		virtual HRESULT __stdcall AddEnum(BSTR Name, VARIANT Position, VARIANT Bases, vsCMAccess Access, CodeEnum **ppCodeEnum) = 0;
		virtual HRESULT __stdcall AddDelegate(BSTR Name, VARIANT Type, VARIANT Position, vsCMAccess Access, CodeDelegate **ppCodeDelegate) = 0;
		virtual HRESULT __stdcall RemoveInterface(VARIANT Element) = 0;
	};
#pragma endregion

	// ================================ CodeNamespace ================================
#pragma region
	const GUID IID_CodeNamespace = { 0x0CFBC2B8, 0x0D4E, 0x11D3, { 0x89, 0x97, 0x00, 0xC0, 0x4F, 0x68, 0x8D, 0xDE } };

	struct __declspec(novtable) CodeNamespace {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeNamespace ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Collection(CodeElements **ppCollection) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pVal) = 0;
		virtual HRESULT __stdcall put_Name(BSTR NewName) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *pVal) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **pProjItem) = 0;
		virtual HRESULT __stdcall get_Kind(vsCMElement *pCodeEltKind) = 0;
		virtual HRESULT __stdcall get_IsCodeType(VARIANT_BOOL *pIsCodeType) = 0;
		virtual HRESULT __stdcall get_InfoLocation(vsCMInfoLocation *pInfoLocation) = 0;
		virtual HRESULT __stdcall get_Children(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall get_StartPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_EndPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall GetStartPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall GetEndPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **ParentObject) = 0;
		virtual HRESULT __stdcall get_Members(CodeElements **ppMembers) = 0;
		virtual HRESULT __stdcall get_DocComment(BSTR *pDocComment) = 0;
		virtual HRESULT __stdcall put_DocComment(BSTR DocComment) = 0;
		virtual HRESULT __stdcall get_Comment(BSTR *pComment) = 0;
		virtual HRESULT __stdcall put_Comment(BSTR Comment) = 0;
		virtual HRESULT __stdcall AddNamespace(BSTR Name, VARIANT Position, CodeNamespace **ppCodeNamespace) = 0;
		virtual HRESULT __stdcall AddClass(BSTR Name, VARIANT Position, VARIANT Bases, VARIANT ImplementedInterfaces, vsCMAccess Access, CodeClass **ppCodeClass) = 0;
		virtual HRESULT __stdcall AddInterface(BSTR Name, VARIANT Position, VARIANT Bases, vsCMAccess Access, CodeInterface **ppCodeInterface) = 0;
		virtual HRESULT __stdcall AddStruct(BSTR Name, VARIANT Position, VARIANT Bases, VARIANT ImplementedInterfaces, vsCMAccess Access, CodeStruct **ppCodeStruct) = 0;
		virtual HRESULT __stdcall AddEnum(BSTR Name, VARIANT Position, VARIANT Bases, vsCMAccess Access, CodeEnum **ppCodeEnum) = 0;
		virtual HRESULT __stdcall AddDelegate(BSTR Name, VARIANT Type, VARIANT Position, vsCMAccess Access, CodeDelegate **ppCodeDelegate) = 0;
		virtual HRESULT __stdcall Remove(VARIANT Element) = 0;
	};
#pragma endregion

	// ================================ TextWindow ================================
#pragma region
	const GUID IID_TextWindow = { 0x2fc54dc9, 0x922b, 0x44eb, { 0x8c, 0xc0, 0xba, 0x18, 0x25, 0x84, 0xdc, 0x4b } };

	struct __declspec(novtable) TextWindow {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == TextWindow ==
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Parent(Window **ppParent) = 0;
		virtual HRESULT __stdcall get_Selection(TextSelection **ppSel) = 0;
		virtual HRESULT __stdcall get_ActivePane(TextPane **ppPane) = 0;
		virtual HRESULT __stdcall get_Panes(TextPanes **ppPanes) = 0;
	};
#pragma endregion

	// ================================ TextPanes ================================
#pragma region
	const GUID IID_TextPanes = { 0xd9013d31, 0x3652, 0x46b2, { 0xa2, 0x5a, 0x29, 0xa8, 0x81, 0xb9, 0xf8, 0x6b } };

	struct __declspec(novtable) TextPanes {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == TextPanes ==
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Parent(TextWindow **ppParent) = 0;
		virtual HRESULT __stdcall get_Count(long *pCount) = 0;
		virtual HRESULT __stdcall Item(VARIANT Index, TextPane **ppPane) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **lppiuReturn) = 0;
	};
#pragma endregion

	// ================================ TextPane ================================
#pragma region
	const GUID IID_TextPane = { 0x0a3bf283, 0x05f8, 0x4669, { 0x9b, 0xcb, 0xa8, 0x4b, 0x64, 0x23, 0x34, 0x9a } };

	enum vsPaneShowHow {
		vsPaneShowCentered = 0,
		vsPaneShowTop = 1,
		vsPaneShowAsIs = 2
	};

	struct __declspec(novtable) TextPane {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == TextPane ==
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Collection(TextPanes **ppPanes) = 0;
		virtual HRESULT __stdcall get_Window(Window **ppWin) = 0;
		virtual HRESULT __stdcall get_Height(long *pHeight) = 0;
		virtual HRESULT __stdcall get_Width(long *pWidth) = 0;
		virtual HRESULT __stdcall get_Selection(TextSelection **ppSel) = 0;
		virtual HRESULT __stdcall get_StartPoint(TextPoint **ppPoint) = 0;
		virtual HRESULT __stdcall Activate() = 0;
		virtual HRESULT __stdcall IsVisible(TextPoint *Point, VARIANT PointOrCount, VARIANT_BOOL *pbResult) = 0;
		virtual HRESULT __stdcall TryToShow(TextPoint *Point, vsPaneShowHow How, VARIANT PointOrCount, VARIANT_BOOL *pbResult) = 0;
	};
#pragma endregion

	// ================================ TextRange ================================
#pragma region
	const GUID IID_TextRange = { 0x72767524, 0xe3b3, 0x43d0, { 0xbb, 0x46, 0xbb, 0xe1, 0xd5, 0x56, 0xa9, 0xff } };

	struct __declspec(novtable) TextRange {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == TextRange ==
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Collection(TextRanges **ppParent) = 0;
		virtual HRESULT __stdcall get_StartPoint(EditPoint **ppPoint) = 0;
		virtual HRESULT __stdcall get_EndPoint(EditPoint **ppPoint) = 0;
	};
#pragma endregion

	// ================================ TextRanges ================================
#pragma region
	const GUID IID_TextRanges = { 0xb6422e9c, 0x9efd, 0x4f87, { 0xbd, 0xdc, 0xc7, 0xfd, 0x8f, 0x2f, 0xd3, 0x03 } };

	struct __declspec(novtable) TextRanges {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == TextRanges ==
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Parent(TextDocument **ppParent) = 0;
		virtual HRESULT __stdcall get_Count(long *pCount) = 0;
		virtual HRESULT __stdcall Item(VARIANT Index, TextRange **ppRange) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **lppiuReturn) = 0;
	};
#pragma endregion

	// ================================ EditPoint ================================
#pragma region
	const GUID IID_EditPoint = { 0xc1ffe800, 0x028b, 0x4475, { 0xa9, 0x07, 0x14, 0xf5, 0x1f, 0x19, 0xbb, 0x7d } };

	enum vsCaseOptions {
		vsCaseOptionsLowercase = 1,
		vsCaseOptionsUppercase = 2,
		vsCaseOptionsCapitalize = 3
	};

	enum vsWhitespaceOptions {
		vsWhitespaceOptionsHorizontal = 0,
		vsWhitespaceOptionsVertical = 1
	};

	struct __declspec(novtable) EditPoint {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == TextPoint ==
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Parent(TextDocument **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Line(long *lppaReturn) = 0;
		virtual HRESULT __stdcall get_LineCharOffset(long *lppaReturn) = 0;
		virtual HRESULT __stdcall get_AbsoluteCharOffset(long *lppaReturn) = 0;
		virtual HRESULT __stdcall get_DisplayColumn(long *lppaReturn) = 0;
		virtual HRESULT __stdcall get_AtEndOfDocument(VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall get_AtStartOfDocument(VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall get_AtEndOfLine(VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall get_AtStartOfLine(VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall get_LineLength(long *lppaReturn) = 0;
		virtual HRESULT __stdcall EqualTo(TextPoint *Point, VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall LessThan(TextPoint *Point, VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall GreaterThan(TextPoint *Point, VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall TryToShow(vsPaneShowHow How, VARIANT PointOrCount, VARIANT_BOOL *pbResult) = 0;
		virtual HRESULT __stdcall get_CodeElement(vsCMElement Scope, CodeElement **ppResult) = 0;
		virtual HRESULT __stdcall CreateEditPoint(EditPoint **lppaReturn) = 0;

		// == EditPoint ==
		virtual HRESULT __stdcall CharLeft(long Count = 1) = 0;
		virtual HRESULT __stdcall CharRight(long Count = 1) = 0;
		virtual HRESULT __stdcall EndOfLine() = 0;
		virtual HRESULT __stdcall StartOfLine() = 0;
		virtual HRESULT __stdcall EndOfDocument() = 0;
		virtual HRESULT __stdcall StartOfDocument() = 0;
		virtual HRESULT __stdcall WordLeft(long Count = 1) = 0;
		virtual HRESULT __stdcall WordRight(long Count = 1) = 0;
		virtual HRESULT __stdcall LineUp(long Count = 1) = 0;
		virtual HRESULT __stdcall LineDown(long Count = 1) = 0;
		virtual HRESULT __stdcall MoveToPoint(TextPoint *Point) = 0;
		virtual HRESULT __stdcall MoveToLineAndOffset(long Line, long Offset) = 0;
		virtual HRESULT __stdcall MoveToAbsoluteOffset(long Offset) = 0;
		virtual HRESULT __stdcall SetBookmark() = 0;
		virtual HRESULT __stdcall ClearBookmark() = 0;
		virtual HRESULT __stdcall NextBookmark(VARIANT_BOOL *pbFound) = 0;
		virtual HRESULT __stdcall PreviousBookmark(VARIANT_BOOL *pbFound) = 0;
		virtual HRESULT __stdcall PadToColumn(long Column) = 0;
		virtual HRESULT __stdcall Insert(BSTR Text) = 0;
		virtual HRESULT __stdcall InsertFromFile(BSTR File) = 0;
		virtual HRESULT __stdcall GetText(VARIANT PointOrCount, BSTR *pbstrText) = 0;
		virtual HRESULT __stdcall GetLines(long Start, long ExclusiveEnd, BSTR *pbstrText) = 0;
		virtual HRESULT __stdcall Copy(VARIANT PointOrCount, VARIANT_BOOL Append = 0) = 0;
		virtual HRESULT __stdcall Cut(VARIANT PointOrCount, VARIANT_BOOL Append = 0) = 0;
		virtual HRESULT __stdcall Delete(VARIANT PointOrCount) = 0;
		virtual HRESULT __stdcall Paste() = 0;
		virtual HRESULT __stdcall ReadOnly(VARIANT PointOrCount, VARIANT_BOOL *lfResult) = 0;
		virtual HRESULT __stdcall FindPattern(BSTR Pattern, long vsFindOptionsValue, EditPoint **EndPoint, TextRanges **Tags, VARIANT_BOOL *pbFound) = 0;
		virtual HRESULT __stdcall ReplacePattern(TextPoint *Point, BSTR Pattern, BSTR Replace, long vsFindOptionsValue, TextRanges **Tags, VARIANT_BOOL *pbFound) = 0;
		virtual HRESULT __stdcall Indent(TextPoint *Point = nullptr, long Count = 1) = 0;
		virtual HRESULT __stdcall Unindent(TextPoint *Point = nullptr, long Count = 1) = 0;
		virtual HRESULT __stdcall SmartFormat(TextPoint *Point) = 0;
		virtual HRESULT __stdcall OutlineSection(VARIANT PointOrCount) = 0;
		virtual HRESULT __stdcall ReplaceText(VARIANT PointOrCount, BSTR Text, long Flags) = 0;
		virtual HRESULT __stdcall ChangeCase(VARIANT PointOrCount, vsCaseOptions How) = 0;
		virtual HRESULT __stdcall DeleteWhitespace(vsWhitespaceOptions Direction = vsWhitespaceOptionsHorizontal) = 0;
	};
#pragma endregion

	// ================================ VirtualPoint ================================
#pragma region
	const GUID IID_VirtualPoint = { 0x42320454, 0x626c, 0x4dd0, { 0x9e, 0xcb, 0x35, 0x7c, 0x4f, 0x19, 0x66, 0xd8 } };

	struct __declspec(novtable) VirtualPoint {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == TextPoint ==
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Parent(TextDocument **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Line(long *lppaReturn) = 0;
		virtual HRESULT __stdcall get_LineCharOffset(long *lppaReturn) = 0;
		virtual HRESULT __stdcall get_AbsoluteCharOffset(long *lppaReturn) = 0;
		virtual HRESULT __stdcall get_DisplayColumn(long *lppaReturn) = 0;
		virtual HRESULT __stdcall get_AtEndOfDocument(VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall get_AtStartOfDocument(VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall get_AtEndOfLine(VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall get_AtStartOfLine(VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall get_LineLength(long *lppaReturn) = 0;
		virtual HRESULT __stdcall EqualTo(TextPoint *Point, VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall LessThan(TextPoint *Point, VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall GreaterThan(TextPoint *Point, VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall TryToShow(vsPaneShowHow How, VARIANT PointOrCount, VARIANT_BOOL *pbResult) = 0;
		virtual HRESULT __stdcall get_CodeElement(vsCMElement Scope, CodeElement **ppResult) = 0;
		virtual HRESULT __stdcall CreateEditPoint(EditPoint **lppaReturn) = 0;

		// == VirtualPoint ==
		virtual HRESULT __stdcall get_VirtualCharOffset(long *pOffset) = 0;
		virtual HRESULT __stdcall get_VirtualDisplayColumn(long *lppaReturn) = 0;
	};
#pragma endregion

	// ================================ TextSelection ================================
#pragma region
	const GUID IID_TextSelection = { 0x1fa0e135, 0x399a, 0x4d2c, { 0xa4, 0xfe, 0xd2, 0x1e, 0x24, 0x80, 0xf9, 0x21 } };

	enum vsStartOfLineOptions {
		vsStartOfLineOptionsFirstColumn = 0,
		vsStartOfLineOptionsFirstText = 1
	};

	enum vsInsertFlags {
		vsInsertFlagsCollapseToEnd = 1,
		vsInsertFlagsCollapseToStart = 2,
		vsInsertFlagsContainNewText = 4,
		vsInsertFlagsInsertAtEnd = 8,
		vsInsertFlagsInsertAtStart = 16
	};

	enum vsSelectionMode {
		vsSelectionModeStream = 0,
		vsSelectionModeBox = 1
	};

	struct __declspec(novtable) TextSelection {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == TextSelection ==
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Parent(TextDocument **ppParent) = 0;
		virtual HRESULT __stdcall get_AnchorPoint(VirtualPoint **ppPoint) = 0;
		virtual HRESULT __stdcall get_ActivePoint(VirtualPoint **ppPoint) = 0;
		virtual HRESULT __stdcall get_AnchorColumn(long *pColumn) = 0;
		virtual HRESULT __stdcall get_BottomLine(long *pLine) = 0;
		virtual HRESULT __stdcall get_BottomPoint(VirtualPoint **ppPoint) = 0;
		virtual HRESULT __stdcall get_CurrentColumn(long *pColumn) = 0;
		virtual HRESULT __stdcall get_CurrentLine(long *pLine) = 0;
		virtual HRESULT __stdcall get_IsEmpty(VARIANT_BOOL *pfEmpty) = 0;
		virtual HRESULT __stdcall get_IsActiveEndGreater(VARIANT_BOOL *pfGreater) = 0;
		virtual HRESULT __stdcall get_Text(BSTR *pText) = 0;
		virtual HRESULT __stdcall put_Text(BSTR Text) = 0;
		virtual HRESULT __stdcall get_TopLine(long *pLine) = 0;
		virtual HRESULT __stdcall get_TopPoint(VirtualPoint **ppPoint) = 0;
		virtual HRESULT __stdcall ChangeCase(vsCaseOptions How) = 0;
		virtual HRESULT __stdcall CharLeft(VARIANT_BOOL Extend = 0, long Count = 1) = 0;
		virtual HRESULT __stdcall CharRight(VARIANT_BOOL Extend = 0, long Count = 1) = 0;
		virtual HRESULT __stdcall ClearBookmark() = 0;
		virtual HRESULT __stdcall Collapse() = 0;
		virtual HRESULT __stdcall OutlineSection() = 0;
		virtual HRESULT __stdcall Copy() = 0;
		virtual HRESULT __stdcall Cut() = 0;
		virtual HRESULT __stdcall Paste() = 0;
		virtual HRESULT __stdcall Delete(long Count = 1) = 0;
		virtual HRESULT __stdcall DeleteLeft(long Count = 1) = 0;
		virtual HRESULT __stdcall DeleteWhitespace(vsWhitespaceOptions Direction = vsWhitespaceOptionsHorizontal) = 0;
		virtual HRESULT __stdcall EndOfLine(VARIANT_BOOL Extend = 0) = 0;
		virtual HRESULT __stdcall StartOfLine(vsStartOfLineOptions Where = vsStartOfLineOptionsFirstColumn, VARIANT_BOOL Extend = 0) = 0;
		virtual HRESULT __stdcall EndOfDocument(VARIANT_BOOL Extend = 0) = 0;
		virtual HRESULT __stdcall StartOfDocument(VARIANT_BOOL Extend = 0) = 0;
		virtual HRESULT __stdcall FindPattern(BSTR Pattern, long vsFindOptionsValue, TextRanges **Tags, VARIANT_BOOL *pfFound) = 0;
		virtual HRESULT __stdcall ReplacePattern(BSTR Pattern, BSTR Replace, long vsFindOptionsValue, TextRanges **Tags, VARIANT_BOOL *pfFound) = 0;
		virtual HRESULT __stdcall FindText(BSTR Pattern, long vsFindOptionsValue, VARIANT_BOOL *pfFound) = 0;
		virtual HRESULT __stdcall ReplaceText(BSTR Pattern, BSTR Replace, long vsFindOptionsValue, VARIANT_BOOL *pfFound) = 0;
		virtual HRESULT __stdcall GotoLine(long Line, VARIANT_BOOL Select = 0) = 0;
		virtual HRESULT __stdcall Indent(long Count = 1) = 0;
		virtual HRESULT __stdcall Unindent(long Count = 1) = 0;
		virtual HRESULT __stdcall Insert(BSTR Text, long vsInsertFlagsCollapseToEndValue = vsInsertFlagsCollapseToEnd) = 0;
		virtual HRESULT __stdcall InsertFromFile(BSTR File) = 0;
		virtual HRESULT __stdcall LineDown(VARIANT_BOOL Extend = 0, long Count = 1) = 0;
		virtual HRESULT __stdcall LineUp(VARIANT_BOOL Extend = 0, long Count = 1) = 0;
		virtual HRESULT __stdcall MoveToPoint(TextPoint *Point, VARIANT_BOOL Extend = 0) = 0;
		virtual HRESULT __stdcall MoveToLineAndOffset(long Line, long Offset, VARIANT_BOOL Extend = 0) = 0;
		virtual HRESULT __stdcall MoveToAbsoluteOffset(long Offset, VARIANT_BOOL Extend = 0) = 0;
		virtual HRESULT __stdcall NewLine(long Count = 1) = 0;
		virtual HRESULT __stdcall SetBookmark() = 0;
		virtual HRESULT __stdcall NextBookmark(VARIANT_BOOL *pbFound) = 0;
		virtual HRESULT __stdcall PreviousBookmark(VARIANT_BOOL *pbFound) = 0;
		virtual HRESULT __stdcall PadToColumn(long Column) = 0;
		virtual HRESULT __stdcall SmartFormat() = 0;
		virtual HRESULT __stdcall SelectAll() = 0;
		virtual HRESULT __stdcall SelectLine() = 0;
		virtual HRESULT __stdcall SwapAnchor() = 0;
		virtual HRESULT __stdcall Tabify() = 0;
		virtual HRESULT __stdcall Untabify() = 0;
		virtual HRESULT __stdcall WordLeft(VARIANT_BOOL Extend = 0, long Count = 1) = 0;
		virtual HRESULT __stdcall WordRight(VARIANT_BOOL Extend = 0, long Count = 1) = 0;
		virtual HRESULT __stdcall get_TextPane(TextPane **ppPane) = 0;
		virtual HRESULT __stdcall get_Mode(vsSelectionMode *pMode) = 0;
		virtual HRESULT __stdcall put_Mode(vsSelectionMode Mode) = 0;
		virtual HRESULT __stdcall get_TextRanges(TextRanges **ppRanges) = 0;
		virtual HRESULT __stdcall Backspace(long Count = 1) = 0;
		virtual HRESULT __stdcall Cancel() = 0;
		virtual HRESULT __stdcall DestructiveInsert(BSTR Text) = 0;
		virtual HRESULT __stdcall MoveTo(long Line, long Column, VARIANT_BOOL Extend = 0) = 0;
		virtual HRESULT __stdcall MoveToDisplayColumn(long Line, long Column, VARIANT_BOOL Extend = 0) = 0;
		virtual HRESULT __stdcall PageUp(VARIANT_BOOL Extend = 0, long Count = 1) = 0;
		virtual HRESULT __stdcall PageDown(VARIANT_BOOL Extend = 0, long Count = 1) = 0;
	};
#pragma endregion

	// ================================ Documents ================================
#pragma region
	const GUID IID_Documents = { 0x9e2cf3ea, 0x140f, 0x413e, { 0xbd, 0x4b, 0x7d, 0x46, 0x74, 0x0c, 0xd2, 0xf4 } };

	enum vsSaveChanges {
		vsSaveChangesYes = 1,
		vsSaveChangesNo = 2,
		vsSaveChangesPrompt = 3
	};

	struct __declspec(novtable) Documents {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Documents ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Parent(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **Enumerator) = 0;
		virtual HRESULT __stdcall Item(VARIANT index, Document **DocumentObject) = 0;
		virtual HRESULT __stdcall get_Count(long *CountOfDocuments) = 0;
		virtual HRESULT __stdcall Add(BSTR Kind, Document **ppDocument) = 0;
		virtual HRESULT __stdcall CloseAll(vsSaveChanges Save = vsSaveChangesPrompt) = 0;
		virtual HRESULT __stdcall SaveAll() = 0;
		virtual HRESULT __stdcall Open(BSTR PathName, BSTR Kind, VARIANT_BOOL ReadOnly, Document **ppDocument) = 0;
	};
#pragma endregion

	// ================================ Document ================================
#pragma region
	const GUID IID_Document = { 0x63eb5c39, 0xca8f, 0x498e, { 0x9a, 0x66, 0x6d, 0xd4, 0xa2, 0x7a, 0xc9, 0x5b } };

	enum vsSaveStatus {
		vsSaveCancelled = 2,
		vsSaveSucceeded = 1
	};

	struct __declspec(novtable) Document {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Document ==
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Kind(BSTR *pKind) = 0;
		virtual HRESULT __stdcall get_Collection(Documents **DocumentsObject) = 0;
		virtual HRESULT __stdcall get_ActiveWindow(Window **pWindow) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall get_Path(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall get_ReadOnly(VARIANT_BOOL *pRetval) = 0;
		virtual HRESULT __stdcall get_Saved(VARIANT_BOOL *pRetval) = 0;
		virtual HRESULT __stdcall put_Saved(VARIANT_BOOL bSaved) = 0;
		virtual HRESULT __stdcall get_Windows(Windows **pWindows) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **pRetVal) = 0;
		virtual HRESULT __stdcall Activate() = 0;
		virtual HRESULT __stdcall Close(vsSaveChanges Save = vsSaveChangesPrompt) = 0;
		virtual HRESULT __stdcall NewWindow(Window **pWindow) = 0;
		virtual HRESULT __stdcall Redo(VARIANT_BOOL *pbRetVal) = 0;
		virtual HRESULT __stdcall Undo(VARIANT_BOOL *pbRetVal) = 0;
		virtual HRESULT __stdcall Save(BSTR FileName, vsSaveStatus *pStatus) = 0;
		virtual HRESULT __stdcall get_Selection(IDispatch **SelectionObject) = 0;
		virtual HRESULT __stdcall Object(BSTR ModelKind, IDispatch **DataModelObject) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall PrintOut() = 0;
		virtual HRESULT __stdcall get_IndentSize(long *pSize) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall put_Language(BSTR Language) = 0;
		virtual HRESULT __stdcall put_ReadOnly(VARIANT_BOOL bReadOnly) = 0;
		virtual HRESULT __stdcall get_TabSize(long *pSize) = 0;
		virtual HRESULT __stdcall ClearBookmarks() = 0;
		virtual HRESULT __stdcall MarkText(BSTR Pattern, long Flags, VARIANT_BOOL *pbRetVal) = 0;
		virtual HRESULT __stdcall ReplaceText(BSTR FindText, BSTR ReplaceText, long Flags, VARIANT_BOOL *pbRetVal) = 0;
		virtual HRESULT __stdcall get_Type(BSTR *pType) = 0;
	};
#pragma endregion

	// ================================ TextDocument ================================
#pragma region
	const GUID IID_TextDocument = { 0xcb218890, 0x1382, 0x472b, { 0x91, 0x18, 0x78, 0x27, 0x00, 0xc8, 0x81, 0x15 } };

	struct __declspec(novtable) TextDocument {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == TextDocument ==
		virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
		virtual HRESULT __stdcall get_Parent(Document **ppParent) = 0;
		virtual HRESULT __stdcall get_Selection(TextSelection **ppSel) = 0;
		virtual HRESULT __stdcall ClearBookmarks() = 0;
		virtual HRESULT __stdcall MarkText(BSTR Pattern, long vsFindOptionsValue, VARIANT_BOOL *pbRetVal) = 0;
		virtual HRESULT __stdcall ReplacePattern(BSTR Pattern, BSTR Replace, long vsFindOptionsValue, TextRanges **Tags, VARIANT_BOOL *pbRetVal) = 0;
		virtual HRESULT __stdcall CreateEditPoint(TextPoint *TextPoint, EditPoint **lppaReturn) = 0;
		virtual HRESULT __stdcall get_StartPoint(TextPoint **pStartPoint) = 0;
		virtual HRESULT __stdcall get_EndPoint(TextPoint **pEndPoint) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall put_Language(BSTR Language) = 0;
		virtual HRESULT __stdcall get_Type(BSTR *pType) = 0;
		virtual HRESULT __stdcall get_IndentSize(long *pSize) = 0;
		virtual HRESULT __stdcall get_TabSize(long *pSize) = 0;
		virtual HRESULT __stdcall ReplaceText(BSTR FindText, BSTR ReplaceText, long vsFindOptionsValue, VARIANT_BOOL *pbRetVal) = 0;
		virtual HRESULT __stdcall PrintOut() = 0;
	};
#pragma endregion

	// ================================ TextPoint ================================
#pragma region
	const GUID IID_TextPoint = { 0x7f59e94e, 0x4939, 0x40d2, { 0x9f, 0x7f, 0xb7, 0x65, 0x1c, 0x25, 0x90, 0x5d } };

	struct __declspec(novtable) TextPoint {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == TextPoint ==
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Parent(TextDocument **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Line(long *lppaReturn) = 0;
		virtual HRESULT __stdcall get_LineCharOffset(long *lppaReturn) = 0;
		virtual HRESULT __stdcall get_AbsoluteCharOffset(long *lppaReturn) = 0;
		virtual HRESULT __stdcall get_DisplayColumn(long *lppaReturn) = 0;
		virtual HRESULT __stdcall get_AtEndOfDocument(VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall get_AtStartOfDocument(VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall get_AtEndOfLine(VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall get_AtStartOfLine(VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall get_LineLength(long *lppaReturn) = 0;
		virtual HRESULT __stdcall EqualTo(TextPoint *Point, VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall LessThan(TextPoint *Point, VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall GreaterThan(TextPoint *Point, VARIANT_BOOL *lppaReturn) = 0;
		virtual HRESULT __stdcall TryToShow(vsPaneShowHow How, VARIANT PointOrCount, VARIANT_BOOL *pbResult) = 0;
		virtual HRESULT __stdcall get_CodeElement(vsCMElement Scope, CodeElement **ppResult) = 0;
		virtual HRESULT __stdcall CreateEditPoint(EditPoint **lppaReturn) = 0;
	};
#pragma endregion

	// ================================ CodeElement ================================
#pragma region
	const GUID IID_CodeElement = { 0x0cfbc2b6, 0x0d4e, 0x11d3, { 0x89, 0x97, 0x00, 0xc0, 0x4f, 0x68, 0x8d, 0xde } };

	struct __declspec(novtable) CodeElement {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeElement ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Collection(CodeElements **ppCollection) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pVal) = 0;
		virtual HRESULT __stdcall put_Name(BSTR NewName) = 0;
		virtual HRESULT __stdcall get_FullName(BSTR *pVal) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **pProjItem) = 0;
		virtual HRESULT __stdcall get_Kind(vsCMElement *pCodeEltKind) = 0;
		virtual HRESULT __stdcall get_IsCodeType(VARIANT_BOOL *pIsCodeType) = 0;
		virtual HRESULT __stdcall get_InfoLocation(vsCMInfoLocation *pInfoLocation) = 0;
		virtual HRESULT __stdcall get_Children(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall get_StartPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_EndPoint(TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall GetStartPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
		virtual HRESULT __stdcall GetEndPoint(vsCMPart Part, TextPoint **ppTextPoint) = 0;
	};
#pragma endregion

	// ================================ CodeElements ================================
#pragma region
	const GUID IID_CodeElements = { 0x0cfbc2b5, 0x0d4e, 0x11d3, { 0x89, 0x97, 0x00, 0xc0, 0x4f, 0x68, 0x8d, 0xde } };

	struct __declspec(novtable) CodeElements {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == CodeElements ==
		virtual HRESULT __stdcall _NewEnum(IUnknown **lppiuReturn) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **ParentObject) = 0;
		virtual HRESULT __stdcall Item(VARIANT index, CodeElement **ppCodeElement) = 0;
		virtual HRESULT __stdcall get_Count(long *CountOfCodeElements) = 0;
		virtual HRESULT __stdcall Reserved1(VARIANT Element) = 0;
		virtual HRESULT __stdcall CreateUniqueID(BSTR Prefix, BSTR *NewName, VARIANT_BOOL *pRootUnique) = 0;
	};
#pragma endregion

	// ================================ FileCodeModel ================================
#pragma region
	const GUID IID_FileCodeModel = { 0xed1a3f99, 0x4477, 0x11d3, { 0x89, 0xbf, 0x00, 0xc0, 0x4f, 0x68, 0x8d, 0xde } };

	struct __declspec(novtable) FileCodeModel {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == FileCodeModel ==
		virtual HRESULT __stdcall get_DTE(DTE **DTEObject) = 0;
		virtual HRESULT __stdcall get_Parent(ProjectItem **pProjItem) = 0;
		virtual HRESULT __stdcall get_Language(BSTR *pLanguage) = 0;
		virtual HRESULT __stdcall get_CodeElements(CodeElements **ppCodeElements) = 0;
		virtual HRESULT __stdcall CodeElementFromPoint(TextPoint *Point, vsCMElement Scope, CodeElement **ppCodeElement) = 0;
		virtual HRESULT __stdcall AddNamespace(BSTR Name, VARIANT Position, CodeNamespace **ppCodeNamespace) = 0;
		virtual HRESULT __stdcall AddClass(BSTR Name, VARIANT Position, VARIANT Bases, VARIANT ImplementedInterfaces, vsCMAccess Access, CodeClass **ppCodeClass) = 0;
		virtual HRESULT __stdcall AddInterface(BSTR Name, VARIANT Position, VARIANT Bases, vsCMAccess Access, CodeInterface **ppCodeInterface) = 0;
		virtual HRESULT __stdcall AddFunction(BSTR Name, vsCMFunction Kind, VARIANT Type, VARIANT Position, vsCMAccess Access, CodeFunction **ppCodeFunction) = 0;
		virtual HRESULT __stdcall AddVariable(BSTR Name, VARIANT Type, VARIANT Position, vsCMAccess Access, CodeVariable **ppCodeVariable) = 0;
		virtual HRESULT __stdcall AddAttribute(BSTR Name, BSTR Value, VARIANT Position, CodeAttribute **ppCodeAttribute) = 0;
		virtual HRESULT __stdcall AddStruct(BSTR Name, VARIANT Position, VARIANT Bases, VARIANT ImplementedInterfaces, vsCMAccess Access, CodeStruct **ppCodeStruct) = 0;
		virtual HRESULT __stdcall AddEnum(BSTR Name, VARIANT Position, VARIANT Bases, vsCMAccess Access, CodeEnum **ppCodeEnum) = 0;
		virtual HRESULT __stdcall AddDelegate(BSTR Name, VARIANT Type, VARIANT Position, vsCMAccess Access, CodeDelegate **ppCodeDelegate) = 0;
		virtual HRESULT __stdcall Remove(VARIANT Element) = 0;
	};
#pragma endregion

	// ================================ ProjectItems ================================
#pragma region
	const GUID IID_ProjectItems = { 0x8e2f1269, 0x185e, 0x43c7, { 0x88, 0x99, 0x95, 0x0a, 0xd2, 0x76, 0x9c, 0xcf } };

	struct __declspec(novtable) ProjectItems {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == ProjectItems ==
		virtual HRESULT __stdcall Item(VARIANT index, ProjectItem **lppcReturn) = 0;
		virtual HRESULT __stdcall get_Parent(IDispatch **lppptReturn) = 0;
		virtual HRESULT __stdcall get_Count(long *lplReturn) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **lppiuReturn) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Kind(BSTR *lpbstrFileName) = 0;
		virtual HRESULT __stdcall AddFromFile(BSTR FileName, ProjectItem **lppcReturn) = 0;
		virtual HRESULT __stdcall AddFromTemplate(BSTR FileName, BSTR Name, ProjectItem **lppcReturn) = 0;
		virtual HRESULT __stdcall AddFromDirectory(BSTR Directory, ProjectItem **lppcReturn) = 0;
		virtual HRESULT __stdcall get_ContainingProject(Project **ppProject) = 0;
		virtual HRESULT __stdcall AddFolder(BSTR Name, BSTR Kind, ProjectItem **pProjectItem) = 0;
		virtual HRESULT __stdcall AddFromFileCopy(BSTR FilePath, ProjectItem **pProjectItem) = 0;
	};
#pragma endregion

	// ================================ ProjectItem ================================
#pragma region
	const GUID IID_ProjectItem = { 0x0b48100a, 0x473e, 0x433c, { 0xab, 0x8f, 0x66, 0xb9, 0x73, 0x9a, 0xb6, 0x20 } };

	struct __declspec(novtable) ProjectItem {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == ProjectItem ==
		virtual HRESULT __stdcall get_IsDirty(VARIANT_BOOL *lpfReturn) = 0;
		virtual HRESULT __stdcall put_IsDirty(VARIANT_BOOL DirtyFlag) = 0;
		virtual HRESULT __stdcall get_FileNames(short Index, BSTR *lpbstrReturn) = 0;
		virtual HRESULT __stdcall SaveAs(BSTR NewFileName, VARIANT_BOOL *lpfReturn) = 0;
		virtual HRESULT __stdcall get_FileCount(short *lpsReturn) = 0;
		virtual HRESULT __stdcall get_Name(BSTR *pbstrReturn) = 0;
		virtual HRESULT __stdcall put_Name(BSTR bstrName) = 0;
		virtual HRESULT __stdcall get_Collection(ProjectItems **lppcReturn) = 0;
		virtual HRESULT __stdcall get_Properties(Properties **ppObject) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Kind(BSTR *lpbstrFileName) = 0;
		virtual HRESULT __stdcall get_ProjectItems(ProjectItems **lppcReturn) = 0;
		virtual HRESULT __stdcall get_IsOpen(BSTR ViewKind, VARIANT_BOOL *lpfReturn) = 0;
		virtual HRESULT __stdcall Open(BSTR ViewKind, Window **lppfReturn) = 0;
		virtual HRESULT __stdcall Remove() = 0;
		virtual HRESULT __stdcall ExpandView() = 0;
		virtual HRESULT __stdcall get_Object(IDispatch **ProjectItemModel) = 0;
		virtual HRESULT __stdcall get_Extender(BSTR ExtenderName, IDispatch **Extender) = 0;
		virtual HRESULT __stdcall get_ExtenderNames(VARIANT *ExtenderNames) = 0;
		virtual HRESULT __stdcall get_ExtenderCATID(BSTR *pRetval) = 0;
		virtual HRESULT __stdcall get_Saved(VARIANT_BOOL *lpfReturn) = 0;
		virtual HRESULT __stdcall put_Saved(VARIANT_BOOL SavedFlag) = 0;
		virtual HRESULT __stdcall get_ConfigurationManager(ConfigurationManager **ppConfigurationManager) = 0;
		virtual HRESULT __stdcall get_FileCodeModel(FileCodeModel **ppFileCodeModel) = 0;
		virtual HRESULT __stdcall Save(BSTR FileName = L"") = 0;
		virtual HRESULT __stdcall get_Document(Document **ppDocument) = 0;
		virtual HRESULT __stdcall get_SubProject(Project **ppProject) = 0;
		virtual HRESULT __stdcall get_ContainingProject(Project **ppProject) = 0;
		virtual HRESULT __stdcall Delete() = 0;
	};
#pragma endregion

	// ================================ LinkedWindows ================================
#pragma region
	const GUID IID_LinkedWindows = { 0xf00ef34a, 0xa654, 0x4c1b, { 0x89, 0x7a, 0x58, 0x5d, 0x5b, 0xcb, 0xb3, 0x5a } };

	struct __declspec(novtable) LinkedWindows {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == LinkedWindows ==
		virtual HRESULT __stdcall get_Parent(Window **ppptReturn) = 0;
		virtual HRESULT __stdcall Item(VARIANT index, Window **lppcReturn) = 0;
		virtual HRESULT __stdcall get_Count(long *lplReturn) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **lppiuReturn) = 0;
		virtual HRESULT __stdcall Remove(Window *Window) = 0;
		virtual HRESULT __stdcall Add(Window *Window) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
	};
#pragma endregion

	// ================================ Window ================================
#pragma region
	const GUID IID_Window = { 0x0beab46b, 0x4c07, 0x4f94, { 0xa8, 0xd7, 0x16, 0x26, 0x02, 0x0e, 0x4e, 0x53 } };

	enum vsWindowState {
		vsWindowStateNormal = 0,
		vsWindowStateMinimize = 1,
		vsWindowStateMaximize = 2
	};

	enum vsWindowType {
		vsWindowTypeCodeWindow = 0,
		vsWindowTypeDesigner = 1,
		vsWindowTypeBrowser = 2,
		vsWindowTypeWatch = 3,
		vsWindowTypeLocals = 4,
		vsWindowTypeImmediate = 5,
		vsWindowTypeSolutionExplorer = 6,
		vsWindowTypeProperties = 7,
		vsWindowTypeFind = 8,
		vsWindowTypeFindReplace = 9,
		vsWindowTypeToolbox = 10,
		vsWindowTypeLinkedWindowFrame = 11,
		vsWindowTypeMainWindow = 12,
		vsWindowTypePreview = 13,
		vsWindowTypeColorPalette = 14,
		vsWindowTypeToolWindow = 15,
		vsWindowTypeDocument = 16,
		vsWindowTypeOutput = 17,
		vsWindowTypeTaskList = 18,
		vsWindowTypeAutos = 19,
		vsWindowTypeCallStack = 20,
		vsWindowTypeThreads = 21,
		vsWindowTypeDocumentOutline = 22,
		vsWindowTypeRunningDocuments = 23
	};

	struct __declspec(novtable) Window {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Window ==
		virtual HRESULT __stdcall get_Collection(Windows **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Visible(VARIANT_BOOL *pfVisible) = 0;
		virtual HRESULT __stdcall put_Visible(VARIANT_BOOL fVisible) = 0;
		virtual HRESULT __stdcall get_Left(long *plLeft) = 0;
		virtual HRESULT __stdcall put_Left(long lLeft) = 0;
		virtual HRESULT __stdcall get_Top(long *plTop) = 0;
		virtual HRESULT __stdcall put_Top(long lTop_r) = 0;
		virtual HRESULT __stdcall get_Width(long *plWidth) = 0;
		virtual HRESULT __stdcall put_Width(long lWidth) = 0;
		virtual HRESULT __stdcall get_Height(long *plHeight) = 0;
		virtual HRESULT __stdcall put_Height(long lHeight_r) = 0;
		virtual HRESULT __stdcall get_WindowState(vsWindowState *plWindowState) = 0;
		virtual HRESULT __stdcall put_WindowState(vsWindowState wstWindowState) = 0;
		virtual HRESULT __stdcall SetFocus() = 0;
		virtual HRESULT __stdcall get_Type(vsWindowType *pKind) = 0;
		virtual HRESULT __stdcall SetKind(vsWindowType eKind) = 0;
		virtual HRESULT __stdcall get_LinkedWindows(LinkedWindows **ppwnsCollection) = 0;
		virtual HRESULT __stdcall get_LinkedWindowFrame(Window **ppwinFrame) = 0;
		virtual HRESULT __stdcall Detach() = 0;
		virtual HRESULT __stdcall Attach(long lWindowHandle) = 0;
		virtual HRESULT __stdcall get_HWnd(long *plWindowHandle) = 0;
		virtual HRESULT __stdcall get_Kind(BSTR *pbstrType) = 0;
		virtual HRESULT __stdcall get_ObjectKind(BSTR *pbstrTypeGUID) = 0;
		virtual HRESULT __stdcall get_Object(IDispatch **ppToolObject) = 0;
		virtual HRESULT __stdcall get_DocumentData(BSTR bstrWhichData, IDispatch **ppDataObject) = 0;
		virtual HRESULT __stdcall get_ProjectItem(ProjectItem **ppProjItem) = 0;
		virtual HRESULT __stdcall get_Project(Project **ppProj) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Document(Document **ppDocument) = 0;
		virtual HRESULT __stdcall get_Selection(IDispatch **ppDocument) = 0;
		virtual HRESULT __stdcall get_Linkable(VARIANT_BOOL *pLinkable) = 0;
		virtual HRESULT __stdcall put_Linkable(VARIANT_BOOL Linkable) = 0;
		virtual HRESULT __stdcall Activate() = 0;
		virtual HRESULT __stdcall Close(vsSaveChanges SaveChanges = vsSaveChangesNo) = 0;
		virtual HRESULT __stdcall get_Caption(BSTR *pbstrTitle) = 0;
		virtual HRESULT __stdcall put_Caption(BSTR pbstrTitle) = 0;
		virtual HRESULT __stdcall SetSelectionContainer(SAFEARRAY **Objects) = 0;
		virtual HRESULT __stdcall get_IsFloating(VARIANT_BOOL *Floating) = 0;
		virtual HRESULT __stdcall put_IsFloating(VARIANT_BOOL Floating) = 0;
		virtual HRESULT __stdcall get_AutoHides(VARIANT_BOOL *Hides) = 0;
		virtual HRESULT __stdcall put_AutoHides(VARIANT_BOOL Hides) = 0;
		virtual HRESULT __stdcall SetTabPicture(VARIANT Picture) = 0;
		virtual HRESULT __stdcall get_ContextAttributes(ContextAttributes **ppVal) = 0;
	};
#pragma endregion

	// ================================ Windows ================================
#pragma region
	const GUID IID_Windows = { 0x2294311A, 0xB7BC, 0x4789, { 0xB3, 0x65, 0x1C, 0x15, 0xFF, 0x2C, 0xD1, 0x7C } };

	enum vsLinkedWindowType {
		vsLinkedWindowTypeDocked = 0,
		vsLinkedWindowTypeVertical = 2,
		vsLinkedWindowTypeHorizontal = 3,
		vsLinkedWindowTypeTabbed = 1
	};

	struct __declspec(novtable) Windows {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;

		// == Windows ==
		virtual HRESULT __stdcall Item(VARIANT index, Window **lppcReturn) = 0;
		virtual HRESULT __stdcall get_Count(long *lplReturn) = 0;
		virtual HRESULT __stdcall _NewEnum(IUnknown **lppiuReturn) = 0;
		virtual HRESULT __stdcall CreateToolWindow(AddIn *AddInInst, BSTR ProgId, BSTR Caption, BSTR GuidPosition, IDispatch **DocObj, Window **lppcReturn) = 0;
		virtual HRESULT __stdcall get_DTE(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall get_Parent(DTE **lppaReturn) = 0;
		virtual HRESULT __stdcall CreateLinkedWindowFrame(Window *Window1, Window *Window2, vsLinkedWindowType Link, Window **LinkedWindowFrame) = 0;
	};
#pragma endregion

	// ================================ IDispatch ================================
#pragma region
	const GUID IID_IDispatch = { 0x00020400, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };

	struct __declspec(novtable) IDispatch {
		// == IUnknown ==
		virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
		virtual ULONG __stdcall AddRef() = 0;
		virtual ULONG __stdcall Release() = 0;

		// == IDispatch ==
		virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) = 0;
		virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
		virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) = 0;
		virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) = 0;
	};
#pragma endregion

	// ================================ _DTE ================================
#pragma region
	const GUID IID__DTE = { 0x04a72314, 0x32e9, 0x48e2,{ 0x9b, 0x87, 0xa6, 0x36, 0x03, 0x45, 0x4f, 0x3e } };

	enum vsDisplay {
		vsDisplayMDI = 1,
		vsDisplayMDITabs = 2
	};

	enum vsIDEMode {
		vsIDEModeDesign = 1,
		vsIDEModeDebug = 2
	};

	enum wizardResult {
		wizardResultSuccess = -1,
		wizardResultFailure = 0,
		wizardResultCancel = 1,
		wizardResultBackOut = 2
	};

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

	enum vsBuildState {
		vsBuildStateNotStarted = 1,
		vsBuildStateInProgress = 2,
		vsBuildStateDone = 3
	};

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

	enum vsConfigurationType {
		vsConfigurationTypeProject = 1,
		vsConfigurationTypeProjectItem = 2
	};

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