#pragma once

//
// Newer Visual Studio SDK ("IVs*" based)
//

// Windows
#include <guiddef.h>
#include <Windows.h>
// OLE
#include <DocObj.h>   

//
// Coherent, in one place, and easier to read and access documentation for IVs* interfaces. 
// 
// Sources: IDLs, genereated headers, MSDN + own experiments/reasoning.
// 
// IGNORE: SCC (Source Code Control), Retargeting, DocView, Unified Settings, VBA, Macros, Web Migration, Help, Discovery
// 
// TODO: Less confusing naming (keep the GUIDs of course). 
//

// Placeholder in the vtable for functions that are either obsolete, bugged, or just unused (as of VS 2022).
#define UNUSED_FUNC(index) virtual HRESULT __stdcall vfunc_UNUSED_##index() = 0;

#define	VSITEMID_NIL	     0xFFFFFFFF
#define	VSITEMID_ROOT      0xFFFFFFFE
#define	VSITEMID_SELECTION 0xFFFFFFFD

namespace vsix {
#pragma region Forward Declarations
  struct IUnknown;

  struct IVsBuildPropertyStorage;
  struct IVsBuildPropertyStorage2;
  struct IVsBuildPropertyStorage3;
  struct IVsHierarchy;
  struct IVsHierarchy2;
  struct IVsSolution;
  struct IVsSolution2;
  struct IVsSolution3;
  struct IVsSolution4;
  struct IVsSolution5;
  struct IVsSolution6;
  struct IVsSolution7;
  struct IVsSolution8;
  struct IVsSolutionBuildManager;
  struct IVsSolutionBuildManager2;
  struct IVsSolutionBuildManager3;
  struct IVsSolutionBuildManager4;
  struct IVsSolutionBuildManager5;
  struct IVsSolutionBuildManager6;
  struct IVsSolutionEvents;
  struct IVsSolutionEvents2;
  struct IVsSolutionEvents3;
  struct IVsSolutionEvents4;
  struct IVsSolutionEvents5;
  struct IVsSolutionEvents6;
  struct IVsSolutionEvents7;
  struct IVsSolutionEvents8;
  struct IVsUIShell;
  struct IVsUIShell2;
  struct IVsUIShell3;
  struct IVsUIShell4;
  struct IVsUIShell5;
  struct IVsUIShell6;
  struct IVsUIShell7;
  struct IVsUpdateSolutionEvents;
  struct IVsUpdateSolutionEvents2;
  struct IVsUpdateSolutionEvents3;
  struct IVsUpdateSolutionEvents4;
  struct IVsUpdateSolutionEvents5;

  struct IVsUIContextManager;
  struct IVsDropdownBar2;
  struct IVsDropdownBarClient3;
  struct IVsCompletionSet3;

  struct IVsCommandWindowCompletion;
  struct IVsImmediateStatementCompletion;
  struct IVsImmediateStatementCompletion2;

  struct IVsTrackProjectDocumentsEvents2;
  struct IVsTrackProjectDocumentsEvents3;
  struct IVsTrackProjectDocumentsEvents4;
  struct IVsTrackProjectDocuments2;
  struct IVsTrackProjectDocuments3;

  struct IVsHierarchyEvents;

  struct IVsUIContextEvents;
  struct IServiceProvider;
  struct IVsWindowFrame;
  struct IVsProject;

  struct IVsTaskBody;
  struct IVsTask;

  struct IVsPropertyBag;
  struct IVsBrowseProjectLocation;
  struct IVsUpdateSolutionEventsAsyncCallback;
  struct IVsUpdateSolutionEventsAsync;

  struct IVsCfg;
  struct IVsProjectCfg;
  struct IVsBuildableProjectCfg;
  struct IVsCfgProvider;
  struct IVsProjectCfgProvider;

  struct IVsTextLayer;
  struct IVsTextLayerMarker;
  struct IVsTextMarker;
  struct IVsTextBuffer;
  struct IVsTextLines;
  struct IVsOutputWindowPane;
  struct IVsOutput;
  struct IVsEnumOutputs;
  struct IVsBuildStatusCallback;
  struct IEnumHierarchies;
  struct IVsWindowFrameEvents;
  struct IVsTextTrackingPoint;
  struct IVsEnumLayerMarkers;
  struct IVsEnumTextSpans;
  struct IVsTextMarkerClient;
  struct IVsTextLineMarker;
  struct IVsEnumLineMarkers;
  struct IVsTextLinesEvents;
  struct IVsTextView;
  struct IVsCompletionSet;
  struct IVsTipWindow;
  struct IVsViewRangeClient;
  struct IVsTextViewFilter;
  struct IVsProjectFactory;
  struct IEnumWindowFrames;
  struct IVsUIHierarchy;
  struct IVsToolWindowToolbar;
  struct IVsBuildPropertyStorageEvents;
  struct IVsPackage;
  struct IDispatch;

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
#pragma endregion

  // ================================ IVsPackage ================================
#pragma region
  const GUID IID_IVsPackage = { 0xD4F3F4B1, 0xE900, 0x4e51, { 0xAD, 0xB3, 0xD5, 0x32, 0x34, 0x8F, 0x83, 0xCB } };

  struct VSPROPSHEETPAGE {
    DWORD dwSize;
    DWORD dwFlags;
    HINSTANCE hInstance;
    WORD wTemplateId;
    DWORD dwTemplateSize;
    BYTE *pTemplate;
    DLGPROC pfnDlgProc;
    LPARAM lParam;
    LPFNPSPCALLBACKA pfnCallback;
    UINT *pcRefParent;
    DWORD dwReserved;
    HWND hwndDlg;
  };

  enum VSPKGRESETFLAGS {
    PKGRF_TOOLBOXITEMS = 0x1,
    PKGRF_TOOLBOXSETUP = 0x2,
    PKGRF_ADDSTDPREVIEWER = 0x4
  };

  struct __declspec(novtable) IVsPackage {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsPackage ==
    virtual HRESULT __stdcall SetSite(IServiceProvider *pSP) = 0;
    virtual HRESULT __stdcall QueryClose(BOOL *pfCanClose) = 0;
    virtual HRESULT __stdcall Close() = 0;
    virtual HRESULT __stdcall GetAutomationObject(const wchar_t *prop_name, IDispatch **dispatch) = 0;
    virtual HRESULT __stdcall CreateTool(const GUID &rguidPersistenceSlot) = 0;
    virtual HRESULT __stdcall ResetDefaults(VSPKGRESETFLAGS grfFlags) = 0;
    virtual HRESULT __stdcall GetPropertyPage(const GUID &rguidPage, VSPROPSHEETPAGE *ppage) = 0;
  };
#pragma endregion

  // ================================ IVsToolWindowToolbarHost ================================
#pragma region
  const GUID IID_IVsToolWindowToolbarHost = { 0xCF7549A9, 0x7A2A, 0x4a6e, { 0xAC, 0xF4, 0x05, 0x45, 0x2C, 0x98, 0xCF, 0x7E } };

  enum VSTWT_LOCATION {
    VSTWT_LEFT = 0,
    VSTWT_TOP = (VSTWT_LEFT + 1),
    VSTWT_RIGHT = (VSTWT_TOP + 1),
    VSTWT_BOTTOM = (VSTWT_RIGHT + 1)
  };

  struct IVsToolWindowToolbarHost {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsToolWindowToolbarHost ==
    virtual HRESULT __stdcall AddToolbar(VSTWT_LOCATION dwLoc, const GUID *pguid, DWORD dwId) = 0;
    virtual HRESULT __stdcall BorderChanged() = 0;
    virtual HRESULT __stdcall ShowHideToolbar(const GUID *pguid, DWORD dwId, BOOL fShow) = 0;
    virtual HRESULT __stdcall ProcessMouseActivation(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) = 0;
    virtual HRESULT __stdcall ForceUpdateUI() = 0;
    virtual HRESULT __stdcall ProcessMouseActivationModal(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, LRESULT *plResult) = 0;
    virtual HRESULT __stdcall Close(DWORD dwReserved) = 0;
    virtual HRESULT __stdcall Show(DWORD dwReserved) = 0;
    virtual HRESULT __stdcall Hide(DWORD dwReserved) = 0;
  };
#pragma endregion

  // ================================ IVsToolWindowToolbar ================================
#pragma region
  const GUID IID_IVsToolWindowToolbar = { 0x4544D333, 0x8D5F, 0x4517, { 0x91, 0x13, 0x35, 0x50, 0xD6, 0x18, 0xF2, 0xAD } };

  struct IVsToolWindowToolbar {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsToolWindowToolbar ==
    virtual HRESULT __stdcall GetBorder(RECT *prc) = 0;
    virtual HRESULT __stdcall SetBorderSpace(const LPCBORDERWIDTHS pbw) = 0;
  };
#pragma endregion

  // ================================ IVsUIHierarchy ================================
#pragma region
  const GUID IID_IVsUIHierarchy = { 0xE82609EA, 0x5169, 0x47f4, { 0x91, 0xd0, 0x69, 0x57, 0x27, 0x2c, 0xbe, 0x9f } };

  enum VSHPROPID {
    VSHPROPID_NIL = -1,
    VSHPROPID_LAST = -1000,
    VSHPROPID_Parent = -1000,
    VSHPROPID_FirstChild = -1001,
    VSHPROPID_NextSibling = -1002,
    VSHPROPID_Root = -1003,
    VSHPROPID_TypeGuid = -1004,
    VSHPROPID_SaveName = -2002,
    VSHPROPID_Caption = -2003,
    VSHPROPID_IconImgList = -2004,
    VSHPROPID_IconIndex = -2005,
    VSHPROPID_Expandable = -2006,
    VSHPROPID_ExpandByDefault = -2011,
    VSHPROPID_ProjectName = -2012,
    VSHPROPID_Name = -2012,
    VSHPROPID_IconHandle = -2013,
    VSHPROPID_OpenFolderIconHandle = -2014,
    VSHPROPID_OpenFolderIconIndex = -2015,
    VSHPROPID_CmdUIGuid = -2016,
    VSHPROPID_SelContainer = -2017,
    VSHPROPID_BrowseObject = -2018,
    VSHPROPID_AltHierarchy = -2019,
    VSHPROPID_AltItemid = -2020,
    VSHPROPID_ProjectDir = -2021,
    VSHPROPID_SortPriority = -2022,
    VSHPROPID_UserContext = -2023,
    VSHPROPID_EditLabel = -2026,
    VSHPROPID_ExtObject = -2027,
    VSHPROPID_ExtSelectedItem = -2028,
    VSHPROPID_StateIconIndex = -2029,
    VSHPROPID_ProjectType = -2030,
    VSHPROPID_TypeName = -2030,
    VSHPROPID_ReloadableProjectFile = -2031,
    VSHPROPID_HandlesOwnReload = -2031,
    VSHPROPID_ParentHierarchy = -2032,
    VSHPROPID_ParentHierarchyItemid = -2033,
    VSHPROPID_ItemDocCookie = -2034,
    VSHPROPID_Expanded = -2035,
    VSHPROPID_ConfigurationProvider = -2036,
    VSHPROPID_ImplantHierarchy = -2037,
    VSHPROPID_OwnerKey = -2038,
    VSHPROPID_StartupServices = -2040,
    VSHPROPID_FirstVisibleChild = -2041,
    VSHPROPID_NextVisibleSibling = -2042,
    VSHPROPID_IsHiddenItem = -2043,
    VSHPROPID_IsNonMemberItem = -2044,
    VSHPROPID_IsNonLocalStorage = -2045,
    VSHPROPID_StorageType = -2046,
    VSHPROPID_ItemSubType = -2047,
    VSHPROPID_OverlayIconIndex = -2048,
    VSHPROPID_DefaultNamespace = -2049,
    VSHPROPID_IsNonSearchable = -2051,
    VSHPROPID_IsFindInFilesForegroundOnly = -2052,
    VSHPROPID_CanBuildFromMemory = -2053,
    VSHPROPID_PreferredLanguageSID = -2054,
    VSHPROPID_ShowProjInSolutionPage = -2055,
    VSHPROPID_AllowEditInRunMode = -2056,
    VSHPROPID_IsNewUnsavedItem = -2057,
    VSHPROPID_ShowOnlyItemCaption = -2058,
    VSHPROPID_ProjectIDGuid = -2059,
    VSHPROPID_DesignerVariableNaming = -2060,
    VSHPROPID_DesignerFunctionVisibility = -2061,
    VSHPROPID_HasEnumerationSideEffects = -2062,
    VSHPROPID_DefaultEnableBuildProjectCfg = -2063,
    VSHPROPID_DefaultEnableDeployProjectCfg = -2064,
    VSHPROPID_FIRST = -2064
  };

  // Represents a UI hierarchy, combining general hierarchy functionality with UI command handling.
  struct IVsUIHierarchy {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsHierarchy ==
    virtual HRESULT __stdcall SetSite(IServiceProvider *pSP) = 0;
    virtual HRESULT __stdcall GetSite(IServiceProvider **ppSP) = 0;
    virtual HRESULT __stdcall QueryClose(BOOL *pfCanClose) = 0;
    virtual HRESULT __stdcall Close() = 0;
    virtual HRESULT __stdcall GetGuidProperty(DWORD itemid, VSHPROPID propid, GUID *pguid) = 0;
    virtual HRESULT __stdcall SetGuidProperty(DWORD itemid, VSHPROPID propid, const GUID &rguid) = 0;
    virtual HRESULT __stdcall GetProperty(DWORD itemid, VSHPROPID propid, VARIANT *pvar) = 0;
    virtual HRESULT __stdcall SetProperty(DWORD itemid, VSHPROPID propid, VARIANT var) = 0;
    virtual HRESULT __stdcall GetNestedHierarchy(DWORD itemid, const IID &iidHierarchyNested, void **ppHierarchyNested, DWORD *pitemidNested) = 0;
    virtual HRESULT __stdcall GetCanonicalName(DWORD itemid, BSTR *pbstrName) = 0;
    virtual HRESULT __stdcall ParseCanonicalName(const wchar_t *pszName, DWORD *pitemid) = 0;
    virtual HRESULT __stdcall AdviseHierarchyEvents(IVsHierarchyEvents *pEventSink, DWORD *pdwCookie) = 0;
    virtual HRESULT __stdcall UnadviseHierarchyEvents(DWORD dwCookie) = 0;

    // == IVsUIHierarchy ==
    virtual HRESULT __stdcall QueryStatusCommand(DWORD itemid, const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText) = 0;
    virtual HRESULT __stdcall ExecCommand(DWORD itemid, const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut) = 0;
  };
#pragma endregion

  // ================================ IEnumWindowFrames ================================
#pragma region
  const GUID IID_IEnumWindowFrames = { 0x8C453B03, 0x8907, 0x435b, { 0x96, 0xd7, 0x57, 0x3c, 0x40, 0x94, 0x8f, 0x5c } };

  struct IEnumWindowFrames {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IEnumWindowFrames ==
    virtual HRESULT __stdcall Next(ULONG celt, IVsWindowFrame **rgelt, ULONG *pceltFetched) = 0;
    virtual HRESULT __stdcall Skip(ULONG celt) = 0;
    virtual HRESULT __stdcall Reset() = 0;
    virtual HRESULT __stdcall Clone(IEnumWindowFrames **ppenum) = 0;
  };
#pragma endregion


  // ================================ IVsProjectFactory ================================
#pragma region
  const GUID IID_IVsProjectFactory = { 0x33FCD00A, 0xBD45, 0x403c, { 0x9c, 0x66, 0x07, 0xba, 0x9a, 0x92, 0x35, 0x01 } };

  enum VSCREATEPROJFLAGS {
    CPF_CLONEFILE = 0x1,
    CPF_OPENFILE = 0x2,
    CPF_OPENDIRECTORY = 0x4,
    CPF_SILENT = 0x8,
    CPF_OVERWRITE = 0x10,
    CPF_NOTINSLNEXPLR = 0x20,
    CPF_NONLOCALSTORE = 0x40
  };

  struct IVsProjectFactory {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsProjectFactory ==
    virtual HRESULT __stdcall CanCreateProject(LPCOLESTR pszFilename, VSCREATEPROJFLAGS grfCreateFlags, BOOL *pfCanCreate) = 0;
    virtual HRESULT __stdcall CreateProject(
      LPCOLESTR pszFilename,
      LPCOLESTR pszLocation,
      LPCOLESTR pszName,
      VSCREATEPROJFLAGS grfCreateFlags,
      const IID &iidProject,
      void **ppvProject,
      BOOL *pfCanceled) = 0;
    virtual HRESULT __stdcall SetSite(IServiceProvider *pSP) = 0;
    virtual HRESULT __stdcall Close() = 0;
  };
#pragma endregion


  // ================================ IVsTextViewFilter ================================
#pragma region
  const GUID IID_IVsTextViewFilter = { 0x6B6F0B32, 0xB88B, 0x40F8, { 0xa8, 0xfe, 0x97, 0x43, 0x8c, 0x5b, 0xdb, 0xef } };

  struct TextSpan {
    long iStartIndex;
    long iStartLine;
    long iEndIndex;
    long iEndLine;
  };

  struct IVsTextViewFilter {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTextViewFilter ==
    virtual HRESULT __stdcall GetWordExtent(long iLine, long iIndex, DWORD dwFlags, TextSpan *pSpan) = 0;
    virtual HRESULT __stdcall GetDataTipText(TextSpan *pSpan, BSTR *pbstrText) = 0;
    virtual HRESULT __stdcall GetPairExtents(long iLine, long iIndex, TextSpan *pSpan) = 0;
  };
#pragma endregion


  // ================================ IVsViewRangeClient ================================
#pragma region
  const GUID IID_IVsViewRangeClient = { 0x30491A5B, 0xA47E, 0x4C9C, { 0x82, 0x04, 0x18, 0x58, 0x66, 0x48, 0xA2, 0x77 } };

  enum TextViewAction {
    TVA_SETCARETPOS = 0,
    TVA_CENTERLINES = (TVA_SETCARETPOS + 1)
  };

  struct IVsViewRangeClient {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsViewRangeClient ==
    virtual HRESULT __stdcall AdjustViewRange(IVsTextView *pView, TextViewAction action, long iLine, long iCount) = 0;
  };
#pragma endregion


  // ================================ IVsTipWindow ================================
#pragma region
  const GUID IID_IVsTipWindow = { 0x64C7FFC4, 0xB9EE, 0x4599, { 0xb1, 0x30, 0xff, 0x9d, 0x89, 0x0e, 0xcf, 0xd4 } };

  enum TipPosPreference {
    TPP_ABOVE = 0,
    TPP_BELOW = (TPP_ABOVE + 1),
    TPP_LEFT = (TPP_BELOW + 1),
    TPP_RIGHT = (TPP_LEFT + 1),
    TPP_DOCKED = (TPP_RIGHT + 1)
  };

  struct TIPSIZEDATA {
    SIZE size;
    TipPosPreference dwPosition;
  };

  struct IVsTipWindow {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTipWindow ==
    virtual HRESULT __stdcall GetContextStream(long *piPos, long *piLength) = 0;
    virtual HRESULT __stdcall GetSizePreferences(const RECT *prcCtxBounds, TIPSIZEDATA *pSizeData) = 0;
    virtual HRESULT __stdcall Paint(HDC hdc, const RECT *prc) = 0;
    virtual void __stdcall Dismiss() = 0;
    virtual LRESULT __stdcall WndProc(HWND hwnd, UINT iMsg, WPARAM wParam, LPARAM lParam) = 0;
  };
#pragma endregion

  // ================================ IVsCompletionSet ================================
#pragma region
  const GUID IID_IVsCompletionSet = { 0x0EF79249, 0xB0BF, 0x4CD0, { 0xa9, 0x66, 0xc4, 0x71, 0x35, 0x46, 0xc3, 0xa5 } };

  struct IVsCompletionSet {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsCompletionSet ==
    virtual DWORD __stdcall GetFlags() = 0;
    virtual long __stdcall GetCount() = 0;
    virtual HRESULT __stdcall GetDisplayText(long iIndex, const WCHAR **ppszText, long *piGlyph) = 0;
    virtual HRESULT __stdcall GetImageList(HANDLE *phImages) = 0;
    virtual HRESULT __stdcall GetDescriptionText(long iIndex, BSTR *pbstrDescription) = 0;
    virtual HRESULT __stdcall GetInitialExtent(long *piLine, long *piStartCol, long *piEndCol) = 0;
    virtual HRESULT __stdcall GetBestMatch(const WCHAR *pszSoFar, long iLength, long *piIndex, DWORD *pdwFlags) = 0;
    virtual HRESULT __stdcall OnCommit(const WCHAR *pszSoFar, long iIndex, BOOL fSelected, WCHAR cCommit, BSTR *pbstrCompleteWord) = 0;
    virtual void __stdcall Dismiss() = 0;
  };
#pragma endregion

  // ================================ IVsTextView ================================
#pragma region
  const GUID IID_IVsTextView = { 0xBB23A14B, 0x7C61, 0x469A, { 0x98, 0x90, 0xa9, 0x56, 0x48, 0xce, 0xd5, 0xe6 } };

  enum vsIndentStyle {
    vsIndentStyleNone = 0,
    vsIndentStyleDefault = (vsIndentStyleNone + 1),
    vsIndentStyleSmart = (vsIndentStyleDefault + 1)
  };

  struct INITVIEW {
    unsigned int fVirtualSpace;
    unsigned int fStreamSelMode;
    unsigned int fOvertype;
    unsigned int fVisibleWhitespace;
    unsigned int fWidgetMargin;
    unsigned int fSelectionMargin;
    unsigned int fDragDropMove;
    unsigned int fHotURLs;
    vsIndentStyle IndentStyle;
  };

  enum TextSelMode {
    SM_STREAM = 0,
    SM_BOX = (SM_STREAM + 1)
  };

  struct IVsTextView {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTextView ==
    virtual HRESULT __stdcall Initialize(IVsTextLines *pBuffer, HWND hwndParent, DWORD InitFlags, const INITVIEW *pInitView) = 0;
    virtual HRESULT __stdcall CloseView() = 0;
    virtual HRESULT __stdcall GetCaretPos(long *piLine, long *piColumn) = 0;
    virtual HRESULT __stdcall SetCaretPos(long iLine, long iColumn) = 0;
    virtual HRESULT __stdcall GetSelectionSpan(TextSpan *pSpan) = 0;
    virtual HRESULT __stdcall GetSelection(long *piAnchorLine, long *piAnchorCol, long *piEndLine, long *piEndCol) = 0;
    virtual HRESULT __stdcall SetSelection(long iAnchorLine, long iAnchorCol, long iEndLine, long iEndCol) = 0;
    virtual TextSelMode __stdcall GetSelectionMode() = 0;
    virtual HRESULT __stdcall SetSelectionMode(TextSelMode iSelMode) = 0;
    virtual HRESULT __stdcall ClearSelection(BOOL fMoveToAnchor) = 0;
    virtual HRESULT __stdcall CenterLines(long iTopLine, long iCount) = 0;
    virtual HRESULT __stdcall GetSelectedText(BSTR *pbstrText) = 0;
    virtual HRESULT __stdcall GetSelectionDataObject(IDataObject **ppIDataObject) = 0;
    virtual HRESULT __stdcall GetTextStream(long iTopLine, long iTopCol, long iBottomLine, long iBottomCol, BSTR *pbstrText) = 0;
    virtual HRESULT __stdcall GetBuffer(IVsTextLines **ppBuffer) = 0;
    virtual HRESULT __stdcall SetBuffer(IVsTextLines *pBuffer) = 0;
    virtual HWND __stdcall GetWindowHandle() = 0;
    virtual HRESULT __stdcall GetScrollInfo(long iBar, long *piMinUnit, long *piMaxUnit, long *piVisibleUnits, long *piFirstVisibleUnit) = 0;
    virtual HRESULT __stdcall SetScrollPosition(long iBar, long iFirstVisibleUnit) = 0;
    virtual HRESULT __stdcall AddCommandFilter(IOleCommandTarget *pNewCmdTarg, IOleCommandTarget **ppNextCmdTarg) = 0;
    virtual HRESULT __stdcall RemoveCommandFilter(IOleCommandTarget *pCmdTarg) = 0;
    virtual HRESULT __stdcall UpdateCompletionStatus(IVsCompletionSet *pCompSet, DWORD dwFlags) = 0;
    virtual HRESULT __stdcall UpdateTipWindow(IVsTipWindow *pTipWindow, DWORD dwFlags) = 0;
    virtual HRESULT __stdcall GetWordExtent(long iLine, long iCol, DWORD dwFlags, TextSpan *pSpan) = 0;
    virtual HRESULT __stdcall RestrictViewRange(long iMinLine, long iMaxLine, IVsViewRangeClient *pClient) = 0;
    virtual HRESULT __stdcall ReplaceTextOnLine(long iLine, long iStartCol, long iCharsToReplace, const WCHAR *pszNewText, long iNewLen) = 0;
    virtual HRESULT __stdcall GetLineAndColumn(long iPos, long *piLine, long *piIndex) = 0;
    virtual HRESULT __stdcall GetNearestPosition(long iLine, long iCol, long *piPos, long *piVirtualSpaces) = 0;
    virtual HRESULT __stdcall UpdateViewFrameCaption() = 0;
    virtual HRESULT __stdcall CenterColumns(long iLine, long iLeftCol, long iColCount) = 0;
    virtual HRESULT __stdcall EnsureSpanVisible(TextSpan span) = 0;
    virtual HRESULT __stdcall PositionCaretForEditing(long iLine, long cIndentLevels) = 0;
    virtual HRESULT __stdcall GetPointOfLineColumn(long iLine, long iCol, POINT *ppt) = 0;
    virtual HRESULT __stdcall GetLineHeight(long *piLineHeight) = 0;
    virtual HRESULT __stdcall HighlightMatchingBrace(DWORD dwFlags, ULONG cSpans, TextSpan *rgBaseSpans) = 0;
    virtual HRESULT __stdcall SendExplicitFocus() = 0;
    virtual HRESULT __stdcall SetTopLine(long iBaseLine) = 0;
  };
#pragma endregion

  // ================================ IVsTextLinesEvents ================================
#pragma region
  const GUID IID_IVsTextLinesEvents = { 0x598D7074, 0xDC17, 0x4162, { 0x9a, 0x2f, 0x97, 0xdd, 0x45, 0x40, 0xc2, 0xdd } };

  struct TextLineChange {
    long iStartIndex;
    long iStartLine;
    long iOldEndIndex;
    long iOldEndLine;
    long iNewEndIndex;
    long iNewEndLine;
  };

  struct IVsTextLinesEvents {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTextLinesEvents ==
    virtual void __stdcall OnChangeLineText(const TextLineChange *pTextLineChange, BOOL fLast) = 0;
    virtual void __stdcall OnChangeLineAttributes(long iFirstLine, long iLastLine) = 0;
  };
#pragma endregion

  // ================================ IVsEnumLineMarkers ================================
#pragma region
  const GUID IID_IVsEnumLineMarkers = { 0x35E981F1, 0x77EF, 0x4343, { 0xaa, 0xa1, 0x87, 0x41, 0xf3, 0x86, 0x27, 0xa2 } };

  struct IVsEnumLineMarkers {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsEnumLineMarkers ==
    virtual HRESULT __stdcall Reset() = 0;
    virtual HRESULT __stdcall Next(IVsTextLineMarker **ppRetval) = 0;
    virtual HRESULT __stdcall GetCount(long *pcMarkers) = 0;
  };
#pragma endregion

  // ================================ IVsTextLineMarker ================================
#pragma region
  const GUID IID_IVsTextLineMarker = { 0x31E2DCA7, 0xCCFF, 0x4E09, { 0xb4, 0x33, 0x17, 0xc7, 0x39, 0xcf, 0x69, 0xad } };

  struct IVsTextLineMarker {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTextMarker ==
    virtual HRESULT __stdcall GetType(long *piMarkerType) = 0;
    virtual HRESULT __stdcall SetType(long iMarkerType) = 0;
    virtual HRESULT __stdcall GetVisualStyle(DWORD *pdwFlags) = 0;
    virtual HRESULT __stdcall SetVisualStyle(DWORD dwFlags) = 0;
    virtual HRESULT __stdcall Invalidate() = 0;
    virtual HRESULT __stdcall DrawGlyph(HDC hdc, RECT *pRect) = 0;
    virtual HRESULT __stdcall GetTipText(BSTR *pbstrText) = 0;
    virtual HRESULT __stdcall UnadviseClient() = 0;
    virtual HRESULT __stdcall GetMarkerCommandInfo(long iItem, BSTR *pbstrText, DWORD *pcmdf) = 0;
    virtual HRESULT __stdcall ExecMarkerCommand(long iItem) = 0;
    virtual HRESULT __stdcall GetBehavior(DWORD *pdwBehavior) = 0;
    virtual HRESULT __stdcall SetBehavior(DWORD dwBehavior) = 0;
    virtual HRESULT __stdcall GetPriorityIndex(long *piPriorityIndex) = 0;

    // == IVsTextLineMarker == 
    virtual HRESULT __stdcall GetLineBuffer(IVsTextLines **ppBuffer) = 0;
    virtual HRESULT __stdcall ResetSpan(long iStartLine, long iStartIndex, long iEndLine, long iEndIndex) = 0;
    virtual HRESULT __stdcall GetCurrentSpan(TextSpan *pSpan) = 0;
  };
#pragma endregion

  // ================================ IVsTextMarkerClient ================================
#pragma region
  const GUID IID_IVsTextMarkerClient = { 0xB1938F1B, 0xD7A9, 0x42F8, { 0x99, 0x60, 0xd0, 0x09, 0x02, 0x7b, 0x3d, 0x2e } };

  struct IVsTextMarkerClient {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTextMarkerClient ==
    virtual void    __stdcall MarkerInvalidated() = 0;
    virtual HRESULT __stdcall GetTipText(IVsTextMarker *pMarker, BSTR *pbstrText) = 0;
    virtual void    __stdcall OnBufferSave(LPCOLESTR pszFileName) = 0;
    virtual void    __stdcall OnBeforeBufferClose() = 0;
    virtual HRESULT __stdcall GetMarkerCommandInfo(IVsTextMarker *pMarker, long iItem, BSTR *pbstrText, DWORD *pcmdf) = 0;
    virtual HRESULT __stdcall ExecMarkerCommand(IVsTextMarker *pMarker, long iItem) = 0;
    virtual void    __stdcall OnAfterSpanReload() = 0;
    virtual HRESULT __stdcall OnAfterMarkerChange(IVsTextMarker *pMarker) = 0;
  };
#pragma endregion

  // ================================ IVsEnumTextSpans ================================
#pragma region
  const GUID IID_IVsEnumTextSpans = { 0x0F343A75, 0x968B, 0x439E, { 0x81, 0xd6, 0x0d, 0x06, 0x6e, 0x53, 0xd2, 0x8d } };

  struct IVsEnumTextSpans {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsEnumTextSpans ==
    virtual HRESULT __stdcall Reset(void) = 0;
    virtual HRESULT __stdcall Next(ULONG cEl, TextSpan *ppOut, ULONG *pcElFetched) = 0;
    virtual HRESULT __stdcall GetCount(ULONG *pcSpans) = 0;
  };
#pragma endregion

  // ================================ IVsEnumLayerMarkers ================================
#pragma region
  const GUID IID_IVsEnumLayerMarkers = { 0x8F591607, 0x2A26, 0x4A9D, { 0xa6, 0xc5, 0x14, 0x7c, 0x2e, 0x51, 0xe8, 0x95 } };
  
  struct IVsEnumLayerMarkers {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsEnumLayerMarkers ==
    virtual HRESULT __stdcall Reset() = 0;
    virtual HRESULT __stdcall Next(IVsTextLayerMarker **ppRetval) = 0;
    virtual HRESULT __stdcall GetCount(long *pcMarkers) = 0;
  };
#pragma endregion

  // ================================ IVsTextTrackingPoint ================================
#pragma region
  const GUID IID_IVsTextTrackingPoint = { 0xD6BF0A8A, 0x3798, 0x49C5, { 0x88, 0x06, 0x64, 0x8a, 0x63, 0x5e, 0xac, 0xc8 } };
  
  struct IVsTextTrackingPoint {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTextTrackingPoint ==
    virtual HRESULT __stdcall GetTextLayer(IVsTextLayer **ppLayer) = 0;
    virtual HRESULT __stdcall GetCurrentLineIndex(long *piLine, long *piIndex) = 0;
    virtual HRESULT __stdcall GetBehavior(DWORD *pdwBehavior) = 0;
    virtual HRESULT __stdcall SetBehavior(DWORD dwBehavior) = 0;
  };
#pragma endregion

  // ================================ IUnknown ================================
#pragma region
  const GUID IID_IUnknown = { 0x00000000, 0x0000, 0x0000, { 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };

  struct IUnknown {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;
  };
#pragma endregion

  // ================================ IVsTextLayer ================================
#pragma region
  const GUID IID_IVsTextLayer = { 0x0E145D3F, 0xBEFC, 0x4FD9, { 0x87, 0x14, 0xb0, 0x1a, 0xe8, 0x9f, 0x43, 0x96 } };

  struct MARKERDATA {
    long iTopLine;
    long iBottomLine;
    long iCount;
    IVsTextLayerMarker **rgpMarker;
    BOOL *rgfLineMarked;
    IVsTextLayer *pLayer;
    MARKERDATA *pNext;
  };

  enum EOLTYPE {
    eolCRLF = 0,
    eolCR = (eolCRLF + 1),
    eolLF = (eolCR + 1),
    eolUNI_LINESEP = (eolLF + 1),
    eolUNI_PARASEP = (eolUNI_LINESEP + 1),
    eolEOF = (eolUNI_PARASEP + 1),
    eolNONE = (eolEOF + 1),
    MAX_EOLTYPES = (eolNONE + 1)
  };

  struct AtomicText {
    AtomicText *pNext;
    long iIndex;
    const WCHAR *pszAtomicText;
    IUnknown *punkAtom;
  };

  struct LINEDATAEX {
    long iLength;
    const WCHAR *pszText;
    EOLTYPE iEolType;
    const ULONG *pAttributes;
    USHORT dwFlags;
    USHORT dwReserved;
    AtomicText *pAtomicTextChain;
  };

  struct IVsTextLayer {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTextLayer ==
    virtual HRESULT __stdcall LocalLineIndexToBase(long iLocalLine, long iLocalIndex, long *piBaseLine, long *piBaseIndex) = 0;
    virtual HRESULT __stdcall BaseLineIndexToLocal(long iBaseLine, long iBaseIndex, long *piLocalLine, long *piLocalIndex) = 0;
    virtual HRESULT __stdcall LocalLineIndexToDeeperLayer(IVsTextLayer *pTargetLayer, long iLocalLine, long iLocalIndex, long *piTargetLine, long *piTargetIndex) = 0;
    virtual HRESULT __stdcall DeeperLayerLineIndexToLocal(DWORD dwFlags, IVsTextLayer *pTargetLayer, long iLayerLine, long iLayerIndex, long *piLocalLine, long *piLocalIndex) = 0;
    virtual HRESULT __stdcall GetBaseBuffer(IVsTextLines **ppiBuf) = 0;
    virtual HRESULT __stdcall LockBufferEx(DWORD dwFlags) = 0;
    virtual HRESULT __stdcall UnlockBufferEx(DWORD dwFlags) = 0;
    virtual HRESULT __stdcall GetLengthOfLine(long iLine, long *piLength) = 0;
    virtual HRESULT __stdcall GetLineCount(long *piLineCount) = 0;
    virtual HRESULT __stdcall GetLastLineIndex(long *piLine, long *piIndex) = 0;
    virtual HRESULT __stdcall GetMarkerData(long iTopLine, long iBottomLine, MARKERDATA *pMarkerData) = 0;
    virtual HRESULT __stdcall ReleaseMarkerData(MARKERDATA *pMarkerData) = 0;
    virtual HRESULT __stdcall GetLineDataEx(DWORD dwFlags, long iLine, long iStartIndex, long iEndIndex, LINEDATAEX *pLineData, MARKERDATA *pMarkerData) = 0;
    virtual HRESULT __stdcall ReleaseLineDataEx(LINEDATAEX *pLineData) = 0;
    virtual HRESULT __stdcall GetLineText(long iStartLine, long iStartIndex, long iEndLine, long iEndIndex, BSTR *pbstrBuf) = 0;
    virtual HRESULT __stdcall CopyLineText(long iStartLine, long iStartIndex, long iEndLine, long iEndIndex, LPWSTR pszBuf, long *pcchBuf) = 0;
    virtual HRESULT __stdcall ReplaceLines(long iStartLine, long iStartIndex, long iEndLine, long iEndIndex, LPCWSTR pszText, long iNewLen, TextSpan *pChangedSpan) = 0;
    virtual HRESULT __stdcall CanReplaceLines(long iStartLine, long iStartIndex, long iEndLine, long iEndIndex, long iNewLen) = 0;
    virtual HRESULT __stdcall CreateTrackingPoint(long iLine, long iIndex, IVsTextTrackingPoint **ppMarker) = 0;
    virtual HRESULT __stdcall EnumLayerMarkers(long iStartLine, long iStartIndex, long iEndLine, long iEndIndex, long iMarkerType, DWORD dwFlags, IVsEnumLayerMarkers **ppEnum) = 0;
    virtual HRESULT __stdcall ReplaceLinesEx(DWORD dwFlags, long iStartLine, long iStartIndex, long iEndLine, long iEndIndex, LPCWSTR pszText, long iNewLen, TextSpan *pChangedSpan) = 0;
    virtual HRESULT __stdcall MapLocalSpansToTextOriginatingLayer(DWORD dwFlags, IVsEnumTextSpans *pLocalSpanEnum, IVsTextLayer **ppTargetLayer, IVsEnumTextSpans **ppTargetSpanEnum) = 0;
  };
#pragma endregion

  // ================================ IVsTextMarker ================================
#pragma region
  const GUID IID_IVsTextMarker = { 0x950122D9, 0x1A51, 0x43CA, { 0x8c, 0xed, 0xb5, 0xd9, 0xe4, 0x2d, 0xe1, 0xb5 } };

  struct IVsTextMarker {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTextMarker ==
    virtual HRESULT __stdcall GetType(long *piMarkerType) = 0;
    virtual HRESULT __stdcall SetType(long iMarkerType) = 0;
    virtual HRESULT __stdcall GetVisualStyle(DWORD *pdwFlags) = 0;
    virtual HRESULT __stdcall SetVisualStyle(DWORD dwFlags) = 0;
    virtual HRESULT __stdcall Invalidate() = 0;
    virtual HRESULT __stdcall DrawGlyph(HDC hdc, RECT *pRect) = 0;
    virtual HRESULT __stdcall GetTipText(BSTR *pbstrText) = 0;
    virtual HRESULT __stdcall UnadviseClient() = 0;
    virtual HRESULT __stdcall GetMarkerCommandInfo(long iItem, BSTR *pbstrText, DWORD *pcmdf) = 0;
    virtual HRESULT __stdcall ExecMarkerCommand(long iItem) = 0;
    virtual HRESULT __stdcall GetBehavior(DWORD *pdwBehavior) = 0;
    virtual HRESULT __stdcall SetBehavior(DWORD dwBehavior) = 0;
    virtual HRESULT __stdcall GetPriorityIndex(long *piPriorityIndex) = 0;
  };
#pragma endregion

  // ================================ IVsTextLayerMarker ================================
#pragma region
  const GUID IID_IVsTextLayerMarker = { 0x28C149D2, 0x8FCB, 0x4AB3, { 0x85, 0x84, 0x9a, 0x27, 0x47, 0xf3, 0xf8, 0xfc } };

  struct IVsTextLayerMarker {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTextMarker ==
    virtual HRESULT __stdcall GetType(long *piMarkerType) = 0;
    virtual HRESULT __stdcall SetType(long iMarkerType) = 0;
    virtual HRESULT __stdcall GetVisualStyle(DWORD *pdwFlags) = 0;
    virtual HRESULT __stdcall SetVisualStyle(DWORD dwFlags) = 0;
    virtual HRESULT __stdcall Invalidate() = 0;
    virtual HRESULT __stdcall DrawGlyph(HDC hdc, RECT *pRect) = 0;
    virtual HRESULT __stdcall GetTipText(BSTR *pbstrText) = 0;
    virtual HRESULT __stdcall UnadviseClient() = 0;
    virtual HRESULT __stdcall GetMarkerCommandInfo(long iItem, BSTR *pbstrText, DWORD *pcmdf) = 0;
    virtual HRESULT __stdcall ExecMarkerCommand(long iItem) = 0;
    virtual HRESULT __stdcall GetBehavior(DWORD *pdwBehavior) = 0;
    virtual HRESULT __stdcall SetBehavior(DWORD dwBehavior) = 0;
    virtual HRESULT __stdcall GetPriorityIndex(long *piPriorityIndex) = 0;

    // == IVsTextLayerMarker ==
    virtual HRESULT __stdcall GetTextLayer(IVsTextLayer **ppLayer) = 0;
    virtual HRESULT __stdcall ResetSpan(long iStartLine, long iStartIndex, long iEndLine, long iEndIndex) = 0;
    virtual HRESULT __stdcall GetCurrentSpan(TextSpan *pSpan) = 0;
    virtual HRESULT __stdcall IsInvalidated() = 0;
    virtual HRESULT __stdcall QueryClientInterface(const IID &iid, void **ppClient) = 0;
    virtual HRESULT __stdcall DrawGlyphEx(DWORD dwFlags, HDC hdc, RECT *pRect, long iLineHeight) = 0;
  };
#pragma endregion

  // ================================ IVsTextBuffer ================================
#pragma region
  const GUID IID_IVsTextBuffer = { 0xC08E5275, 0x0D26, 0x4DE9, { 0x88, 0x92, 0x99, 0x40, 0x24, 0xc2, 0x37, 0x50 } };

  struct IVsTextBuffer {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTextBuffer ==
    virtual HRESULT __stdcall LockBuffer() = 0;
    virtual HRESULT __stdcall UnlockBuffer() = 0;
    virtual HRESULT __stdcall InitializeContent(const WCHAR *pszText, long iLength) = 0;
    virtual HRESULT __stdcall GetStateFlags(DWORD *pdwReadOnlyFlags) = 0;
    virtual HRESULT __stdcall SetStateFlags(DWORD dwReadOnlyFlags) = 0;
    virtual HRESULT __stdcall GetPositionOfLine(long iLine, long *piPosition) = 0;
    virtual HRESULT __stdcall GetPositionOfLineIndex(long iLine, long iIndex, long *piPosition) = 0;
    virtual HRESULT __stdcall GetLineIndexOfPosition(long iPosition, long *piLine, long *piColumn) = 0;
    virtual HRESULT __stdcall GetLengthOfLine(long iLine, long *piLength) = 0;
    virtual HRESULT __stdcall GetLineCount(long *piLineCount) = 0;
    virtual HRESULT __stdcall GetSize(long *piLength) = 0;
    virtual HRESULT __stdcall GetLanguageServiceID(GUID *pguidLangService) = 0;
    virtual HRESULT __stdcall SetLanguageServiceID(const GUID &guidLangService) = 0;
    virtual HRESULT __stdcall GetUndoManager(IOleUndoManager **ppUndoManager) = 0;
    virtual HRESULT __stdcall Reserved1() = 0;
    virtual HRESULT __stdcall Reserved2() = 0;
    virtual HRESULT __stdcall Reserved3() = 0;
    virtual HRESULT __stdcall Reserved4() = 0;
    virtual HRESULT __stdcall Reload(BOOL fUndoable) = 0;
    virtual HRESULT __stdcall LockBufferEx(DWORD dwFlags) = 0;
    virtual HRESULT __stdcall UnlockBufferEx(DWORD dwFlags) = 0;
    virtual HRESULT __stdcall GetLastLineIndex(long *piLine, long *piIndex) = 0;
    virtual HRESULT __stdcall Reserved5() = 0;
    virtual HRESULT __stdcall Reserved6() = 0;
    virtual HRESULT __stdcall Reserved7() = 0;
    virtual HRESULT __stdcall Reserved8() = 0;
    virtual HRESULT __stdcall Reserved9() = 0;
    virtual HRESULT __stdcall Reserved10() = 0;
  };
#pragma endregion

  // ================================ IVsTextLines ================================
#pragma region
  const GUID IID_IVsTextLines = { 0xECF3E19D, 0x149C, 0x43AA, { 0x80, 0xc2, 0xd0, 0xa4, 0x69, 0x46, 0xda, 0xa3 } };

  struct LINEDATA {
    long iLength;
    const WCHAR *pszText;
    const ULONG *pAttributes;
    EOLTYPE iEolType;
    BOOL fMarkers;
  };

  struct IVsTextLines {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTextBuffer ==
    virtual HRESULT __stdcall LockBuffer() = 0;
    virtual HRESULT __stdcall UnlockBuffer() = 0;
    virtual HRESULT __stdcall InitializeContent(const WCHAR *pszText, long iLength) = 0;
    virtual HRESULT __stdcall GetStateFlags(DWORD *pdwReadOnlyFlags) = 0;
    virtual HRESULT __stdcall SetStateFlags(DWORD dwReadOnlyFlags) = 0;
    virtual HRESULT __stdcall GetPositionOfLine(long iLine, long *piPosition) = 0;
    virtual HRESULT __stdcall GetPositionOfLineIndex(long iLine, long iIndex, long *piPosition) = 0;
    virtual HRESULT __stdcall GetLineIndexOfPosition(long iPosition, long *piLine, long *piColumn) = 0;
    virtual HRESULT __stdcall GetLengthOfLine(long iLine, long *piLength) = 0;
    virtual HRESULT __stdcall GetLineCount(long *piLineCount) = 0;
    virtual HRESULT __stdcall GetSize(long *piLength) = 0;
    virtual HRESULT __stdcall GetLanguageServiceID(GUID *pguidLangService) = 0;
    virtual HRESULT __stdcall SetLanguageServiceID(const GUID &guidLangService) = 0;
    virtual HRESULT __stdcall GetUndoManager(IOleUndoManager **ppUndoManager) = 0;
    virtual HRESULT __stdcall Reserved1() = 0;
    virtual HRESULT __stdcall Reserved2() = 0;
    virtual HRESULT __stdcall Reserved3() = 0;
    virtual HRESULT __stdcall Reserved4() = 0;
    virtual HRESULT __stdcall Reload(BOOL fUndoable) = 0;
    virtual HRESULT __stdcall LockBufferEx(DWORD dwFlags) = 0;
    virtual HRESULT __stdcall UnlockBufferEx(DWORD dwFlags) = 0;
    virtual HRESULT __stdcall GetLastLineIndex(long *piLine, long *piIndex) = 0;
    virtual HRESULT __stdcall Reserved5() = 0;
    virtual HRESULT __stdcall Reserved6() = 0;
    virtual HRESULT __stdcall Reserved7() = 0;
    virtual HRESULT __stdcall Reserved8() = 0;
    virtual HRESULT __stdcall Reserved9() = 0;
    virtual HRESULT __stdcall Reserved10() = 0;

    // == IVsTextLines ==
    virtual HRESULT __stdcall GetMarkerData(long iTopLine, long iBottomLine, MARKERDATA *pMarkerData) = 0;
    virtual HRESULT __stdcall ReleaseMarkerData(MARKERDATA *pMarkerData) = 0;
    virtual HRESULT __stdcall GetLineData(long iLine, LINEDATA *pLineData, MARKERDATA *pMarkerData) = 0;
    virtual HRESULT __stdcall ReleaseLineData(LINEDATA *pLineData) = 0;
    virtual HRESULT __stdcall GetLineText(long iStartLine, long iStartIndex, long iEndLine, long iEndIndex, BSTR *pbstrBuf) = 0;
    virtual HRESULT __stdcall CopyLineText(long iStartLine, long iStartIndex, long iEndLine, long iEndIndex, LPWSTR pszBuf, long *pcchBuf) = 0;
    virtual HRESULT __stdcall ReplaceLines(long iStartLine, long iStartIndex, long iEndLine, long iEndIndex, LPCWSTR pszText, long iNewLen, TextSpan *pChangedSpan) = 0;
    virtual HRESULT __stdcall CanReplaceLines(long iStartLine, long iStartIndex, long iEndLine, long iEndIndex, long iNewLen) = 0;
    virtual HRESULT __stdcall CreateLineMarker(long iMarkerType, long iStartLine, long iStartIndex, long iEndLine, long iEndIndex, IVsTextMarkerClient *pClient, IVsTextLineMarker **ppMarker) = 0;
    virtual HRESULT __stdcall EnumMarkers(long iStartLine, long iStartIndex, long iEndLine, long iEndIndex, long iMarkerType, DWORD dwFlags, IVsEnumLineMarkers **ppEnum) = 0;
    virtual HRESULT __stdcall FindMarkerByLineIndex(long iMarkerType, long iStartingLine, long iStartingIndex, DWORD dwFlags, IVsTextLineMarker **ppMarker) = 0;
    virtual HRESULT __stdcall AdviseTextLinesEvents(IVsTextLinesEvents *pSink, DWORD *pdwCookie) = 0;
    virtual HRESULT __stdcall UnadviseTextLinesEvents(DWORD dwCookie) = 0;
    virtual HRESULT __stdcall GetPairExtents(const TextSpan *pSpanIn, TextSpan *pSpanOut) = 0;
    virtual HRESULT __stdcall ReloadLines(long iStartLine, long iStartIndex, long iEndLine, long iEndIndex, LPCWSTR pszText, long iNewLen, TextSpan *pChangedSpan) = 0;
    virtual HRESULT __stdcall IVsTextLinesReserved1(long iLine, LINEDATA *pLineData, BOOL fAttributes) = 0;
    virtual HRESULT __stdcall GetLineDataEx(DWORD dwFlags, long iLine, long iStartIndex, long iEndIndex, LINEDATAEX *pLineData, MARKERDATA *pMarkerData) = 0;
    virtual HRESULT __stdcall ReleaseLineDataEx(LINEDATAEX *pLineData) = 0;
    virtual HRESULT __stdcall CreateEditPoint(long iLine, long iIndex, IDispatch **ppEditPoint) = 0;
    virtual HRESULT __stdcall ReplaceLinesEx(DWORD dwFlags, long iStartLine, long iStartIndex, long iEndLine, long iEndIndex, LPCWSTR pszText, long iNewLen, TextSpan *pChangedSpan) = 0;
    virtual HRESULT __stdcall CreateTextPoint(long iLine, long iIndex, IDispatch **ppTextPoint) = 0;
  };
#pragma endregion

  // ================================ IVsOutputWindowPane ================================
#pragma region
  const GUID IID_IVsOutputWindowPane = { 0x9B878A55, 0x296A, 0x404D, { 0x80, 0xc4, 0x14, 0x68, 0xbb, 0x7c, 0xdc, 0x43 } };

  enum VSTASKPRIORITY {
    TP_HIGH = 0,
    TP_NORMAL = (TP_HIGH + 1),
    TP_LOW = (TP_NORMAL + 1)
  };

  enum VSTASKCATEGORY {
    CAT_ALL = 1,
    CAT_BUILDCOMPILE = 10,
    CAT_COMMENTS = 20,
    CAT_CODESENSE = 30,
    CAT_SHORTCUTS = 40,
    CAT_USER = 50,
    CAT_MISC = 60,
    CAT_HTML = 70
  };

  enum VSTASKBITMAP {
    BMP_COMPILE = -1,
    BMP_SQUIGGLE = -2,
    BMP_COMMENT = -3,
    BMP_SHORTCUT = -4,
    BMP_USER = -5
  };

  struct IVsOutputWindowPane {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsOutputWindowPane ==
    virtual HRESULT __stdcall OutputString(LPCOLESTR pszOutputString) = 0;
    virtual HRESULT __stdcall Activate() = 0;
    virtual HRESULT __stdcall Hide() = 0;
    virtual HRESULT __stdcall Clear() = 0;
    virtual HRESULT __stdcall FlushToTaskList() = 0;
    virtual HRESULT __stdcall OutputTaskItemString(LPCOLESTR pszOutputString, VSTASKPRIORITY nPriority, VSTASKCATEGORY nCategory, LPCOLESTR pszSubcategory, VSTASKBITMAP nBitmap, LPCOLESTR pszFilename, ULONG nLineNum, LPCOLESTR pszTaskItemText) = 0;
    virtual HRESULT __stdcall OutputTaskItemStringEx(LPCOLESTR pszOutputString, VSTASKPRIORITY nPriority, VSTASKCATEGORY nCategory, LPCOLESTR pszSubcategory, VSTASKBITMAP nBitmap, LPCOLESTR pszFilename, ULONG nLineNum, LPCOLESTR pszTaskItemText, LPCOLESTR pszLookupKwd) = 0;
    virtual HRESULT __stdcall GetName(BSTR *pbstrPaneName) = 0;
    virtual HRESULT __stdcall SetName(LPCOLESTR pszPaneName) = 0;
    virtual HRESULT __stdcall OutputStringThreadSafe(LPCOLESTR pszOutputString) = 0;
  };
#pragma endregion

  // ================================ IVsBuildStatusCallback ================================
#pragma region
  const GUID IID_IVsBuildStatusCallback = { 0xA17326AD, 0xC97B, 0x4278, { 0x86, 0xe2, 0x72, 0x16, 0x3c, 0x4c, 0x6a, 0x8c } };

  struct IVsBuildStatusCallback {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsBuildStatusCallback ==
    virtual HRESULT __stdcall BuildBegin(BOOL *pfContinue) = 0;
    virtual HRESULT __stdcall BuildEnd(BOOL fSuccess) = 0;
    virtual HRESULT __stdcall Tick(BOOL *pfContinue) = 0;
  };
#pragma endregion

  // ================================ IVsBuildableProjectCfg ================================
#pragma region
  const GUID IID_IVsBuildableProjectCfg = { 0x8588E475, 0xBB33, 0x4763, { 0xb4, 0xba, 0x03, 0x22, 0xf8, 0x39, 0xaa, 0x3c } };

  struct IVsBuildableProjectCfg {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsBuildableProjectCfg ==
    virtual HRESULT __stdcall get_ProjectCfg(IVsProjectCfg **ppIVsProjectCfg) = 0;
    virtual HRESULT __stdcall AdviseBuildStatusCallback(IVsBuildStatusCallback *pIVsBuildStatusCallback, DWORD *pdwCookie) = 0;
    virtual HRESULT __stdcall UnadviseBuildStatusCallback(DWORD dwCookie) = 0;
    virtual HRESULT __stdcall StartBuild(IVsOutputWindowPane *pIVsOutputWindowPane, DWORD dwOptions) = 0;
    virtual HRESULT __stdcall StartClean(IVsOutputWindowPane *pIVsOutputWindowPane, DWORD dwOptions) = 0;
    virtual HRESULT __stdcall StartUpToDateCheck(IVsOutputWindowPane *pIVsOutputWindowPane, DWORD dwOptions) = 0;
    virtual HRESULT __stdcall QueryStatus(BOOL *pfBuildDone) = 0;
    virtual HRESULT __stdcall Stop(BOOL fSync) = 0;
    virtual HRESULT __stdcall Wait(DWORD dwMilliseconds, BOOL fTickWhenMessageQNotEmpty) = 0;
    virtual HRESULT __stdcall QueryStartBuild(DWORD dwOptions, BOOL *pfSupported, BOOL *pfReady) = 0;
    virtual HRESULT __stdcall QueryStartClean(DWORD dwOptions, BOOL *pfSupported, BOOL *pfReady) = 0;
    virtual HRESULT __stdcall QueryStartUpToDateCheck(DWORD dwOptions, BOOL *pfSupported, BOOL *pfReady) = 0;
  };
#pragma endregion

  // ================================ IVsCfgProvider ================================
#pragma region
  const GUID IID_IVsCfgProvider = { 0xEEABD2BE, 0x4F4F, 0x4CCB, { 0x86, 0xad, 0x9f, 0x46, 0x9c, 0x5c, 0x96, 0x86 } };

  enum VSCFGFLAGS {
    CFGFLAG_CfgsAreNotBrowsable = 0x1,
    CFGFLAG_CfgsUseIndependentPages = 0x2
  };

  struct IVsCfgProvider {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsCfgProvider ==
    virtual HRESULT __stdcall GetCfgs(ULONG celt, IVsCfg *rgpcfg[], ULONG *pcActual, VSCFGFLAGS *prgfFlags) = 0;
  };
#pragma endregion

  // ================================ IVsProjectCfgProvider ================================
#pragma region
  const GUID IID_IVsProjectCfgProvider = { 0x803E46E2, 0x6A0D, 0x4D5D, { 0x9f, 0x84, 0x6c, 0xe1, 0x24, 0x8b, 0x06, 0x8d } };

  struct IVsProjectCfgProvider {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsCfgProvider ==
    virtual HRESULT __stdcall GetCfgs(ULONG celt, IVsCfg *rgpcfg[], ULONG *pcActual, VSCFGFLAGS *prgfFlags) = 0;

    // == IVsProjectCfgProvider ==
    virtual HRESULT __stdcall OpenProjectCfg(LPCOLESTR szProjectCfgCanonicalName, IVsProjectCfg **ppIVsProjectCfg) = 0;
    virtual HRESULT __stdcall get_UsesIndependentConfigurations(BOOL *pfUsesIndependentConfigurations) = 0;
  };
#pragma endregion

  // ================================ IEnumHierarchies ================================
#pragma region
  const GUID IID_IEnumHierarchies = { 0xBEC77711, 0x2DF9, 0x44d7, { 0xb4, 0x78, 0xa4, 0x53, 0xc2, 0xe8, 0xa1, 0x34 } };

  struct IEnumHierarchies {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IEnumHierarchies ==
    virtual HRESULT __stdcall Next(ULONG celt, IVsHierarchy **rgelt, ULONG *pceltFetched) = 0;
    virtual HRESULT __stdcall Skip(ULONG celt) = 0;
    virtual HRESULT __stdcall Reset() = 0;
    virtual HRESULT __stdcall Clone(IEnumHierarchies **ppenum) = 0;
  };
#pragma endregion

  // ================================ IVsWindowFrameEvents ================================
#pragma region
  const GUID IID_IVsWindowFrameEvents = { 0x15D6E42B, 0x36AD, 0x4AF9, { 0xa1, 0x44, 0xc6, 0xf0, 0x70, 0x27, 0xa6, 0xed } };

  struct IVsWindowFrameEvents {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsWindowFrameEvents ==
    virtual HRESULT __stdcall OnFrameCreated(IVsWindowFrame *frame) = 0;
    virtual HRESULT __stdcall OnFrameDestroyed(IVsWindowFrame *frame) = 0;
    virtual HRESULT __stdcall OnFrameIsVisibleChanged(IVsWindowFrame *frame, VARIANT_BOOL newIsVisible) = 0;
    virtual HRESULT __stdcall OnFrameIsOnScreenChanged(IVsWindowFrame *frame, VARIANT_BOOL newIsOnScreen) = 0;
    virtual HRESULT __stdcall OnActiveFrameChanged(IVsWindowFrame *oldFrame, IVsWindowFrame *newFrame) = 0;
  };
#pragma endregion

  // ================================ IVsOutput ================================
#pragma region
  const GUID IID_IVsOutput = { 0x0238DCC5, 0x62D6, 0x4DAC, { 0xa9, 0x77, 0x2c, 0x6a, 0x36, 0xc5, 0x02, 0xf4 } };

  struct IVsOutput {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsOutput ==
    virtual HRESULT __stdcall get_DisplayName(BSTR *pbstrDisplayName) = 0;
    virtual HRESULT __stdcall get_CanonicalName(BSTR *pbstrCanonicalName) = 0;
    virtual HRESULT __stdcall get_DeploySourceURL(BSTR *pbstrDeploySourceURL) = 0;
    virtual HRESULT __stdcall get_Type(GUID *pguidType) = 0;
  };
#pragma endregion

  // ================================ IVsEnumOutputs ================================
#pragma region
  const GUID IID_IVsEnumOutputs = { 0x0A8AC2FB, 0x87BC, 0x4795, { 0x8c, 0x8b, 0x47, 0xe8, 0x77, 0xf4, 0x8f, 0xe8 } };

  struct IVsEnumOutputs {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsEnumOutputs ==
    virtual HRESULT __stdcall Reset() = 0;
    virtual HRESULT __stdcall Next(ULONG cElements, IVsOutput *rgpIVsOutput[], ULONG *pcElementsFetched) = 0;
    virtual HRESULT __stdcall Skip(ULONG cElements) = 0;
    virtual HRESULT __stdcall Clone(IVsEnumOutputs **ppIVsEnumOutputs) = 0;
  };
#pragma endregion

  // ================================ IVsCfg ================================
#pragma region
  const GUID IID_IVsCfg = { 0xB8F932A5, 0x5037, 0x48C9, { 0xab, 0x3a, 0xa4, 0xab, 0xba, 0x79, 0x35, 0x8b } };

  struct IVsCfg {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsCfg ==
    virtual HRESULT __stdcall get_DisplayName(BSTR *pbstrDisplayName) = 0;
    virtual HRESULT __stdcall get_IsDebugOnly(BOOL *pfIsDebugOnly) = 0;
    virtual HRESULT __stdcall get_IsReleaseOnly(BOOL *pfIsReleaseOnly) = 0;
  };
#pragma endregion

  // ================================ IVsProjectCfg ================================
#pragma region
  const GUID IID_IVsProjectCfg = { 0x2DBDF061, 0x439B, 0x4822, { 0x97, 0x27, 0xca, 0x3e, 0xd9, 0x18, 0xb6, 0x58 } };

  struct IVsProjectCfg {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsCfg ==
    virtual HRESULT __stdcall get_DisplayName(BSTR *pbstrDisplayName) = 0;
    virtual HRESULT __stdcall get_IsDebugOnly(BOOL *pfIsDebugOnly) = 0;
    virtual HRESULT __stdcall get_IsReleaseOnly(BOOL *pfIsReleaseOnly) = 0;

    // == IVsProjectCfg ==
    virtual HRESULT __stdcall EnumOutputs(IVsEnumOutputs **ppIVsEnumOutputs) = 0;
    virtual HRESULT __stdcall OpenOutput(LPCOLESTR szOutputCanonicalName, IVsOutput **ppIVsOutput) = 0;
    virtual HRESULT __stdcall get_ProjectCfgProvider(IVsProjectCfgProvider **ppIVsProjectCfgProvider) = 0;
    virtual HRESULT __stdcall get_BuildableProjectCfg(IVsBuildableProjectCfg **ppIVsBuildableProjectCfg) = 0;
    virtual HRESULT __stdcall get_CanonicalName(BSTR *pbstrCanonicalName) = 0;
    virtual HRESULT __stdcall get_Platform(GUID *pguidPlatform) = 0;
    virtual HRESULT __stdcall get_IsPackaged(BOOL *pfIsPackaged) = 0;
    virtual HRESULT __stdcall get_IsSpecifyingOutputSupported(BOOL *pfIsSpecifyingOutputSupported) = 0;
    virtual HRESULT __stdcall get_TargetCodePage(UINT *puiTargetCodePage) = 0;
    virtual HRESULT __stdcall get_UpdateSequenceNumber(ULARGE_INTEGER *puliUSN) = 0;
    virtual HRESULT __stdcall get_RootURL(BSTR *pbstrRootURL) = 0;
  };
#pragma endregion

  // ================================ IVsEnumGuids ================================
#pragma region
  const GUID IID_IVsEnumGuids = { 0xBEC804F7, 0xF5DE, 0x4F3E, { 0x8e, 0xbb, 0xda, 0xb2, 0x66, 0x49, 0xf3, 0x3f } };

  struct IVsEnumGuids {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsEnumGuids ==
    virtual HRESULT __stdcall Next(ULONG celt, GUID rgelt[], ULONG *pceltFetched) = 0;
    virtual HRESULT __stdcall Skip(ULONG celt) = 0;
    virtual HRESULT __stdcall Reset() = 0;
    virtual HRESULT __stdcall Clone(IVsEnumGuids **ppenum) = 0;
  };
#pragma endregion

  // ================================ IVsSaveOptionsDlg ================================
#pragma region
  const GUID IID_IVsSaveOptionsDlg = { 0xC3E2ED14, 0x4E64, 0x4c26, { 0x84, 0xD7, 0x68, 0xcc, 0xd0, 0x71, 0xa0, 0xc8 } };

  struct IVsSaveOptionsDlg {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSaveOptionsDlg ==
    virtual HRESULT __stdcall ShowSaveOptionsDlg(DWORD dwReserved, HWND hwndDlgParent, WCHAR *pszFileName) = 0;
  };
#pragma endregion

  // ================================ IVsUIAccelerator ================================
#pragma region
  const GUID IID_IVsUIAccelerator = { 0x4E25556D, 0x941D, 0x4c29, { 0xa1, 0x71, 0x38, 0x4e, 0xa8, 0x4f, 0x67, 0x05 } };

  enum VSUIACCELMODIFIERS {
    VSAM_NONE = 0,
    VSAM_SHIFT = 0x1,
    VSAM_CONTROL = 0x2,
    VSAM_ALT = 0x4,
    VSAM_WINDOWS = 0x8
  };

  struct IVsUIAccelerator {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIAccelerator ==
    virtual HRESULT __stdcall get_Message(MSG *pMsg) = 0;
    virtual HRESULT __stdcall get_Modifiers(VSUIACCELMODIFIERS *pdwModifiers) = 0;
  };
#pragma endregion

  // ================================ IVsUIEnumDataSourceVerbs ================================
#pragma region
  const GUID IID_IVsUIEnumDataSourceVerbs = { 0x51C2FFFB, 0x35FA, 0x4ad2, { 0x81, 0xb1, 0x11, 0x81, 0x6c, 0x48, 0x2a, 0xaa } };

  struct IVsUIEnumDataSourceVerbs {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIEnumDataSourceVerbs ==
    virtual HRESULT __stdcall Next(ULONG celt, BSTR *rgelt, ULONG *pceltFetched) = 0;
    virtual HRESULT __stdcall Skip(ULONG celt) = 0;
    virtual HRESULT __stdcall Reset() = 0;
    virtual HRESULT __stdcall Clone(IVsUIEnumDataSourceVerbs **ppEnum) = 0;
  };
#pragma endregion

  // ================================ IVsUIDispatch ================================
#pragma region
  const GUID IID_IVsUIDispatch = { 0x0DF3E43A, 0x5356, 0x4A33, { 0x8a, 0xc1, 0x3b, 0xe6, 0xe3, 0x33, 0x7c, 0x37 } };

  struct IVsUIDispatch {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIDispatch ==
    virtual HRESULT __stdcall Invoke(LPCOLESTR verb, VARIANT pvaIn, VARIANT *pvaOut) = 0;
    virtual HRESULT __stdcall EnumVerbs(IVsUIEnumDataSourceVerbs **ppEnum) = 0;
  };
#pragma endregion

  // ================================ IVsUISimpleDataSource ================================
#pragma region
  const GUID IID_IVsUISimpleDataSource = { 0x110596DC, 0x7A19, 0x4E04, { 0x91, 0x06, 0x1d, 0xb0, 0x58, 0x0f, 0x77, 0xe9 } };

  struct IVsUISimpleDataSource {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIDispatch ==
    virtual HRESULT __stdcall Invoke(LPCOLESTR verb, VARIANT pvaIn, VARIANT *pvaOut) = 0;
    virtual HRESULT __stdcall EnumVerbs(IVsUIEnumDataSourceVerbs **ppEnum) = 0;

    // == IVsUISimpleDataSource ==
    virtual HRESULT __stdcall Close() = 0;
  };
#pragma endregion

  // ================================ IVsUIElement ================================
#pragma region
  const GUID IID_IVsUIElement = { 0x62C0A03E, 0x4979, 0x4b4e, { 0x90, 0xf0, 0x56, 0xdf, 0x90, 0x52, 0x1f, 0x79 } };

  struct IVsUIElement {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIElement ==
    virtual HRESULT __stdcall get_DataSource(IVsUISimpleDataSource **ppDataSource) = 0;
    virtual HRESULT __stdcall put_DataSource(IVsUISimpleDataSource *pDataSource) = 0;
    virtual HRESULT __stdcall TranslateAccelerator(IVsUIAccelerator *pAccel) = 0;
    virtual HRESULT __stdcall GetUIObject(IUnknown **ppUnk) = 0;
  };
#pragma endregion

  // ================================ IVsToolbarTrayHost ================================
#pragma region
  const GUID IID_IVsToolbarTrayHost = { 0x2B3321EE, 0x693F, 0x4b46, { 0x95, 0x36, 0xe4, 0x4d, 0xad, 0x8c, 0x6e, 0x60 } };

  struct IVsToolbarTrayHost {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsToolbarTrayHost ==
    virtual HRESULT __stdcall AddToolbar(const GUID &pGuid, DWORD dwId) = 0;
    virtual HRESULT __stdcall GetToolbarTray(IVsUIElement **ppToolbarTray) = 0;
    virtual HRESULT __stdcall Close() = 0;
  };
#pragma endregion

  // ================================ IVsImageButton ================================
#pragma region
  const GUID IID_IVsImageButton = { 0x61DF9CCE, 0xE88E, 0x4fe2, { 0x99, 0x76, 0x77, 0xa4, 0xf4, 0x78, 0xe2, 0x4b } };

  struct VSDRAWITEMSTRUCT {
    UINT CtlType;
    UINT CtlID;
    UINT itemID;
    UINT itemAction;
    UINT itemState;
    HWND hwndItem;
    HDC hDC;
    RECT rcItem;
    ULONG_PTR itemData;
  };

  struct IVsImageButton {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsImageButton ==
    virtual HRESULT __stdcall Draw(VSDRAWITEMSTRUCT *pDrawItemStruct, BOOL fHot) = 0;
  };
#pragma endregion

  // ================================ IVsGradient ================================
#pragma region
  const GUID IID_IVsGradient = { 0xfd3f680a, 0xd5c1, 0x437a, { 0x8a, 0x21, 0x80, 0x84, 0x31, 0x0b, 0xf0, 0x37 } };

  struct IVsGradient {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsGradient ==
    virtual HRESULT __stdcall DrawGradient(HWND hwnd, HDC hdc, RECT *gradientRect, RECT *sliceRect) = 0;
    virtual HRESULT __stdcall GetGradientVector(int cVector, COLORREF *rgVector) = 0;
  };
#pragma endregion

  // ================================ IVsBatchProjectActionContext ================================
#pragma region
  const GUID IID_IVsBatchProjectActionContext = { 0x49ED97F3, 0xEAE0, 0x47AC, { 0x9A, 0x2E, 0xdc, 0x15, 0xd0, 0x45, 0x9f, 0x7b } };

  enum VSBatchProjectAction {
    BPA_NONE = 0,
    BPA_UNLOAD = 1,
    BPA_LOAD = 2,
    BPA_RELOAD = 3,
    BPA_RELOADSOLUTION = 4
  };

  enum VSBatchProjectActionResult {
    BPAR_UNCHANGED = 0,
    BPAR_UNLOADED = 1,
    BPAR_RELOADED = 3,
    BPAR_SELFRELOADED = 4,
    BPAR_ERROR = 5
  };

  struct IVsBatchProjectActionContext {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsBatchProjectActionContext ==
    virtual HRESULT __stdcall GetAction(VSBatchProjectAction *pAction) = 0;
    virtual HRESULT __stdcall GetProjectsCount(DWORD *pdwProjects) = 0;
    virtual HRESULT __stdcall GetProjectsInfo(DWORD dwProjects, GUID rgProjectsGuid[], VSBatchProjectActionResult rgExpectedResult[]) = 0;
    virtual HRESULT __stdcall GetCurrentResult(const GUID &guidProject, VSBatchProjectActionResult *pResult) = 0;
  };
#pragma endregion

  // ================================ IVsPropertyBag ================================
#pragma region
  const GUID IID_IVsPropertyBag = { 0xaaeeac4c, 0x3bf3, 0x492c, { 0x92, 0x7d, 0x84, 0xab, 0x7d, 0x93, 0xd6, 0xdf } };

  struct IVsPropertyBag {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsPropertyBag ==
    virtual HRESULT __stdcall GetValue(LPCOLESTR szName, VARIANT *pVarValue) = 0;
    virtual HRESULT __stdcall SetValue(LPCOLESTR szName, VARIANT *pVarValue) = 0;
  };
#pragma endregion

    // ================================ IVsBrowseProjectLocation ================================
#pragma region
    const GUID IID_IVsBrowseProjectLocation = { 0x368FC032, 0xAE91, 0x44a2, { 0xbe, 0x6b, 0x09, 0x3a, 0x8a, 0x9e, 0x63, 0xcc } };

  struct IVsBrowseProjectLocation {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsBrowseProjectLocation ==
    virtual HRESULT __stdcall BrowseProjectLocation(LPCOLESTR pszStartDirectory, BSTR *pbstrProjectLocation) = 0;
  };
#pragma endregion

  // ================================ IVsUpdateSolutionEventsAsyncCallback ================================
#pragma region
  const GUID IID_IVsUpdateSolutionEventsAsyncCallback = { 0x02D0878C, 0x53F5, 0x4CE9, { 0xb5, 0x5c, 0x35, 0x77, 0xda, 0xe6, 0x47, 0x61 } };

  struct IVsUpdateSolutionEventsAsyncCallback {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUpdateSolutionEventsAsyncCallback ==
    virtual HRESULT __stdcall CompleteLastUpdateAction() = 0;
  };
#pragma endregion

    // ================================ IVsUpdateSolutionEventsAsync ================================
#pragma region
    const GUID IID_IVsUpdateSolutionEventsAsync = { 0x703ECC2C, 0x7631, 0x46A9, { 0xad, 0x1e, 0x19, 0xd9, 0x59, 0x2c, 0x7a, 0x6b } };

  struct IVsUpdateSolutionEventsAsync {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUpdateSolutionEventsAsync ==
    virtual HRESULT __stdcall UpdateSolution_EndLastUpdateActionAsync(IVsUpdateSolutionEventsAsyncCallback *pCallback) = 0;
  };
#pragma endregion

      // ================================ IVsTaskBody ================================
#pragma region
  const GUID IID_IVsTaskBody = { 0x05a07459, 0x551f, 0x4cdf, { 0xb3, 0x8a, 0x16, 0x08, 0x9d, 0x08, 0x31, 0x10 } };

  struct IVsTaskBody {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTaskBody ==
    virtual HRESULT __stdcall OnContextChanged(GUID uiContext, VARIANT_BOOL active) = 0;
  };
#pragma endregion

      // ================================ IVsTask ================================
#pragma region
  const GUID IID_IVsTask = { 0x0b98eab8, 0x00bb, 0x45d0, { 0xae, 0x2f, 0x3d, 0xe3, 0x5c, 0xd6, 0x82, 0x35 } };

  // From Microsoft.VisualStudio.Shell.Interop.11.0.xml
  // Specifies how the task is run.
  enum VSTASKRUNCONTEXT {
    // Runs the task on the background thread pool with normal priority.
    VSTC_BACKGROUNDTHREAD = 0,
    // Runs the task on UI thread using RPC callback to be executed as soon as possible. Note: This context may cause reentrancy.
    VSTC_UITHREAD_SEND = 1,
    // Runs the task on the UI thread using background priority(that is, below user input). Tasks are scheduled even while modal dialogs are open. Tasks are scheduled to occur when no user input is pending, plus a short delay. Appropriate for short tasks or slightly longer tasks.
    VSTC_UITHREAD_BACKGROUND_PRIORITY = 2,
    // Runs the task on the UI thread when Visual Studio is idle. Tasks are not scheduled till most modal dialogs have been dismissed. Appropriate for very short running tasks.
    VSTC_UITHREAD_IDLE_PRIORITY = 3,
    // Runs the task on the current context(that is, the UI thread or the background thread).
    VSTC_CURRENTCONTEXT = 4,
    // Runs the task on background thread pool and sets the background mode on the thread while the task is running. This is useful for IO heavy background tasks that are not time critical.
    VSTC_BACKGROUNDTHREAD_LOW_IO_PRIORITY = 5,
    // Runs the task on UI thread using Dispatcher with Normal priority.
    VSTC_UITHREAD_NORMAL_PRIORITY = 6
  };

  enum VSTASKCONTINUATIONOPTIONS {
    VSTCO_None = 0,
    VSTCO_PreferFairness = 1,
    VSTCO_LongRunning = 2,
    VSTCO_AttachedToParent = 4,
    VSTCO_DenyChildAttach = 8,
    VSTCO_LazyCancelation = 32,
    VSTCO_NotOnRanToCompletion = 0x10000,
    VSTCO_NotOnFaulted = 0x20000,
    VSTCO_OnlyOnCanceled = 0x30000,
    VSTCO_NotOnCanceled = 0x40000,
    VSTCO_OnlyOnFaulted = 0x50000,
    VSTCO_OnlyOnRanToCompletion = 0x60000,
    VSTCO_ExecuteSynchronously = 0x80000,
    VSTCO_IndependentlyCanceled = 0x40000000,
    VSTCO_NotCancelable = 0x80000000,
    VSTCO_Default = VSTCO_NotOnFaulted
  };

  enum VSTASKWAITOPTIONS {
    VSTWO_None = 0,
    VSTWO_AbortOnTaskCancellation = 0x1
  };

  struct IVsTask {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTask ==
    virtual HRESULT __stdcall ContinueWith(VSTASKRUNCONTEXT context, IVsTaskBody *pTaskBody, IVsTask **ppTask) = 0;
    virtual HRESULT __stdcall ContinueWithEx(VSTASKRUNCONTEXT context, VSTASKCONTINUATIONOPTIONS options, IVsTaskBody *pTaskBody, VARIANT pAsyncState, IVsTask **ppTask) = 0;
    virtual HRESULT __stdcall Start() = 0;
    virtual HRESULT __stdcall Cancel() = 0;
    virtual HRESULT __stdcall GetResult(VARIANT *pResult) = 0;
    virtual HRESULT __stdcall AbortIfCanceled() = 0;
    virtual HRESULT __stdcall Wait() = 0;
    virtual HRESULT __stdcall WaitEx(int millisecondsTimeout, VSTASKWAITOPTIONS options, VARIANT_BOOL *pTaskCompleted) = 0;
    virtual HRESULT __stdcall get_IsFaulted(VARIANT_BOOL *pResult) = 0;
    virtual HRESULT __stdcall get_IsCompleted(VARIANT_BOOL *pResult) = 0;
    virtual HRESULT __stdcall get_IsCanceled(VARIANT_BOOL *pResult) = 0;
    virtual HRESULT __stdcall get_AsyncState(VARIANT *pAsyncState) = 0;
    virtual HRESULT __stdcall get_Description(BSTR *ppDescriptionText) = 0;
    virtual HRESULT __stdcall put_Description(LPCOLESTR pDescriptionText) = 0;
  };
#pragma endregion

    // ================================ IVsUIContextEvents ================================
#pragma region
  const GUID IID_IVsUIContextEvents = { 0x0393d191, 0x94ac, 0x4997, { 0x93, 0x10, 0x2e, 0xac, 0x67, 0x49, 0x58, 0x16 } };

  struct IVsUIContextEvents {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIContextEvents ==
    virtual HRESULT __stdcall OnContextChanged(GUID uiContext, VARIANT_BOOL active) = 0;
  };
#pragma endregion

  // ================================ IVsWindowFrame ================================
#pragma region 
  const GUID IID_IVsWindowFrame = { 0x11138F8A, 0x38C0, 0x4436, { 0xB5, 0xA6, 0x2f, 0x5E, 0xF2, 0xC3, 0xE2, 0x42 } };

  enum VSRDTSAVEOPTIONS {
    RDTSAVEOPT_SaveIfDirty = 0,
    RDTSAVEOPT_PromptSave = 0x1,
    RDTSAVEOPT_ForceSave = 0x2,
    RDTSAVEOPT_SaveNoChildren = 0x4,
    RDTSAVEOPT_SaveOnlyChildren = 0x8,
    RDTSAVEOPT_ActivateDocOnErr = 0x10,
    RDTSAVEOPT_DocClose = 0x10000,
    RDTSAVEOPT_Reserved = 0xffff0000
  };

  enum FRAMECLOSE {
    FRAMECLOSE_NoSave = (0x100 | RDTSAVEOPT_DocClose),
    FRAMECLOSE_SaveIfDirty = (0x200 | RDTSAVEOPT_DocClose),
    FRAMECLOSE_PromptSave = (0x400 | RDTSAVEOPT_DocClose)
  };

  enum VSSETFRAMEPOS {
    SFP_maskFrameMode = 0xf,
    SFP_fDock = 0x1,
    SFP_fTab = 0x2,
    SFP_fFloat = 0x3,
    SFP_fMdiChild = 0x4,
    SFP_maskPosition = 0xf0,
    SFP_fDockTop = 0x10,
    SFP_fDockBottom = 0x20,
    SFP_fDockLeft = 0x30,
    SFP_fDockRight = 0x40,
    SFP_fTabFirst = 0x10,
    SFP_fTabLast = 0x20,
    SFP_fTabPrevious = 0x30,
    SFP_fTabNext = 0x40,
    SFP_fSize = 0x40000000,
    SFP_fMove = 0x80000000
  };

  enum VSFPROPID {
    VSFPROPID_NIL = -1,
    VSFPROPID_LAST = -3000,
    VSFPROPID_Type = -3000,
    VSFPROPID_DocView = -3001,
    VSFPROPID_SPFrame = -3002,
    VSFPROPID_SPProjContext = -3003,
    VSFPROPID_Caption = -3004,
    VSFPROPID_WindowState = -3007,
    VSFPROPID_FrameMode = -3008,
    VSFPROPID_IsWindowTabbed = -3009,
    VSFPROPID_UserContext = -3010,
    VSFPROPID_ViewHelper = -3011,
    VSFPROPID_ShortCaption = -3012,
    VSFPROPID_WindowHelpKeyword = -3013,
    VSFPROPID_WindowHelpCmdText = -3014,
    VSFPROPID_DocCookie = -4000,
    VSFPROPID_OwnerCaption = -4001,
    VSFPROPID_EditorCaption = -4002,
    VSFPROPID_pszMkDocument = -4003,
    VSFPROPID_DocData = -4004,
    VSFPROPID_Hierarchy = -4005,
    VSFPROPID_ItemID = -4006,
    VSFPROPID_CmdUIGuid = -4007,
    VSFPROPID_CreateDocWinFlags = -4008,
    VSFPROPID_guidEditorType = -4009,
    VSFPROPID_pszPhysicalView = -4010,
    VSFPROPID_InheritKeyBindings = -4011,
    VSFPROPID_RDTDocData = -4012,
    VSFPROPID_AltDocData = -4013,
    VSFPROPID_GuidPersistenceSlot = -5000,
    VSFPROPID_GuidAutoActivate = -5001,
    VSFPROPID_CreateToolWinFlags = -5002,
    VSFPROPID_ExtWindowObject = -5003,
    VSFPROPID_MultiInstanceToolNum = -5004,
    VSFPROPID_BitmapResource = -5006,
    VSFPROPID_BitmapIndex = -5007,
    VSFPROPID_ToolbarHost = -5008,
    VSFPROPID_HideToolwinContainer = -5009,
    VSFPROPID_FIRST = -5009
  };

  struct IVsWindowFrame {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsWindowFrame ==
    virtual HRESULT __stdcall Show() = 0;
    virtual HRESULT __stdcall Hide() = 0;
    virtual HRESULT __stdcall IsVisible() = 0;
    virtual HRESULT __stdcall ShowNoActivate() = 0;
    virtual HRESULT __stdcall CloseFrame(FRAMECLOSE grfSaveOptions) = 0;
    virtual HRESULT __stdcall SetFramePos(VSSETFRAMEPOS dwSFP, const GUID &rguidRelativeTo, int x, int y, int cx, int cy) = 0;
    virtual HRESULT __stdcall GetFramePos(VSSETFRAMEPOS *pdwSFP, GUID *pguidRelativeTo, int *px, int *py, int *pcx, int *pcy) = 0;
    virtual HRESULT __stdcall GetProperty(VSFPROPID propid, VARIANT *pvar) = 0;
    virtual HRESULT __stdcall SetProperty(VSFPROPID propid, VARIANT var) = 0;
    virtual HRESULT __stdcall GetGuidProperty(VSFPROPID propid, GUID *pguid) = 0;
    virtual HRESULT __stdcall SetGuidProperty(VSFPROPID propid, const GUID &rguid) = 0;
    virtual HRESULT __stdcall QueryViewInterface(const IID &riid, void **ppv) = 0;
    virtual HRESULT __stdcall IsOnScreen(BOOL *pfOnScreen) = 0;
  };
#pragma endregion

  // ================================ IVsProject ================================
#pragma region
  const GUID IID_IVsProject = { 0xCD4028ED, 0xC4D8, 0x44ba, { 0x89, 0x0f, 0xef, 0xfb, 0x02, 0xa3, 0x80, 0xc6 } };

  enum VSDOCUMENTPRIORITY {
    DP_Intrinsic = 60,
    DP_Standard = 50,
    DP_NonMember = 40,
    DP_CanAddAsNonMember = 30,
    DP_External = 20,
    DP_CanAddAsExternal = 10,
    DP_Unsupported = 0
  };

  enum VSADDITEMOPERATION {
    VSADDITEMOP_OPENFILE = 1,
    VSADDITEMOP_CLONEFILE = 2,
    VSADDITEMOP_RUNWIZARD = 3,
    VSADDITEMOP_LINKTOFILE = 4
  };

  enum VSADDRESULT {
    ADDRESULT_Success = -1,
    ADDRESULT_Failure = 0,
    ADDRESULT_Cancel = 1
  };

  struct __declspec(novtable) IVsProject {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsProject ==
    virtual HRESULT __stdcall IsDocumentInProject(LPCOLESTR pszMkDocument, BOOL * pfFound, VSDOCUMENTPRIORITY * pdwPriority, DWORD *pitemid) = 0;
    virtual HRESULT __stdcall GetMkDocument(DWORD itemid, BSTR *pbstrMkDocument) = 0;
    virtual HRESULT __stdcall OpenItem(DWORD itemid, const GUID &rguidLogicalView, IUnknown *punkDocDataExisting, IVsWindowFrame **ppWindowFrame) = 0;
    virtual HRESULT __stdcall GetItemContext(DWORD itemid, IServiceProvider **ppSP) = 0;
    virtual HRESULT __stdcall GenerateUniqueItemName(DWORD itemidLoc, LPCOLESTR pszExt, LPCOLESTR pszSuggestedRoot, BSTR *pbstrItemName) = 0;
    virtual HRESULT __stdcall AddItem(DWORD itemidLoc, VSADDITEMOPERATION dwAddItemOperation, LPCOLESTR pszItemName, ULONG cFilesToOpen, LPCOLESTR rgpszFilesToOpen[], HWND hwndDlgOwner, VSADDRESULT *pResult) = 0;
  };
#pragma endregion

  // ================================ IVsTrackProjectDocumentsEvents2 ================================
#pragma region
  const GUID IID_IVsTrackProjectDocumentsEvents2 = { 0x53544c4d, 0x663d, 0x11d3, { 0xa6, 0x0d, 0x00, 0x50, 0x04, 0x77, 0x5a, 0xb1 } };

  enum VSQUERYADDFILEFLAGS {
    VSQUERYADDFILEFLAGS_NoFlags = 0,
    VSQUERYADDFILEFLAGS_IsSpecialFile = 1,
    VSQUERYADDFILEFLAGS_IsNestedProjectFile = 2
  };

  enum VSQUERYADDFILERESULTS {
    VSQUERYADDFILERESULTS_AddOK = 0,
    VSQUERYADDFILERESULTS_AddNotOK = 1
  };

  enum VSQUERYREMOVEFILEFLAGS {
    VSQUERYREMOVEFILEFLAGS_NoFlags = 0,
    VSQUERYREMOVEFILEFLAGS_IsSpecialFile = 1,          // "Special" file i.e. an invisible file associated with another file in the project
    VSQUERYREMOVEFILEFLAGS_IsNestedProjectFile = 2     // Nested project (file) i.e. the file sitting on the root node of a nested project
  };

  enum VSQUERYREMOVEFILEFLAGS2 {
    //VSQUERYREMOVEFILEFLAGS_NoFlags = 0,
    //VSQUERYREMOVEFILEFLAGS_IsSpecialFile = 1, 
    //VSQUERYREMOVEFILEFLAGS_IsNestedProjectFile = 2, 
    VSQUERYREMOVEFILEFLAGS_IsRemovedFromProjectOnly = 4 // Distinguish: "Remove From Project" VS. "Delete". If it's set, which means the file is just removed from project, but still exists in dis
  };

  enum VSQUERYREMOVEDIRECTORYFLAGS {
    VSQUERYREMOVEDIRECTORYFLAGS_padding = 0 // No flags yet
  };

  enum VSQUERYREMOVEDIRECTORYFLAGS2 {
    // VSQUERYREMOVEDIRECTORYFLAGS_padding, 
    VSQUERYREMOVEDIRECTORYFLAGS_IsRemovedFromProjectOnly = 1 // Distinguish: "Remove From Project" VS. "Delete". If it's set, which means the directory is just removed from project, but still exists in dis
  };

  enum VSREMOVEFILEFLAGS {
    VSREMOVEFILEFLAGS_NoFlags = 0,
    VSREMOVEFILEFLAGS_IsDirectoryBased = 1,
    VSREMOVEFILEFLAGS_RemoveFromSourceControlDoneExternally = 2,  // Already been removed from source control
    VSREMOVEFILEFLAGS_IsSpecialFile = 4,                          // "Special" file i.e. an invisible file associated with another file in the project
    VSREMOVEFILEFLAGS_IsNestedProjectFile = 8                     // Nested project (file) i.e. the file sitting on the root node of a nested project
  };

  enum VSREMOVEFILEFLAGS2 {
    //VSREMOVEFILEFLAGS_NoFlags = 0,
    //VSREMOVEFILEFLAGS_IsDirectoryBased = 1,
    //VSREMOVEFILEFLAGS_RemoveFromSourceControlDoneExternally = 2, 
    //VSREMOVEFILEFLAGS_IsSpecialFile = 4, 
    //VSREMOVEFILEFLAGS_IsNestedProjectFile = 8, 
    VSREMOVEFILEFLAGS_IsRemovedFromProjectOnly = 16 // Distinguish: "Remove From Project" VS. "Delete". If it's set, which means the file is just removed from project, but still exists in disk
  };

  enum VSREMOVEDIRECTORYFLAGS {
    VSREMOVEDIRECTORYFLAGS_NoFlags = 0,
    VSREMOVEDIRECTORYFLAGS_IsDirectoryBased = 1,
    VSREMOVEDIRECTORYFLAGS_RemoveFromSourceControlDoneExternally = 2 // Dir has already been removed from source control
  };

  enum VSREMOVEDIRECTORYFLAGS2 {
    //VSREMOVEDIRECTORYFLAGS_NoFlags = 0,
    //VSREMOVEDIRECTORYFLAGS_IsDirectoryBased = 1,
    //VSREMOVEDIRECTORYFLAGS_RemoveFromSourceControlDoneExternally = 2, 
    VSREMOVEDIRECTORYFLAGS_IsRemovedFromProjectOnly = 4 // Distinguish: "Remove From Project" VS. "Delete". If it's set, which means the directory is just removed from project, but still exists in disk
  };

  enum VSQUERYREMOVEFILERESULTS {
    VSQUERYREMOVEFILERESULTS_RemoveOK = 0,
    VSQUERYREMOVEFILERESULTS_RemoveNotOK = 1
  };

  enum VSQUERYREMOVEDIRECTORYRESULTS {
    VSQUERYREMOVEDIRECTORYRESULTS_RemoveOK = 0,
    VSQUERYREMOVEDIRECTORYRESULTS_RemoveNotOK = 1
  };
  enum VSADDFILEFLAGS {
    VSADDFILEFLAGS_NoFlags = 0,
    VSADDFILEFLAGS_AddToSourceControlDoneExternally = 1,
    VSADDFILEFLAGS_IsSpecialFile = 2,
    VSADDFILEFLAGS_IsNestedProjectFile = 4
  };

  enum VSADDDIRECTORYFLAGS {
    VSADDDIRECTORYFLAGS_NoFlags = 0,
    VSADDDIRECTORYFLAGS_AddToSourceControlDoneExternally = 1
  };

  enum VSQUERYRENAMEFILEFLAGS {
    VSQUERYRENAMEFILEFLAGS_NoFlags = 0,
    VSQUERYRENAMEFILEFLAGS_IsSpecialFile = 1,
    VSQUERYRENAMEFILEFLAGS_IsNestedProjectFile = 2,
    VSQUERYRENAMEFILEFLAGS_Directory = 0x20
  };

  enum VSQUERYRENAMEFILERESULTS {
    VSQUERYRENAMEFILERESULTS_RenameOK = 0,
    VSQUERYRENAMEFILERESULTS_RenameNotOK = 1
  };

  enum VSRENAMEFILEFLAGS {
    VSRENAMEFILEFLAGS_NoFlags = 0,
    VSRENAMEFILEFLAGS_FromShellCommand = 0x1,
    VSRENAMEFILEFLAGS_FromScc = 0x2,
    VSRENAMEFILEFLAGS_FromFileChange = 0x4,
    VSRENAMEFILEFLAGS_WasQueried = 0x8,
    VSRENAMEFILEFLAGS_AlreadyOnDisk = 0x10,
    VSRENAMEFILEFLAGS_Directory = 0x20,
    VSRENAMEFILEFLAGS_RenameInSourceControlDoneExternally = 0x40,
    VSRENAMEFILEFLAGS_IsSpecialFile = 0x80,
    VSRENAMEFILEFLAGS_IsNestedProjectFile = 0x100,
    VSRENAMEFILEFLAGS_ALL = 0x1ff,
    VSRENAMEFILEFLAGS_INVALID = 0xfffffe00
  };

  enum VSQUERYRENAMEDIRECTORYFLAGS {
    VSQUERYRENAMEDIRECTORYFLAGS_padding = 0
  };

  enum VSQUERYRENAMEDIRECTORYRESULTS {
    VSQUERYRENAMEDIRECTORYRESULTS_RenameOK = 0,
    VSQUERYRENAMEDIRECTORYRESULTS_RenameNotOK = 1
  };

  enum VSRENAMEDIRECTORYFLAGS {
    VSRENAMEDIRECTORYFLAGS_NoFlags = 0,
    VSRENAMEDIRECTORYFLAGS_RenameInSourceControlDoneExternally = 1
  };

  enum VSQUERYADDDIRECTORYFLAGS {
    VSQUERYADDDIRECTORYFLAGS_padding = 0
  };

  enum VSQUERYADDDIRECTORYRESULTS {
    VSQUERYADDDIRECTORYRESULTS_AddOK = 0,
    VSQUERYADDDIRECTORYRESULTS_AddNotOK = 1
  };

  // Notifies clients of changes made to project files or directories.
  // Projects use this service to receive events that happen on files and folders
  //
  // Before adding, renaming, or deleting a file or directory, the project must call the appropriate TrackProjectDocuments2 method in order to check whether this operation is allowed.  
  // When the operation is completion, the project must notify the TrackProjectDocuments2 service, so that appropriate notifications can be sent out.
  // 
  // This interface sinks these events.  
  // When the project issues the calls to TrackProjectDocuments2, the service passes the notifications and queries on to the sinks.
  // 
  // When you subscribe to the IVsTrackProjectDocumentEvents2 events, you will receive event notification for all projects. 
  // Generally, you will not receive batched notification of these events unless two projects coordinate, as in the case of a nested project and the parent project.
  // Before adding, renaming, or deleting a file or directory, each project must call the appropriate OnQuery* method from IVsTrackProjectDocuments2 to check whether the operation is allowed. 
  // After the operation is completed, the project must then notify the OnAfter* method in IVsTrackProjectDocuments2. 
  // The environment sends out the appropriate event notifications after each call.
  //
  // The parameters to these functions generally consist of: 
  // (Note all arrays have an associated parameter that is the size of that array.)
  // 
  // a) The relevant IVsProject, or array of IVsProject pointers
  // 
  // b) Flags regarding the operation taking place
  // 
  // c) Array of documents, sorted by project.  If there is only one project, then the
  // ordering of files does not matter.  If there is more than one project, the
  // files must be grouped by their associated projects.
  //
  // d) Array of "first indices".  These indices relate the array of projects to the
  // array of documents.  Their is one "first index" for each project, which points
  // to the first file in the array of documents that is controled by that project.
  // Since the array of documents is sorted by projects, all the indices greater than
  // one first index, and less than the next first index belong to a given project.
  // 
  // Example: 
  // Projects:     Visual Basic        |  Visual C++     |    Visual C#
  // Indices:            0             |      5          |       8
  // Documents: a, b, c, d, e (0-4)    | f, g, h (5-7)   | i, j, k, l (8-11)
  // Documents a,b,c,d,e  belong to Visual Basic
  // Documents f,g,h      belong to Visual C++
  // Documents i, j, k, l belong to Visual C#
  //
  // e) Result flags: Is this operation allowed?
  //
  // A source control package implements this interface if it needs to track changes in a project, such as when files or directories are added, removed, or renamed. 
  // It is recommended that this interface be implemented; otherwise, the user may need to manually refresh the source control display to see any changes in status.
  //
  // Called by the environment in response to the addition, removal, or renaming of files or directories in a project.
  struct __declspec(novtable) IVsTrackProjectDocumentsEvents2 {
    // == IUnknown methods ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTrackProjectDocumentsEvents2 methods ==

    // Notifies the client when a project has requested to add files.
    // pProject: Project requesting to add files.
    // cFiles: Number of files to add.
    // rgpszMkDocuments: Array of files to add to the project.
    // rgFlags: Array of flags associated with each file.
    // pSummaryResult: Summary result object indicating if the add operation can proceed.
    // rgResults: Array of results for each individual file.
    virtual HRESULT __stdcall OnQueryAddFiles(
      IVsProject *pProject,
      int cFiles,
      const LPCOLESTR rgpszMkDocuments[],
      const VSQUERYADDFILEFLAGS rgFlags[],
      VSQUERYADDFILERESULTS *pSummaryResult,
      VSQUERYADDFILERESULTS rgResults[]
    ) = 0;

    // Notifies the client after a project has added files.
    // cProjects: Number of projects involved.
    // cFiles: Number of files added.
    // rgpProjects: Array of projects involved.
    // rgFirstIndices: Array of indices identifying which project each file belongs to.
    // rgpszMkDocuments: Array of file paths added.
    // rgFlags: Array of flags associated with each file.
    virtual HRESULT __stdcall OnAfterAddFilesEx(
      int cProjects,
      int cFiles,
      IVsProject *rgpProjects[],
      const int rgFirstIndices[],
      const LPCOLESTR rgpszMkDocuments[],
      const VSADDFILEFLAGS rgFlags[]
    ) = 0;

    // Notifies the client after directories are added to the project.
    // cProjects: Number of projects involved.
    // cDirectories: Number of directories added.
    // rgpProjects: Array of projects involved.
    // rgFirstIndices: Array of indices identifying which project each directory belongs to.
    // rgpszMkDocuments: Array of directory paths added.
    // rgFlags: Array of flags associated with each directory.
    virtual HRESULT __stdcall OnAfterAddDirectoriesEx(
      int cProjects,
      int cDirectories,
      IVsProject *rgpProjects[],
      const int rgFirstIndices[],
      const LPCOLESTR rgpszMkDocuments[],
      const VSADDDIRECTORYFLAGS rgFlags[]
    ) = 0;

    // Notifies the client after files have been removed from the project.
    // cProjects: Number of projects involved.
    // cFiles: Number of files removed.
    // rgpProjects: Array of projects involved.
    // rgFirstIndices: Array of indices identifying which project each file belonged to.
    // rgpszMkDocuments: Array of file paths removed.
    // rgFlags: Array of flags associated with each file.
    virtual HRESULT __stdcall OnAfterRemoveFiles(
      int cProjects,
      int cFiles,
      IVsProject *rgpProjects[],
      const int rgFirstIndices[],
      const LPCOLESTR rgpszMkDocuments[],
      const VSREMOVEFILEFLAGS rgFlags[]
    ) = 0;

    // Notifies the client after directories have been removed from the project.
    // cProjects: Number of projects involved.
    // cDirectories: Number of directories removed.
    // rgpProjects: Array of projects involved.
    // rgFirstIndices: Array of indices identifying which project each directory belonged to.
    // rgpszMkDocuments: Array of directory paths removed.
    // rgFlags: Array of flags associated with each directory.
    virtual HRESULT __stdcall OnAfterRemoveDirectories(
      int cProjects,
      int cDirectories,
      IVsProject *rgpProjects[],
      const int rgFirstIndices[],
      const LPCOLESTR rgpszMkDocuments[],
      const VSREMOVEDIRECTORYFLAGS rgFlags[]
    ) = 0;

    // Notifies the client when a project has requested to rename files.
    // pProject: Project requesting to rename files.
    // cFiles: Number of files to rename.
    // rgszMkOldNames: Array of current file paths.
    // rgszMkNewNames: Array of new file paths.
    // rgflags: Array of flags associated with each file.
    // pSummaryResult: Summary result object indicating if the rename operation can proceed.
    // rgResults: Array of results for each individual file.
    virtual HRESULT __stdcall OnQueryRenameFiles(
      IVsProject *pProject,
      int cFiles,
      const LPCOLESTR rgszMkOldNames[],
      const LPCOLESTR rgszMkNewNames[],
      const VSQUERYRENAMEFILEFLAGS rgflags[],
      VSQUERYRENAMEFILERESULTS *pSummaryResult,
      VSQUERYRENAMEFILERESULTS rgResults[]
    ) = 0;

    // Notifies the client after files have been renamed in the project.
    // cProjects: Number of projects involved.
    // cFiles: Number of files renamed.
    // rgpProjects: Array of projects involved.
    // rgFirstIndices: Array of indices identifying which project each file belongs to.
    // rgszMkOldNames: Array of old file paths.
    // rgszMkNewNames: Array of new file paths.
    // rgflags: Array of flags associated with each file.
    virtual HRESULT __stdcall OnAfterRenameFiles(
      int cProjects,
      int cFiles,
      IVsProject *rgpProjects[],
      const int rgFirstIndices[],
      const LPCOLESTR rgszMkOldNames[],
      const LPCOLESTR rgszMkNewNames[],
      const VSRENAMEFILEFLAGS rgflags[]
    ) = 0;

    // Notifies the client when a project has requested to rename directories.
    // pProject: Project requesting to rename directories.
    // cDirs: Number of directories to rename.
    // rgszMkOldNames: Array of current directory paths.
    // rgszMkNewNames: Array of new directory paths.
    // rgflags: Array of flags associated with each directory.
    // pSummaryResult: Summary result object indicating if the rename operation can proceed.
    // rgResults: Array of results for each individual directory.
    virtual HRESULT __stdcall OnQueryRenameDirectories(
      IVsProject *pProject,
      int cDirs,
      const LPCOLESTR rgszMkOldNames[],
      const LPCOLESTR rgszMkNewNames[],
      const VSQUERYRENAMEDIRECTORYFLAGS rgflags[],
      VSQUERYRENAMEDIRECTORYRESULTS *pSummaryResult,
      VSQUERYRENAMEDIRECTORYRESULTS rgResults[]
    ) = 0;

    // Notifies the client after directories have been renamed in the project.
    // cProjects: Number of projects involved.
    // cDirs: Number of directories renamed.
    // rgpProjects: Array of projects involved.
    // rgFirstIndices: Array of indices identifying which project each directory belongs to.
    // rgszMkOldNames: Array of old directory paths.
    // rgszMkNewNames: Array of new directory paths.
    // rgflags: Array of flags associated with each directory.
    virtual HRESULT __stdcall OnAfterRenameDirectories(
      int cProjects,
      int cDirs,
      IVsProject *rgpProjects[],
      const int rgFirstIndices[],
      const LPCOLESTR rgszMkOldNames[],
      const LPCOLESTR rgszMkNewNames[],
      const VSRENAMEDIRECTORYFLAGS rgflags[]
    ) = 0;

    // Notifies the client when a project has requested to add directories.
    // pProject: Project requesting to add directories.
    // cDirectories: Number of directories to add.
    // rgpszMkDocuments: Array of directory paths to add to the project.
    // rgFlags: Array of flags associated with each directory.
    // pSummaryResult: Summary result object indicating if the add operation can proceed.
    // rgResults: Array of results for each individual directory.
    virtual HRESULT __stdcall OnQueryAddDirectories(
      IVsProject *pProject,
      int cDirectories,
      const LPCOLESTR rgpszMkDocuments[],
      const VSQUERYADDDIRECTORYFLAGS rgFlags[],
      VSQUERYADDDIRECTORYRESULTS *pSummaryResult,
      VSQUERYADDDIRECTORYRESULTS rgResults[]
    ) = 0;

    // Notifies the client when a project has requested to remove files.
    // pProject: Project requesting to remove files.
    // cFiles: Number of files to remove.
    // rgpszMkDocuments: Array of file paths to remove from the project.
    // rgFlags: Array of flags associated with each file.
    // pSummaryResult: Summary result object indicating if the remove operation can proceed.
    // rgResults: Array of results for each individual file.
    virtual HRESULT __stdcall OnQueryRemoveFiles(
      IVsProject *pProject,
      int cFiles,
      const LPCOLESTR rgpszMkDocuments[],
      const VSQUERYREMOVEFILEFLAGS rgFlags[],
      VSQUERYREMOVEFILERESULTS *pSummaryResult,
      VSQUERYREMOVEFILERESULTS rgResults[]
    ) = 0;

    // Notifies the client when a project has requested to remove directories.
    // pProject: Project requesting to remove directories.
    // cDirectories: Number of directories to remove.
    // rgpszMkDocuments: Array of directory paths to remove from the project.
    // rgFlags: Array of flags associated with each directory.
    // pSummaryResult: Summary result object indicating if the remove operation can proceed.
    // rgResults: Array of results for each individual directory.
    virtual HRESULT __stdcall OnQueryRemoveDirectories(
      IVsProject *pProject,
      int cDirectories,
      const LPCOLESTR rgpszMkDocuments[],
      const VSQUERYREMOVEDIRECTORYFLAGS rgFlags[],
      VSQUERYREMOVEDIRECTORYRESULTS *pSummaryResult,
      VSQUERYREMOVEDIRECTORYRESULTS rgResults[]
    ) = 0;

    // Notifies the client when the source control status of files has changed.
    // cProjects: Number of projects involved.
    // cFiles: Number of files whose status has changed.
    // rgpProjects: Array of projects involved.
    // rgFirstIndices: Array of indices identifying which project each file belongs to.
    // rgpszMkDocuments: Array of file paths.
    // rgdwSccStatus: Array of source control status values for each file.
    virtual HRESULT __stdcall OnAfterSccStatusChanged(
      int cProjects,
      int cFiles,
      IVsProject *rgpProjects[],
      const int rgFirstIndices[],
      const LPCOLESTR rgpszMkDocuments[],
      const DWORD rgdwSccStatus[]
    ) = 0;
};
#pragma endregion

  // ================================ IVsTrackProjectDocumentsEvents3 ================================
#pragma region
  const GUID IID_IVsTrackProjectDocumentsEvents3 = { 0x53544c4d, 0xBD74, 0x4D21, { 0xA7, 0x9F, 0x2C, 0x19, 0x0E, 0x38, 0xAB, 0x6F } };

  enum HANDSOFFMODE {
    HANDSOFFMODE_ReadAccess = 0x1,
    HANDSOFFMODE_WriteAccess = 0x2,
    HANDSOFFMODE_DeleteAccess = 0x4,
    HANDSOFFMODE_AsyncOperation = 0x80000000,
    HANDSOFFMODE_FullAccess = ((HANDSOFFMODE_DeleteAccess | HANDSOFFMODE_ReadAccess) | HANDSOFFMODE_WriteAccess),
    HANDSOFFMODE_ReadWriteAccess = (HANDSOFFMODE_ReadAccess | HANDSOFFMODE_WriteAccess)
  };

  // Extends IVsTrackProjectDocumentsEvents2 to provide additional notifications for batch operations and file access coordination.
  struct __declspec(novtable) IVsTrackProjectDocumentsEvents3 {
    // == IUnknown methods ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTrackProjectDocumentsEvents3 methods ==

    // Indicates that a project is about to start a batch query process.
    virtual HRESULT __stdcall OnBeginQueryBatch() = 0;

    // Determines whether it is okay to proceed with the actual batch operation after successful completion of a batch query process.
    // pfActionOK: [out] Returns nonzero if it is okay to continue with the proposed batch process; returns zero if the proposed batch process should not proceed.
    virtual HRESULT __stdcall OnEndQueryBatch(BOOL *pfActionOK) = 0;

    // Indicates that a batch query process has been canceled.
    virtual HRESULT __stdcall OnCancelQueryBatch() = 0;

    // Determines if it is okay to add a collection of files (possibly from source control) whose final destination may be different from a source location.
    // pProject: [in] Pointer to the project.
    // cFiles: [in] Number of files to be added.
    // rgpszNewMkDocuments: [in] Array of new file paths.
    // rgpszSrcMkDocuments: [in] Array of source file paths.
    // rgFlags: [in] Array of flags indicating the add operation.
    // pSummaryResult: [out] Summary result of the query.
    // rgResults: [out] Array of results for each file.
    virtual HRESULT __stdcall OnQueryAddFilesEx(IVsProject *pProject, int cFiles, const LPCOLESTR rgpszNewMkDocuments[], const LPCOLESTR rgpszSrcMkDocuments[], const VSQUERYADDFILEFLAGS rgFlags[], VSQUERYADDFILERESULTS *pSummaryResult, VSQUERYADDFILERESULTS rgResults[]) = 0;

    // Accesses a specified set of files and asks all implementers of this method to release any locks that may exist on those files.
    // grfRequiredAccess: [in] Specifies the required access mode.
    // cFiles: [in] Number of files.
    // rgpszMkDocuments: [in] Array of file paths.
    virtual HRESULT __stdcall HandsOffFiles(HANDSOFFMODE grfRequiredAccess, int cFiles, const LPCOLESTR rgpszMkDocuments[]) = 0;

    // Called when a project has completed operations on a set of files.
    // cFiles: [in] Number of files.
    // rgpszMkDocuments: [in] Array of file paths.
    virtual HRESULT __stdcall HandsOnFiles(int cFiles, const LPCOLESTR rgpszMkDocuments[]) = 0;
  };
#pragma endregion

  // ================================ IVsTrackProjectDocumentsEvents4 ================================
#pragma region
  const GUID IID_IVsTrackProjectDocumentsEvents4 = { 0x53544c4d, 0xDB74, 0x4078, { 0x96, 0x42, 0x5D, 0x1B, 0xC0, 0xBB, 0x5B, 0x26 } };

  // Notifies clients of additional changes made to project files or directories.
  // Extends IVsTrackProjectDocumentsEvents2 to provide additional notifications for project document events.
  // Implemented by clients that try to keep track of changes to the contents of projects in the Solution.
  struct __declspec(novtable) IVsTrackProjectDocumentsEvents4 {
    // == IUnknown methods ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTrackProjectDocumentsEvents4 methods ==

    // Called to determine whether files can be removed from the project.
    // pProject: [in] Project from which the files will be removed.
    // cFiles: [in] Number of files to be removed.
    // rgpszMkDocuments: [in, size_is(cFiles)] Array of paths for the files to be removed.
    // rgFlags: [in, size_is(cFiles)] Array of flags. For a list of rgFlags values, see VSQUERYREMOVEFILEFLAGS2.
    // pSummaryResult: [out] Summary result object. This object is a summation of the yes and no results for the array of files passed in rgpszMkDocuments. 
    //   If the result for a single file is no, then this parameter is equal to VSQUERYREMOVEFILERESULTS_RemoveNotOK; if the results for all files are yes, 
    //   then this parameter is equal to VSQUERYREMOVEFILERESULTS_RemoveOK. For a list of pSummaryResult values, see VSQUERYREMOVEFILERESULTS.
    // rgResults: [out, size_is(cFiles)] Array of results. For a list of rgResults values, see VSQUERYREMOVEFILERESULTS.
    virtual HRESULT __stdcall OnQueryRemoveFilesEx(
      IVsProject *pProject,
      int cFiles,
      const LPCOLESTR rgpszMkDocuments[],
      const VSQUERYREMOVEFILEFLAGS2 rgFlags[],
      VSQUERYREMOVEFILERESULTS *pSummaryResult,
      VSQUERYREMOVEFILERESULTS rgResults[]
    ) = 0;

    // Called to determine whether directories can be removed from the project.
    // pProject: [in] Project from which the directories will be removed.
    // cDirectories: [in] Number of directories to be removed.
    // rgpszMkDocuments: [in, size_is(cDirectories)] Array of paths for the directories to remove.
    // rgFlags: [in, size_is(cDirectories)] Array of flags. For a list of rgFlags values, see VSQUERYREMOVEDIRECTORYFLAGS2.
    // pSummaryResult: [out] Summary result object. This object is a summation of the yes and no results for the array of directories passed in rgpszMkDocuments. 
    //   If the result for a single directory is no, then this parameter is equal to VSQUERYREMOVEDIRECTORYRESULTS_RemoveNotOK; if the results for all files are yes, 
    //   then this parameter is equal to VSQUERYREMOVEDIRECTORYRESULTS_RemoveOK. For a list of pSummaryResult values, see VSQUERYREMOVEDIRECTORYRESULTS.
    // rgResults: [out, size_is(cDirectories)] Array of results. For a list of rgResults values, see VSQUERYREMOVEDIRECTORYRESULTS.
    virtual HRESULT __stdcall OnQueryRemoveDirectoriesEx(
      IVsProject *pProject,
      int cDirectories,
      const LPCOLESTR rgpszMkDocuments[],
      const VSQUERYREMOVEDIRECTORYFLAGS2 rgFlags[],
      VSQUERYREMOVEDIRECTORYRESULTS *pSummaryResult,
      VSQUERYREMOVEDIRECTORYRESULTS rgResults[]
    ) = 0;

    // Notifies that files have been removed from the project.
    // cProjects: [in] Number of projects from which files were removed.
    // cFiles: [in] Number of files removed.
    // rgpProjects: [in, size_is(cProjects)] Array of projects from which files were removed.
    // rgFirstIndices: [in, size_is(cProjects)] Array of first indices identifying to which project each file belongs. For more information, see IVsTrackProjectDocumentsEvents2.
    // rgpszMkDocuments: [in, size_is(cFiles)] Array of paths for the files that were removed.
    // rgFlags: [in, size_is(cFiles)] Array of flags. For a list of rgFlags values, see VSREMOVEFILEFLAGS2.
    virtual HRESULT __stdcall OnAfterRemoveFilesEx(
      int cProjects,
      int cFiles,
      IVsProject *rgpProjects[],
      const int rgFirstIndices[],
      const LPCOLESTR rgpszMkDocuments[],
      const VSREMOVEFILEFLAGS2 rgFlags[]
    ) = 0;

    // Notifies that directories have been removed from the project.
    // cProjects: [in] Number of projects from which directories were removed.
    // cDirectories: [in] Number of directories removed.
    // rgpProjects: [in, size_is(cProjects)] Array of projects from which directories were removed.
    // rgFirstIndices: [in, size_is(cProjects)] Array of first indices identifying to which project each directory belongs. For more information, see IVsTrackProjectDocumentsEvents2.
    // rgpszMkDocuments: [in, size_is(cDirectories)] Array of paths for the directories that were removed.
    // rgFlags: [in, size_is(cDirectories)] Array of flags. For a list of rgFlags values, see VSREMOVEDIRECTORYFLAGS2.
    virtual HRESULT __stdcall OnAfterRemoveDirectoriesEx(
      int cProjects,
      int cDirectories,
      IVsProject *rgpProjects[],
      const int rgFirstIndices[],
      const LPCOLESTR rgpszMkDocuments[],
      const VSREMOVEDIRECTORYFLAGS2 rgFlags[]
    ) = 0;
  };
#pragma endregion

  // ================================ IVsTrackProjectDocuments2 ================================
#pragma region
  const GUID IID_IVsTrackProjectDocuments2 = { 0x53544C4D, 0x6639, 0x11d3, { 0xa6, 0x0d, 0x00, 0x50, 0x04, 0x77, 0x5a, 0xb1 } };

  // Interface to broadcast document events
  // The point of the SVsTrackProjectDocuments is to bottleneck certain shell events. Packages can Advise with the service at SetSite time.
  // Projects use this service to signal that they want to receive document events, and to signal events that are about to happen.
  // Before adding, renaming, or deleting a file or directory, the project must call the appropriate TrackProjectDocuments2 method in order to check whether this operation is allowed.
  // When the operation is completed, the project must notify the TrackProjectDocuments2 service, so that appropriate notifications can be sent out.
  // You need to make these calls even if your project does not support source control.
  // 
  // You must use IVsTrackProjectDocuments2 for all files, not just master files.
  // For example, if you have a form with a.resx file and other files, you need to tell the environment about all of the files.
  // Do not call the methods of IVsTrackProjectDocuments2 at project open or close.
  // Any entity that requires the information provided through IVsTrackProjectDocuments2 at startup can wait for the OnAfterOpenSolution event and iterate through the solution to find the information required.
  // On shutdown, this information is not needed.
  // 
  // Access to IVsTrackProjectDocuments2 is provided from the SVsTrackProjectDocuments service.
  // 
  // For each call on IVsTrackProjectDocuments2, there are two methods, the OnQuery *method and the OnAfter *method.
  // Call the appropriate OnQuery *method to request permission to add, remove, or rename a file or directory in a project.
  // From this call, you might receive notification that the operation cannot proceed.
  // For example, if the Enterprise Framework and Template(EFT) project system does not permit the user to add a file that does not meet policy, the project must be prepared to not perform the add, remove, or rename.
  // If permission is granted, the project must complete the add, rename, or remove action and then call the appropriate OnAfter* method to inform the environment of the changes made to the project.
  // The IVsTrackProjectDocuments2 method also applies to directories, but directory calls are optional.
  // If your project system has directory information, then provide this information to the environment using these methods.
  // However, if the project system does not have this information, then the environment will infer it.
  //
  // All directory calls are optional.However, if you call one of the OnQuery *directory methods and the call was successful, then you are required to call the corresponding OnAfter *directory method.
  //
  // Implemented by the environment. This interface is the mechanism for gathering the information regarding when a file or directory is added, removed, or renamed in a project.
  // Called by projects to query the environment as to whether a file or directory can be added, removed, or renamed in a solution.
  // For all actions approved by the environment, the appropriate method is then called after the action is completed.
  // IVsTrackProjectDocuments2 must be used by all projects, regardless of whether they support source control.
  struct __declspec(novtable) IVsTrackProjectDocuments2 {
    // == IUnknown methods ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsTrackProjectDocuments2 methods ==

    // Begins a batch of OnAfter* method calls.
    // The IVsTrackProjectDocuments2 interface allows projects to tell the environment when a project item has been added, removed, or renamed.
    // During these operations, user interface(UI) is sometimes displayed.
    // IVsTrackProjectDocuments2.BeginBatch informs the environment that you are going to call several IVsTrackProjectDocuments2.
    // OnAfter *methods, and that you would like the user to receive only one UI display for those calls.
    // Grouping these calls using IVsTrackProjectDocuments2.BeginBatch and EndBatch increases the likelihood that the environment will display only one UI display; however, this is not guaranteed.
    //
    // Call IVsTrackProjectDocuments2.BeginBatch to start the batch, make multiple IVsTrackProjectDocuments2 calls, and then call EndBatch to display the UI.
    //
    // You can batch only OnAfter *methods.OnQuery *methods cannot be batched.
    virtual HRESULT __stdcall BeginBatch() = 0;

    // Ends the batch started by BeginBatch and displays any UI generated within the batch.
    virtual HRESULT __stdcall EndBatch() = 0;

    // Displays the UI for the calls completed so far without ending the batch.
    virtual HRESULT __stdcall Flush() = 0;

    // Determines whether files can be added to the project.
    // pProject: Pointer to the IVsProject interface.
    // cFiles: Number of files.
    // rgpszMkDocuments: Array of file paths.
    // rgFlags: Array of flags for each file.
    // pSummaryResult: Pointer to the summary result.
    // rgResults: Array of results for each file.
    virtual HRESULT __stdcall OnQueryAddFiles(
      IVsProject *pProject,
      int cFiles,
      const LPCOLESTR rgpszMkDocuments[],
      const VSQUERYADDFILEFLAGS rgFlags[],
      VSQUERYADDFILERESULTS *pSummaryResult,
      VSQUERYADDFILERESULTS rgResults[]
    ) = 0;

    // Called after files have been added to the project.
    // pProject: Pointer to the IVsProject interface.
    // cFiles: Number of files.
    // rgpszMkDocuments: Array of file paths.
    // rgFlags: Array of flags for each file.
    virtual HRESULT __stdcall OnAfterAddFilesEx(
      IVsProject *pProject,
      int cFiles,
      const LPCOLESTR rgpszMkDocuments[],
      const VSADDFILEFLAGS rgFlags[]
    ) = 0;

    // Called after files have been added to the project.
    // pProject: Pointer to the IVsProject interface.
    // cFiles: Number of files.
    // rgpszMkDocuments: Array of file paths.
    virtual HRESULT __stdcall OnAfterAddFiles(
      IVsProject *pProject,
      int cFiles,
      const LPCOLESTR rgpszMkDocuments[]
    ) = 0;

    // Called after directories have been added to the project.
    // pProject: Pointer to the IVsProject interface.
    // cDirectories: Number of directories.
    // rgpszMkDocuments: Array of directory paths.
    // rgFlags: Array of flags for each directory.
    virtual HRESULT __stdcall OnAfterAddDirectoriesEx(
      IVsProject *pProject,
      int cDirectories,
      const LPCOLESTR rgpszMkDocuments[],
      const VSADDDIRECTORYFLAGS rgFlags[]
    ) = 0;

    // Called after directories have been added to the project.
    // pProject: Pointer to the IVsProject interface.
    // cDirectories: Number of directories.
    // rgpszMkDocuments: Array of directory paths.
    virtual HRESULT __stdcall OnAfterAddDirectories(
      IVsProject *pProject,
      int cDirectories,
      const LPCOLESTR rgpszMkDocuments[]
    ) = 0;

    // Called after files have been removed from the project.
    // pProject: Pointer to the IVsProject interface.
    // cFiles: Number of files.
    // rgpszMkDocuments: Array of file paths.
    // rgFlags: Array of flags for each file.
    virtual HRESULT __stdcall OnAfterRemoveFiles(
      IVsProject *pProject,
      int cFiles,
      const LPCOLESTR rgpszMkDocuments[],
      const VSREMOVEFILEFLAGS rgFlags[]
    ) = 0;

    // Called after directories have been removed from the project.
    // pProject: Pointer to the IVsProject interface.
    // cDirectories: Number of directories.
    // rgpszMkDocuments: Array of directory paths.
    // rgFlags: Array of flags for each directory.
    virtual HRESULT __stdcall OnAfterRemoveDirectories(
      IVsProject *pProject,
      int cDirectories,
      const LPCOLESTR rgpszMkDocuments[],
      const VSREMOVEDIRECTORYFLAGS rgFlags[]
    ) = 0;

    // Determines whether files can be renamed in the project.
    // pProject: Pointer to the IVsProject interface.
    // cFiles: Number of files.
    // rgszMkOldNames: Array of old file paths.
    // rgszMkNewNames: Array of new file paths.
    // rgflags: Array of flags for each file.
    // pSummaryResult: Pointer to the summary result.
    // rgResults: Array of results for each file.
    virtual HRESULT __stdcall OnQueryRenameFiles(
      IVsProject *pProject,
      int cFiles,
      const LPCOLESTR rgszMkOldNames[],
      const LPCOLESTR rgszMkNewNames[],
      const VSQUERYRENAMEFILEFLAGS rgflags[],
      VSQUERYRENAMEFILERESULTS *pSummaryResult,
      VSQUERYRENAMEFILERESULTS rgResults[]
    ) = 0;

    // Determines whether a file can be renamed in the project.
    // pProject: Pointer to the IVsProject interface.
    // pszMkOldName: Old file path.
    // pszMkNewName: New file path.
    // flags: Flags for the file.
    // pfRenameCanContinue: Pointer to the result indicating if rename can continue.
    virtual HRESULT __stdcall OnQueryRenameFile(
      IVsProject *pProject,
      LPCOLESTR pszMkOldName,
      LPCOLESTR pszMkNewName,
      VSRENAMEFILEFLAGS flags,
      BOOL *pfRenameCanContinue
    ) = 0;

    // Called after files have been renamed in the project.
    // pProject: Pointer to the IVsProject interface.
    // cFiles: Number of files.
    // rgszMkOldNames: Array of old file paths.
    // rgszMkNewNames: Array of new file paths.
    // rgflags: Array of flags for each file.
    virtual HRESULT __stdcall OnAfterRenameFiles(
      IVsProject *pProject,
      int cFiles,
      const LPCOLESTR rgszMkOldNames[],
      const LPCOLESTR rgszMkNewNames[],
      const VSRENAMEFILEFLAGS rgflags[]
    ) = 0;

    // rgflags: Flags for the file.
    virtual HRESULT __stdcall OnAfterRenameFile(
      IVsProject *pProject,
      LPCOLESTR pszMkOldName,
      LPCOLESTR pszMkNewName,
      VSRENAMEFILEFLAGS flags
    ) = 0;

    // Determines whether directories can be renamed in the project.
    // pProject: Pointer to the IVsProject interface.
    // cDirs: Number of directories.
    // rgszMkOldNames: Array of old directory paths.
    // rgszMkNewNames: Array of new directory paths.
    // rgflags: Flags for each directory.
    // pSummaryResult: Pointer to the summary result.
    // rgResults: Array of results for each directory.
    virtual HRESULT __stdcall OnQueryRenameDirectories(
      IVsProject *pProject,
      int cDirs,
      const LPCOLESTR rgszMkOldNames[],
      const LPCOLESTR rgszMkNewNames[],
      const VSQUERYRENAMEDIRECTORYFLAGS rgflags[],
      VSQUERYRENAMEDIRECTORYRESULTS *pSummaryResult,
      VSQUERYRENAMEDIRECTORYRESULTS rgResults[]
    ) = 0;

    // Called after directories have been renamed in the project.
    // pProject: Pointer to the IVsProject interface.
    // cDirs: Number of directories.
    // rgszMkOldNames: Array of old directory paths.
    // rgszMkNewNames: Array of new directory paths.
    // rgflags: Flags for each directory.
    virtual HRESULT __stdcall OnAfterRenameDirectories(
      IVsProject *pProject,
      int cDirs,
      const LPCOLESTR rgszMkOldNames[],
      const LPCOLESTR rgszMkNewNames[],
      const VSRENAMEDIRECTORYFLAGS rgflags[]
    ) = 0;

    // Advises the client of changes to project documents.
    // pEventSink: Pointer to the event sink interface.
    // pdwCookie: Pointer to the event cookie.
    virtual HRESULT __stdcall AdviseTrackProjectDocumentsEvents(
      IVsTrackProjectDocumentsEvents2 *pEventSink,
      DWORD *pdwCookie
    ) = 0;

    // Unadvises the client of changes to project documents.
    // dwCookie: The event cookie to unadvise.
    virtual HRESULT __stdcall UnadviseTrackProjectDocumentsEvents(
      DWORD dwCookie
    ) = 0;

    // Determines whether directories can be added to the project.
    // pProject: Pointer to the IVsProject interface.
    // cDirectories: Number of directories.
    // rgpszMkDocuments: Array of directory paths.
    // rgFlags: Flags for each directory.
    // pSummaryResult: Pointer to the summary result.
    // rgResults: Array of results for each directory.
    virtual HRESULT __stdcall OnQueryAddDirectories(
      IVsProject *pProject,
      int cDirectories,
      const LPCOLESTR rgpszMkDocuments[],
      const VSQUERYADDDIRECTORYFLAGS rgFlags[],
      VSQUERYADDDIRECTORYRESULTS *pSummaryResult,
      VSQUERYADDDIRECTORYRESULTS rgResults[]
    ) = 0;

    // Determines whether files can be removed from the project.
    // pProject: Pointer to the IVsProject interface.
    // cFiles: Number of files.
    // rgpszMkDocuments: Array of file paths.
    // rgFlags: Flags for each file.
    // pSummaryResult: Pointer to the summary result.
    // rgResults: Array of results for each file.
    virtual HRESULT __stdcall OnQueryRemoveFiles(
      IVsProject *pProject,
      int cFiles,
      const LPCOLESTR rgpszMkDocuments[],
      const VSQUERYREMOVEFILEFLAGS rgFlags[],
      VSQUERYREMOVEFILERESULTS *pSummaryResult,
      VSQUERYREMOVEFILERESULTS rgResults[]
    ) = 0;

    // Determines whether directories can be removed from the project.
    // pProject: Pointer to the IVsProject interface.
    // cDirectories: Number of directories.
    // rgpszMkDocuments: Array of directory paths.
    // rgFlags: Flags for each directory.
    // pSummaryResult: Pointer to the summary result.
    // rgResults: Array of results for each directory.
    virtual HRESULT __stdcall OnQueryRemoveDirectories(
      IVsProject *pProject,
      int cDirectories,
      const LPCOLESTR rgpszMkDocuments[],
      const VSQUERYREMOVEDIRECTORYFLAGS rgFlags[],
      VSQUERYREMOVEDIRECTORYRESULTS *pSummaryResult,
      VSQUERYREMOVEDIRECTORYRESULTS rgResults[]
    ) = 0;

    // Called after source control status has changed for a list of files.
    // pProject: Pointer to the IVsProject interface.
    // cFiles: Number of files.
    // rgpszMkDocuments: Array of file paths.
    // rgdwSccStatus: Array of source control status values for each file.
    virtual HRESULT __stdcall OnAfterSccStatusChanged(
      IVsProject *pProject,
      int cFiles,
      const LPCOLESTR rgpszMkDocuments[],
      const DWORD rgdwSccStatus[]
    ) = 0;
};
#pragma endregion

  // ================================ IVsTrackProjectDocuments3 ================================
#pragma region
const GUID IID_IVsTrackProjectDocuments3 = { 0x53544c4d, 0x9097, 0x4325, { 0x92, 0x70, 0x75, 0x4e, 0xb8, 0x5a, 0x63, 0x51 } };

// Provides batch processing and coordination of file access for project documents.
struct __declspec(novtable) IVsTrackProjectDocuments3 {
  // == IUnknown methods ==
  virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
  virtual ULONG __stdcall AddRef() = 0;
  virtual ULONG __stdcall Release() = 0;

  // == IVsTrackProjectDocuments3 methods ==

  // These events enable batch handling of a group of OnQueryRenameXXX, 
  // OnQueryRemoveXXX, or OnQueryAddXXX events. For example, this batching
  // can be used by a SCC provider to offer a single checkout prompt 
  // dialog even when a subfolder tree of files is moved or deleted.
  // A project system must guarantee to balance each call to BeginQueryBatch
  // with either a call to EndQueryBatch or CancelQueryBatch. CancelQueryBatch
  // should be called if one of the individual OnQuery calls returned NotOK.
  // EndQueryBatch should be called if none of the individual OnQuery calls
  // were vetoed. This call to EndQueryBatch will return a final OK/NotOK
  // decision. The call pattern of these methods should look something like 
  // the following:
  //		BeginQueryBatch();
  //		QueryFolderRename(pFolderNode, strOldPath, strNewPath, &fRenameCanCont);
  //		if (fRenameCanCont)
  //			EndQueryBatch(&fRenameCanCont);
  //		else
  //			CancelQueryBatch();
  //
  // where assume QueryFolderRename is an internal project helper function
  // that causes a series of item level OnQueryRenameFile events to be fired.
  // Note: Batching nests.

  // Starts a batch query process, allowing multiple file operations to be confirmed in a single user prompt.
  virtual HRESULT __stdcall BeginQueryBatch() = 0;

  // Ends the batch query process and determines whether the batched operations should proceed.
  // pfActionOK: [out] Pointer to a BOOL that receives TRUE if the operations can proceed; otherwise, FALSE.
  virtual HRESULT __stdcall EndQueryBatch(BOOL *pfActionOK) = 0;

  // Cancels the batch query process, indicating that the batched operations should not proceed.
  virtual HRESULT __stdcall CancelQueryBatch() = 0;

  // This OnQueryAddFilesEx is an extension of the original OnQueryAddFiles 
  // method which passes both the path to the original source files being
  // added as well as the path of resultant files newly added to the project.
  // This method is usefull for directory based project systems which often
  // make a copy of the file being added to the project to a new location 
  // within the project directory. Clients that manage a coordinated set 
  // of files can use this information to find the other related source files 
  // that may need to be added as well (e.g. form/report files with 
  // associated code behind/resource files).
  // 
  // Determines if it is permissible to add a collection of files, possibly from source control, to the project.
  // pProject: Pointer to the project making the request.
  // cFiles: The number of files to be added.
  // rgpszNewMkDocuments: Array of file paths indicating the final destination of the files.
  // rgpszSrcMkDocuments: Array of file paths specifying the source location of the files.
  // rgFlags: Array of flags from the VSQUERYADDFILEFLAGS enumeration, one for each file.
  // pSummaryResult: [out] Pointer to a VSQUERYADDFILERESULTS value indicating the overall result.
  // rgResults: [out] Array that receives the result for each file.
  virtual HRESULT __stdcall OnQueryAddFilesEx(
    IVsProject *pProject,
    int cFiles,
    const LPCOLESTR rgpszNewMkDocuments[],
    const LPCOLESTR rgpszSrcMkDocuments[],
    const VSQUERYADDFILEFLAGS rgFlags[],
    VSQUERYADDFILERESULTS *pSummaryResult,
    VSQUERYADDFILERESULTS rgResults[]
  ) = 0;

  // Hands[Off/On]Files is called to tell all parties that may hold files locked
  // to release the files so that they may be operated on by the caller. These
  // methods can be called with files that are members of projects as well as with
  // files that are not members of projects. HandsOffFiles should be called on any 
  // file that someone needs to read to, write to, rename or delete. The 
  // "grfRequiredAccess" parameter indicates the access permissions required by the 
  // caller. If a party holding a lock on a file allows the requested permissions, 
  // then no releasing of their lock is required. Multiple parties are allowed to
  // hold ReadAccess on a file simultaneously, but only one party may hold WriteAccess 
  // or DeleteAccess at a time. A holder of a shared read lock on a file that does
  // not deny writing to the file does not have to release the file if someone 
  // requests WriteAccess, but if the caller requires DeleteAccess in order to rename
  // or delete a file, then all locks must be released. This parameter lets lock holders 
  // optimize how they manage their locks.
  //
  // For example, a proper project system should call this with files it manages as 
  // part of its build operation:
  //     1. it should request HANDSOFFMODE_ReadAccess for all source files
  //     2. it should request HANDSOFFMODE_FullAccess for all output files
  // This would tell all parties to release the files that the project wants to 
  // access as part of a build operation including anyone that would block the
  // project from reading its source files or regenerating its output files. 
  // Similarly, the Source Code Control (SCC) integration VSPackage will call this
  // prior to retrieving new copies of files from the SCC server. A well behaved
  // project system would call this for all save, copy, rename, move, and delete 
  // operations. HandsOffFiles should be called after all calls to OnQueryXXX have 
  // returned that the operation is OK to proceed.
  //
  // HandsOnFiles will not always be called, e.g. in scenarios where files are
  // deleted, renamed or moved it will not be called. Normally HandsOnFiles will
  // only be called in scenarios where contents of files are modified in place or
  // contents of files are read or copied. HandsOnFiles must always be called
  // to terminate an Async operation.
  //
  // Sync vs. Async Operations: Normally the request for access to a file is for
  // a synchronous operation that completes before execution returns to the main
  // message loop of the IDE. In some cases access to files is needed for an 
  // asynchronous operation which returns to the main message loop and continues 
  // for an extended period of time. A classic scenario for this is the Solution/Project
  // Build operation. The Build operation is asynchronous; the user may work
  // while the build operation is proceeding. To indicate that HandsOffFiles is 
  // is required for an asynchronous operation, HANDSOFFMODE_AsyncOperation should 
  // be or-ed together with the Access flags required. When this flag is specified,
  // then the caller is making a guarantee that HandsOnFiles will be called to 
  // indicate the end of the Async operation. 
  // For example, well behaved clients that hold build output files open, will let go 
  // of the file(s) when HandsOffFiles is called and track the state that an 
  // AsyncOperation is in effect. They will avoid opening the file(s) again until 
  // HandsOnFiles is called. Otherwise, they would interfere with the Build operation. 
  // They should treat this as an E_ACCESSDENIED situation even though in reality the 
  // file might not be locked in the file system.
  //
  // NOTE: if HandsOffFiles returned an error, then you should not call HandsOnFiles.
  //
  // NOTE: incompatible HANDSOFFMODE_AsyncOperations do not nest. A call to 
  // IVsTrackProjectDocuments3:: HandsOffFiles will return E_ACCESSDENIED if there 
  // already is a pending incompatible AsyncOperation. There can be multiple
  // nested Async ReadAccess operations but WriteAccess and DeleteAccess operations
  // are not allowed to be nested. Thus Async ReadAccess operations must be refcounted.
  //
  // NOTE: The IVsTrackProjectDocuments3:: HandsOffFiles event is cancelable by 
  // returning an error HRESULT (normally E_ACCESSDENIED) and setting a useful error
  // message via a call to ::SetErrorInfo. The E_NOTIMPL error code is treated 
  // as if the event sink does not exist and does not block the HandsOff operation.
  // Any client that refuses to honor the request to HandsOff will end up blocking
  // the desired operation from succeeding. If one event sink returns an error 
  // then all sinks that were previously called for HandsOffFiles will be called with
  // HandsOnFiles.

  // Requests that any locks on the specified files be released to allow for operations such as editing or deletion.
  // grfRequiredAccess: Specifies the required access level, using values from the HANDSOFFMODE enumeration.
  // cFiles: The number of files.
  // rgpszMkDocuments: Array of file paths for which locks should be released.
  virtual HRESULT __stdcall HandsOffFiles(
    HANDSOFFMODE grfRequiredAccess,
    int cFiles,
    const LPCOLESTR rgpszMkDocuments[]
  ) = 0;

  // Indicates that operations on the specified files are complete, allowing any necessary locks to be reinstated.
  // cFiles: The number of files.
  // rgpszMkDocuments: Array of file paths that were previously subject to HandsOffFiles.
  virtual HRESULT __stdcall HandsOnFiles(
    int cFiles,
    const LPCOLESTR rgpszMkDocuments[]
  ) = 0;
};
#pragma endregion

  // ================================ IVsCommandWindowCompletion ================================
#pragma region
  const GUID IID_IVsCommandWindowCompletion = { 0x34896BBB, 0xA3D5, 0x4c80, { 0xBC, 0xCE, 0xE9, 0x27, 0x1B, 0xEE, 0xDC, 0x11 } };

  // Implemented by the environment on the command window tool window to coordinate statement completion.
  // This interface is called by the debugger to coordinate statement completion when the command window is in immediate mode.
  struct __declspec(novtable) IVsCommandWindowCompletion {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsCommandWindowCompletion ==

    // Sets the current context for statement completion. 
    // The debugger calls this every time the current statement changes due to stepping, changing stack frames, hitting a breakpoint, etc. 
    // The command window listens for this event and calls the appropriate method, such as SetCompletionContext on the current language service.
    // 
    // pszFilePath: [in] The path to the file containing the current statement.
    // pBuffer: [in] The text buffer containing the current statement.
    // ptsCurStatement: [in] The current statement’s text span.
    // punkContext: [in] May be used to pass additional context in future. If none is available, NULL is passed.
    //   (e.g. The Command Window will listen for this event and in turn call IVsImmediateStatementCompletion::SetCompletionContext on the current language service.)
    virtual HRESULT __stdcall SetCompletionContext(
      LPCOLESTR pszFilePath,
      IVsTextLines *pBuffer,
      const TextSpan *ptsCurStatement,
      IUnknown *punkContext) = 0;
  };
#pragma endregion

  // ================================ IVsImmediateStatementCompletion ================================
#pragma region
  const GUID IID_IVsImmediateStatementCompletion = { 0x5CE7AE60, 0x7B66, 0x4301, { 0x88, 0x92, 0x90, 0xBC, 0x0B, 0x49, 0xA8, 0x9B } };

  // Implemented by a language service object that supports statement completion and other IntelliSense features in the immediate mode of the Command Window (i.e. when the debugger is in break mode). 
  // Called by the Command Window of the Environment.
  // Accessed by calling QueryInterface on the same object that implements IVsLanguageInfo.
  struct __declspec(novtable) IVsImmediateStatementCompletion {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsImmediateStatementCompletion ==

    // Informs the language service that it must add or remove its IVsTextViewFilter from the command filter chain for the command window's IVsTextView.
    // Called by the Command Window when the debugger sets an active language service (via IVsCommandWindow::SetCurrentLanguageService).
    // fInstall: [in] Flag that determines whether to add (true) or remove (false) the TextViewFilter from the CommandFilter chain for the Command Window's IVsTextView (ie. call IVsTextView::AddCommandFilter or RemoveCommandFilter). 
    // pCmdWinView: [in] The text view containing the command filter chain to be modified.
    // fInitialEnable: [in] Flag that determines whether statement completion should be active upon return from this method. Value is true for active statement completion. If the filter is being removed, this parameter is ignored.
    virtual HRESULT __stdcall InstallStatementCompletion(BOOL fInstall, IVsTextView *pCmdWinView, BOOL fInitialEnable) = 0;

    UNUSED_FUNC(0);
    UNUSED_FUNC(1);
  };
#pragma endregion

  // ================================ IVsImmediateStatementCompletion2 ================================
#pragma region
  const GUID IID_IVsImmediateStatementCompletion2 = { 0x58F03D6E, 0xF781, 0x4e44, { 0xBD, 0x88, 0x3B, 0xDE, 0x81, 0x7C, 0xBD, 0xCD } };

  // This interface is implemented by a language service that supports statement completion and other IntelliSense features in the immediate mode of the command window. This mode occurs when the debugger is in break mode.
  // This interface is accessed via QueryInterface on the same object that implements IVsLanguageInfo.
  // This interface is called by the Command Window of the Environment.
  // This interface was created when Statement Completion support was added to the Watch Window, and the support for multiple filters in the language was required.
  // Only use this interface from now on, do not use the old interface.
  struct __declspec(novtable) IVsImmediateStatementCompletion2 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsImmediateStatementCompletion ==

    // Informs the language service that it must add or remove its IVsTextViewFilter from the command filter chain for the command window's IVsTextView.
    // Called by the Command Window when the debugger sets an active language service (via IVsCommandWindow::SetCurrentLanguageService).
    // fInstall: [in] Flag that determines whether to add (true) or remove (false) the TextViewFilter from the CommandFilter chain for the Command Window's IVsTextView (ie. call IVsTextView::AddCommandFilter or RemoveCommandFilter). 
    // pCmdWinView: [in] The text view containing the command filter chain to be modified.
    // fInitialEnable: [in] Flag that determines whether statement completion should be active upon return from this method. Value is true for active statement completion. If the filter is being removed, this parameter is ignored.
    virtual HRESULT __stdcall InstallStatementCompletion(BOOL fInstall, IVsTextView *pCmdWinView, BOOL fInitialEnable) = 0;

    UNUSED_FUNC(0);
    UNUSED_FUNC(1);

    // == IVsImmediateStatementCompletion2 ==

    // Enables or disables statement completion.
    // Called by the Command Window when user actions require statement completion to be turned on or off. 
    // The language service should not install or remove its filter from the filter chain. 
    // If statement completion is being enabled, the iStartIndex and iEndIndex instruct the language service to consider only the indicated part of the current line for statement completion purposes.
    // If iEndIndex is -1, it indicates that the rest of the line is to be used.
    // 
    // fEnable: [in] Flag indicating whether to enable statement completion. True indicates statement completion is enabled.
    // iStartIndex: [in] If fEnable is true, the index in the current line which marks the start of the portion to be used for statement completion. Otherwise ignored.
    // iEndIndex: [in] If fEnable is true, the index in the current line which marks the end of the portion to be used for statement completion. If value is -1, it indicates that the rest of the line is to be used. Ignored on disable of statement completion.
    // pTextView: [in] The text view.
    virtual HRESULT __stdcall EnableStatementCompletion(BOOL fEnable, long iStartIndex, long iEndIndex, IVsTextView *pTextView) = 0;

    // Sets the current context for statement completion for the Command Window.
    // The Command Window calls this method to forward the information that the debugger passes via IVsCommandWindowCompletion::SetCompletionContext.
    // 
    // pszFilePath: [in] The path to the file containing the current statement.
    // pBuffer: [in] The text buffer containing the current statement.
    // ptsCurStatement: [in] The current statement’s text span.
    // punkContext: [in] may be used to pass additional context in future. If none is available, NULL is passed.
    // pTextView: [in] The text view.
    virtual HRESULT __stdcall SetCompletionContext(LPCOLESTR pszFilePath, IVsTextLines *pBuffer, const TextSpan *ptsCurStatement, IUnknown *punkContext, IVsTextView *pTextView) = 0;

    // Retrieves the text view filter associated with the specified text view.
    // This is useful when an operation needs to be performed on the filter, such as GetWordExtent.
    // 
    // pTextView: The text view.
    // ppFilter: [out] Pointer to the text view filter.
    virtual HRESULT __stdcall GetFilter(IVsTextView *pTextView, IVsTextViewFilter **ppFilter) = 0;
  };
#pragma endregion

  // ================================ IVsUIContextManager ================================
#pragma region
  const GUID IID_IVsUIContextManager = { 0xeceae828, 0x2b6f, 0x48ad, { 0xbe, 0x7d, 0x61, 0xb9, 0x9c, 0x2e, 0xc4, 0x66 } };
  const GUID SID_SVsUIContextManager = { 0xf215afcd, 0xdc11, 0x4dc5, { 0x99, 0xd4, 0x9d, 0x0b, 0x6b, 0xc8, 0x49, 0x57 } };
  
  // Describes the state of a given UI context.
  enum UIContextState {
    // The UI context is unknown to the system. This means it has never been set. 
    // Unknown UI contexts are implicitly considered to have a state of Inactive. 
    // If all you care about is active or not, this can be considered Inactive. 
    // Though this value allows you to differentiate a UI context having been set to inactive vs being implicitly inactive.
    UIContextState_NeverSet = 0x0,

    // The UI context is set to an active state.
    UIContextState_Active = 0x1,

    // The UI context is set to an inactive state.
    UIContextState_Inactive = 0x2
  };

  // Exposes the VS UI context subsystem.
  // Safe to access from any thread except for SetUIContextState, which must occur on the UI thread.
  // VS2022
  struct __declspec(novtable) IVsUIContextManager {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIContextManager ==

    // Returns a value indicating whether the UI context subsystem is available. This is generally true except in UI-less scenarios.
    virtual HRESULT __stdcall get_AreUIContextsAvailable(VARIANT_BOOL *pfAvailable) = 0;

    virtual HRESULT __stdcall GetUIContextState(const GUID &uiContext, UIContextState *pState) = 0;

    virtual HRESULT __stdcall SetUIContextState(const GUID &uiContext, VARIANT_BOOL isActive) = 0;

    // Advises for change events for all UI contexts.
    // Safe to access from any thread, though do be aware that if this method is called off the UI thread while a 
    // context is actively being set on the UI thread, the registered callback may miss the change notification.
    // callback: The callback to invoke when a UI context value changes.
    // pCookie: A cookie representing your subscription.
    virtual HRESULT __stdcall AdviseUIContextEvents(IVsUIContextEvents *callback, DWORD *pCookie) = 0;

    virtual HRESULT __stdcall AdviseSpecificUIContextEvents(IVsUIContextEvents *callback, const GUID &uiContext, DWORD *pCookie) = 0;

    // Safe to access from any thread.
    virtual HRESULT __stdcall UnadviseUIContextEvents(DWORD cookie) = 0;
  };
#pragma endregion

  // ================================ IVsDropdownBar2 ================================
#pragma region
  const GUID IID_IVsDropdownBar2 = { 0xbc766334, 0x6e04, 0x442c, { 0x92, 0xf1, 0xf6, 0x87, 0xec, 0xf8, 0xf7, 0x41 } };

  // Provides control of the drop-down bar at the top of a code window.
  struct __declspec(novtable) IVsDropdownBar2 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsDropdownBar2 ==

    // Gets the current enabled status of a dropdown bar.
    // iCombo: The index of the combo box.
    // enabled: true if the dropdown bar is enabled, otherwise false.
    virtual HRESULT __stdcall SetComboEnabled(long iCombo, BOOL enabled) = 0;

    // Enables or disables a dropdown bar.
    // iCombo: The index of the combo box.
    // enabled: true if enabled, otherwise false.
    virtual HRESULT __stdcall GetComboEnabled(long iCombo, BOOL *enabled) = 0;
  };
#pragma endregion

  // ================================ IVsDropdownBarClient3 ================================
#pragma region
  const GUID IID_IVsDropdownBarClient3 = { 0x9f0dc18b, 0x2ed7, 0x435e, { 0xb0, 0xd2, 0x0e, 0xd5, 0x57, 0x7b, 0x96, 0x35 } };

  // Describes the contents of the dropdown bar combinations.
  struct __declspec(novtable) IVsDropdownBarClient3 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsDropdownBarClient3 ==

    // Gets the width of the column by percent. If all the columns return 0, the default width is used. The value should be between 0 and 100.
    // iCombo: The index of the combo.
    // piWidthPercent: The width in per cent.
    virtual HRESULT __stdcall GetComboWidth(long iCombo, long *piWidthPercent) = 0;

    // Gets the automation name and ID for the drop down box.
    // iCombo: The index of the combo.
    // pbstrName: The automation name of the combo.
    // pbstrId: The ID of the combo.
    virtual HRESULT __stdcall GetAutomationProperties(long iCombo, BSTR *pbstrName, BSTR *pbstrId) = 0;

    // The successor to IVsDropdownBarClient.GetEntryImage. 
    // This method also returns the image list. 
    // If this method returns a non-zero result, the dropdown bar will fallback and call IVsDropdownBarClient.GetEntryImage.
    virtual HRESULT __stdcall GetEntryImage(long iCombo, long iIndex, long *piImageIndex, HANDLE *phImageList) = 0;
  };
#pragma endregion

  // ================================ IVsCompletionSet3 ================================
#pragma region
  const GUID IID_IVsCompletionSet3 = { 0xB30B03B0, 0xDB30, 0x43C7, { 0xA9, 0x59, 0xC8, 0x15, 0x22, 0x98, 0x8F, 0x8E } };

  // Provides statement completion capabilities for the language service.
  struct __declspec(novtable) IVsCompletionSet3 {
    // == IUnknown methods ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsCompletionSet3 methods ==

    // Retrieves the image index of the context icon for a specified item.
    // If piGlyph is set to -1, no icon is displayed.
    // iIndex: The index of the item.
    // piGlyph: [out] Pointer to the image index.
    virtual HRESULT __stdcall GetContextIcon(long iIndex, long *piGlyph) = 0;

    // Retrieves the list of context images (glyphs) supported by this completion set.
    // phImageList: [out] Pointer to the handle of the image list.
    virtual HRESULT __stdcall GetContextImageList(HANDLE *phImageList) = 0;
  };
#pragma endregion

  // ================================ IVsUpdateSolutionEvents ================================
#pragma region
  const GUID IID_IVsUpdateSolutionEvents = { 0xa9f86308, 0x5ea7, 0x485d, { 0xba, 0xb8, 0xe8, 0x98, 0x9c, 0x3c, 0xfb, 0xdc } }; 
  
  // Implemented by clients (VSPackages) to sink solution and project build events.
  // 
  // [UpdateSolution_Begin] -> Dependency checking and UI prompts -> [UpdateSolution_StartUpdate] -> Build -> [UpdateSolution_Done] 
  // [OnActiveProjectCfgChange] can fire at any point after an active project configuration has changed. 
  // [UpdateSolution_Cancel] can fire when the build/update process is being actively canceled. E.g. after [UpdateSolution_Begin] or [UpdateSolution_StartUpdate]. 
  struct __declspec(novtable) IVsUpdateSolutionEvents {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUpdateSolutionEvents ==
    
    // Informs listeners that an update actions(s) may be about to begin.
    // This is sent BEFORE any dependency checking or UI prompts.
    // A component can set *pbCancel to TRUE to cancel the update actions.
    // If this event is sent the UpdateSolution_Done event should be expected whether or not it is canceled.
    //
    // Fired before any build actions have begun. This is the last chance to cancel the build before any building begins.
    virtual HRESULT __stdcall UpdateSolution_Begin(BOOL *pfCancelUpdate) = 0;

    // Called when a build is completed.
    // fSucceeded: true if no update actions failed.
    // fModified: true if any update action succeeded.
    // fCancelCommand: true if update actions were canceled.
    virtual HRESULT __stdcall UpdateSolution_Done(BOOL fSucceeded, BOOL fModified, BOOL fCancelCommand) = 0;

    // Fired just before the first project configuration is about to be built (the first update action begins). It is the last chance to cancel the update.
    // pfCancelUpdate: Pointer to a flag indicating cancel update.
    virtual HRESULT __stdcall UpdateSolution_StartUpdate(BOOL *pfCancelUpdate) = 0;

    // Called when a build (update action) is being cancelled.
    virtual HRESULT __stdcall UpdateSolution_Cancel() = 0;

    // Fired after the active project configuration for a project in the solution has changed.
    // If pIVsHierarchy is 0, sinks for this event have to assume that every project in the solution may have changed, even if there is only one project active in the solution.
    virtual HRESULT __stdcall OnActiveProjectCfgChange(IVsHierarchy *pIVsHierarchy) = 0;
  };
#pragma endregion

  // ================================ IVsUpdateSolutionEvents2 ================================
#pragma region
  const GUID IID_IVsUpdateSolutionEvents2 = { 0xf59dbc1a, 0x91c3, 0x45c9, { 0x97, 0x96, 0x1c, 0xab, 0x55, 0x85, 0x02, 0xdd } };
  struct __declspec(novtable) IVsUpdateSolutionEvents2 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUpdateSolutionEvents ==
    virtual HRESULT __stdcall UpdateSolution_Begin(BOOL *pfCancelUpdate) = 0;
    virtual HRESULT __stdcall UpdateSolution_Done(BOOL fSucceeded, BOOL fModified, BOOL fCancelCommand) = 0;
    virtual HRESULT __stdcall UpdateSolution_StartUpdate(BOOL *pfCancelUpdate) = 0;
    virtual HRESULT __stdcall UpdateSolution_Cancel() = 0;
    virtual HRESULT __stdcall OnActiveProjectCfgChange(IVsHierarchy *pIVsHierarchy) = 0;

    // == IVsUpdateSolutionEvents2 ==
    virtual HRESULT __stdcall UpdateProjectCfg_Begin(IVsHierarchy *pHierProj, IVsCfg *pCfgProj, IVsCfg *pCfgSln, DWORD dwAction, BOOL *pfCancel) = 0;
    virtual HRESULT __stdcall UpdateProjectCfg_Done(IVsHierarchy *pHierProj, IVsCfg *pCfgProj, IVsCfg *pCfgSln, DWORD dwAction, BOOL fSuccess, BOOL fCancel) = 0;
  };
#pragma endregion

  // ================================ IVsUpdateSolutionEvents3 (BROKEN) ================================
#pragma region
  const GUID IID_IVsUpdateSolutionEvents3 = { 0x40025c28, 0x3303, 0x42ca, { 0xbe, 0xd8, 0x0f, 0x3b, 0xd8, 0x56, 0xbd, 0x6d } };

  // Defines events for changes in the solution configuration. 
  // To monitor these events, implement the interface and use it as an argument to IVsSolutionBuildManager3::AdviseUpdateSolutionEvents3
  struct __declspec(novtable) IVsUpdateSolutionEvents3 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUpdateSolutionEvents3 ==
    virtual HRESULT __stdcall OnBeforeActiveSolutionCfgChange(IVsCfg *pOldActiveSlnCfg, IVsCfg *pNewActiveSlnCfg) = 0;
    virtual HRESULT __stdcall OnAfterActiveSolutionCfgChange(IVsCfg *pOldActiveSlnCfg, IVsCfg *pNewActiveSlnCfg) = 0;
  };
#pragma endregion

  // ================================ IVsUpdateSolutionEvents4 ================================
#pragma region
  const GUID IID_IVsUpdateSolutionEvents4 = { 0x84ca83ee, 0xee80, 0x42c1, { 0x99, 0xce, 0x1d, 0xe8, 0x3f, 0x2f, 0xca, 0x3a } };
  struct __declspec(novtable) IVsUpdateSolutionEvents4 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUpdateSolutionEvents4 ==
    virtual HRESULT __stdcall UpdateSolution_QueryDelayFirstUpdateAction(int *pfDelay) = 0;
    virtual HRESULT __stdcall UpdateSolution_BeginFirstUpdateAction() = 0;
    virtual HRESULT __stdcall UpdateSolution_EndLastUpdateAction() = 0;
    virtual HRESULT __stdcall UpdateSolution_BeginUpdateAction(DWORD dwAction) = 0;
    virtual HRESULT __stdcall UpdateSolution_EndUpdateAction(DWORD dwAction) = 0;
    virtual HRESULT __stdcall OnActiveProjectCfgChangeBatchBegin() = 0;
    virtual HRESULT __stdcall OnActiveProjectCfgChangeBatchEnd() = 0;
  };
#pragma endregion

  // ================================ IVsUpdateSolutionEvents5 ================================
#pragma region
  const GUID IID_IVsUpdateSolutionEvents5 = { 0x95498691, 0xcb06, 0x4bc1, { 0x8a, 0x83, 0xf8, 0x4c, 0x6d, 0x56, 0x5e, 0x21 } };
  struct __declspec(novtable) IVsUpdateSolutionEvents5 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUpdateSolutionEvents5 ==
    virtual HRESULT __stdcall UpdateSolution_QueryDelayBuildAction(DWORD dwAction, IVsTask **pDelayTask) = 0;
  };
#pragma endregion

  // ================================ IVsSolutionEvents ================================
#pragma region
  const GUID IID_IVsSolutionEvents = { 0xa8516b56, 0x7421, 0x4dbd, { 0xab, 0x87, 0x57, 0xaf, 0x7a, 0x2e, 0x85, 0xde } };
  struct __declspec(novtable) IVsSolutionEvents {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionEvents ==
    virtual HRESULT __stdcall OnAfterOpenProject(IVsHierarchy *pHierarchy, BOOL fAdded) = 0;
    virtual HRESULT __stdcall OnQueryCloseProject(IVsHierarchy *pHierarchy, BOOL fRemoving, BOOL *pfCancel) = 0;
    virtual HRESULT __stdcall OnBeforeCloseProject(IVsHierarchy *pHierarchy, BOOL fRemoved) = 0;
    virtual HRESULT __stdcall OnAfterLoadProject(IVsHierarchy *pStubHierarchy, IVsHierarchy *pRealHierarchy) = 0;
    virtual HRESULT __stdcall OnQueryUnloadProject(IVsHierarchy *pRealHierarchy, BOOL *pfCancel) = 0;
    virtual HRESULT __stdcall OnBeforeUnloadProject(IVsHierarchy *pRealHierarchy, IVsHierarchy *pStubHierarchy) = 0;
    virtual HRESULT __stdcall OnAfterOpenSolution(IUnknown *pUnkReserved, BOOL fNewSolution) = 0;
    virtual HRESULT __stdcall OnQueryCloseSolution(IUnknown *pUnkReserved, BOOL *pfCancel) = 0;
    virtual HRESULT __stdcall OnBeforeCloseSolution(IUnknown *pUnkReserved) = 0;
    virtual HRESULT __stdcall OnAfterCloseSolution(IUnknown *pUnkReserved) = 0;
  };
#pragma endregion

  // ================================ IVsSolutionEvents2 ================================
#pragma region
  const GUID IID_IVsSolutionEvents2 = { 0xa711df67, 0xb00a, 0x4e82, { 0xa9, 0x90, 0x51, 0xb2, 0xb4, 0x50, 0xea, 0x0f } };
  struct __declspec(novtable) IVsSolutionEvents2 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionEvents ==
    virtual HRESULT __stdcall OnAfterOpenProject(IVsHierarchy *pHierarchy, BOOL fAdded) = 0;
    virtual HRESULT __stdcall OnQueryCloseProject(IVsHierarchy *pHierarchy, BOOL fRemoving, BOOL *pfCancel) = 0;
    virtual HRESULT __stdcall OnBeforeCloseProject(IVsHierarchy *pHierarchy, BOOL fRemoved) = 0;
    virtual HRESULT __stdcall OnAfterLoadProject(IVsHierarchy *pStubHierarchy, IVsHierarchy *pRealHierarchy) = 0;
    virtual HRESULT __stdcall OnQueryUnloadProject(IVsHierarchy *pRealHierarchy, BOOL *pfCancel) = 0;
    virtual HRESULT __stdcall OnBeforeUnloadProject(IVsHierarchy *pRealHierarchy, IVsHierarchy *pStubHierarchy) = 0;
    virtual HRESULT __stdcall OnAfterOpenSolution(IUnknown *pUnkReserved, BOOL fNewSolution) = 0;
    virtual HRESULT __stdcall OnQueryCloseSolution(IUnknown *pUnkReserved, BOOL *pfCancel) = 0;
    virtual HRESULT __stdcall OnBeforeCloseSolution(IUnknown *pUnkReserved) = 0;
    virtual HRESULT __stdcall OnAfterCloseSolution(IUnknown *pUnkReserved) = 0;

    // == IVsSolutionEvents2 ==
    virtual HRESULT __stdcall OnAfterMergeSolution(IUnknown *pUnkReserved) = 0;
  };
#pragma endregion

  // ================================ IVsSolutionEvents3 ================================
#pragma region
  const GUID IID_IVsSolutionEvents3 = { 0xf1de2d75, 0x3b95, 0x4510, { 0x9b, 0x2b, 0x56, 0x5b, 0xc0, 0xe3, 0x88, 0x77 } };
  struct __declspec(novtable) IVsSolutionEvents3 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionEvents ==
    virtual HRESULT __stdcall OnAfterOpenProject(IVsHierarchy *pHierarchy, BOOL fAdded) = 0;
    virtual HRESULT __stdcall OnQueryCloseProject(IVsHierarchy *pHierarchy, BOOL fRemoving, BOOL *pfCancel) = 0;
    virtual HRESULT __stdcall OnBeforeCloseProject(IVsHierarchy *pHierarchy, BOOL fRemoved) = 0;
    virtual HRESULT __stdcall OnAfterLoadProject(IVsHierarchy *pStubHierarchy, IVsHierarchy *pRealHierarchy) = 0;
    virtual HRESULT __stdcall OnQueryUnloadProject(IVsHierarchy *pRealHierarchy, BOOL *pfCancel) = 0;
    virtual HRESULT __stdcall OnBeforeUnloadProject(IVsHierarchy *pRealHierarchy, IVsHierarchy *pStubHierarchy) = 0;
    virtual HRESULT __stdcall OnAfterOpenSolution(IUnknown *pUnkReserved, BOOL fNewSolution) = 0;
    virtual HRESULT __stdcall OnQueryCloseSolution(IUnknown *pUnkReserved, BOOL *pfCancel) = 0;
    virtual HRESULT __stdcall OnBeforeCloseSolution(IUnknown *pUnkReserved) = 0;
    virtual HRESULT __stdcall OnAfterCloseSolution(IUnknown *pUnkReserved) = 0;

    // == IVsSolutionEvents2 ==
    virtual HRESULT __stdcall OnAfterMergeSolution(IUnknown *pUnkReserved) = 0;

    // == IVsSolutionEvents3 ==
    virtual HRESULT __stdcall OnBeforeOpeningChildren(IVsHierarchy *pHierarchy) = 0;
    virtual HRESULT __stdcall OnAfterOpeningChildren(IVsHierarchy *pHierarchy) = 0;
    virtual HRESULT __stdcall OnBeforeClosingChildren(IVsHierarchy *pHierarchy) = 0;
    virtual HRESULT __stdcall OnAfterClosingChildren(IVsHierarchy *pHierarchy) = 0;
  };
#pragma endregion

  // ================================ IVsSolutionEvents4 ================================
#pragma region
  const GUID IID_IVsSolutionEvents4 = { 0x23ec4d20, 0x54a9, 0x4365, { 0x82, 0xc8, 0xab, 0xdf, 0xba, 0x68, 0x6e, 0xcf } };
  struct __declspec(novtable) IVsSolutionEvents4 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionEvents4 ==
    virtual HRESULT __stdcall OnAfterRenameProject(IVsHierarchy *pHierarchy) = 0;
    virtual HRESULT __stdcall OnQueryChangeProjectParent(IVsHierarchy *pHierarchy, IVsHierarchy *pNewParentHier, BOOL *pfCancel) = 0;
    virtual HRESULT __stdcall OnAfterChangeProjectParent(IVsHierarchy *pHierarchy) = 0;
    virtual HRESULT __stdcall OnAfterAsynchOpenProject(IVsHierarchy *pHierarchy, BOOL fAdded) = 0;
  };
#pragma endregion

  // ================================ IVsSolutionEvents5 ================================
#pragma region
  const GUID IID_IVsSolutionEvents5 = { 0xaf530689, 0x9987, 0x48be, { 0xaf, 0x20, 0xd9, 0x39, 0x2a, 0x9c, 0x67, 0xff } };
  struct __declspec(novtable) IVsSolutionEvents5 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionEvents5 ==
    virtual HRESULT __stdcall OnBeforeOpenProject(const GUID &guidProjectID, const GUID &guidProjectType, const wchar_t *pszFileName) = 0;
  };
#pragma endregion

  // ================================ IVsSolutionEvents6 ================================
#pragma region
  const GUID IID_IVsSolutionEvents6 = { 0x9ad84ab1, 0x5c4e, 0x4084, { 0xb1, 0x61, 0x21, 0xb6, 0x69, 0x62, 0x37, 0xcb } };
  struct __declspec(novtable) IVsSolutionEvents6 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionEvents6 ==
    virtual HRESULT __stdcall OnBeforeProjectRegisteredInRunningDocumentTable(GUID projectID, const wchar_t *projectFullPath) = 0;
    virtual HRESULT __stdcall OnAfterProjectRegisteredInRunningDocumentTable(GUID projectID, const wchar_t *projectFullPath, DWORD docCookie) = 0;
  };
#pragma endregion

  // ================================ IVsSolutionEvents7 ================================
#pragma region
  const GUID IID_IVsSolutionEvents7 = { 0xa459c228, 0x5617, 0x4136, { 0xbc, 0xbe, 0xc2, 0x82, 0xdf, 0x6d, 0x9a, 0x62 } };
  struct __declspec(novtable) IVsSolutionEvents7 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionEvents7 ==
    virtual HRESULT __stdcall OnAfterOpenFolder(const wchar_t *folderPath) = 0;
    virtual HRESULT __stdcall OnBeforeCloseFolder(const wchar_t *folderPath) = 0;
    virtual HRESULT __stdcall OnQueryCloseFolder(const wchar_t *folderPath, BOOL *pfCancel) = 0;
    virtual HRESULT __stdcall OnAfterCloseFolder(const wchar_t *folderPath) = 0;

    virtual HRESULT __stdcall Obsolete() = 0;
  };
#pragma endregion

  // ================================ IVsSolutionEvents8 ================================
#pragma region
  const GUID IID_IVsSolutionEvents8 = { 0x36b55a6c, 0x8da0, 0x428f, { 0x82, 0x8c, 0x80, 0xde, 0x91, 0x05, 0xf9, 0xc1 } };
  struct __declspec(novtable) IVsSolutionEvents8 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionEvents8 ==
    virtual HRESULT __stdcall OnAfterRenameSolution(const wchar_t *old_name, const wchar_t *new_name) = 0;
  };
#pragma endregion

  // ================================ IVsSolutionBuildManager ================================
#pragma region
  const GUID IID_IVsSolutionBuildManager = { 0x93e969d6, 0x1aa0, 0x455f, { 0xb2, 0x08, 0x6e, 0xd3, 0xc8, 0x2b, 0x5c, 0x58 } };
  // Pass    IID_IVsSolutionBuildManager for service IID

  enum VSDBGLAUNCHFLAGS {
    DBGLAUNCH_Silent = 0x1,
    DBGLAUNCH_LocalDeploy = 0x2,
    DBGLAUNCH_NoDebug = 0x4,
    DBGLAUNCH_DetachOnStop = 0x8,
    DBGLAUNCH_Selected = 0x10,
    DBGLAUNCH_StopDebuggingOnEnd = 0x20,
    DBGLAUNCH_WaitForAttachComplete = 0x40
  };

  struct __declspec(novtable) IVsSolutionBuildManager {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionBuildManager ==
    virtual HRESULT __stdcall DebugLaunch(VSDBGLAUNCHFLAGS grfLaunch) = 0;
    virtual HRESULT __stdcall StartSimpleUpdateSolutionConfiguration(DWORD dwFlags, DWORD dwDefQueryResults, BOOL fSuppressUI) = 0;
    virtual HRESULT __stdcall AdviseUpdateSolutionEvents(IVsUpdateSolutionEvents *pIVsUpdateSolutionEvents, DWORD *pdwCookie) = 0;
    virtual HRESULT __stdcall UnadviseUpdateSolutionEvents(DWORD dwCookie) = 0;
    virtual HRESULT __stdcall UpdateSolutionConfigurationIsActive(BOOL *pfIsActive) = 0;
    virtual HRESULT __stdcall CanCancelUpdateSolutionConfiguration(BOOL *pfCanCancel) = 0;
    virtual HRESULT __stdcall CancelUpdateSolutionConfiguration() = 0;
    virtual HRESULT __stdcall QueryDebugLaunch(VSDBGLAUNCHFLAGS grfLaunch, BOOL *pfCanLaunch) = 0;
    virtual HRESULT __stdcall QueryBuildManagerBusy(BOOL *pfBuildManagerBusy) = 0;
    virtual HRESULT __stdcall FindActiveProjectCfg(IVsHierarchy *pvReserved1, const wchar_t *pvReserved2, IVsHierarchy *pIVsHierarchy_RequestedProject, IVsProjectCfg **ppIVsProjectCfg_Active) = 0;
    virtual HRESULT __stdcall get_IsDebug(BOOL *pfIsDebug) = 0;
    virtual HRESULT __stdcall put_IsDebug(BOOL fIsDebug) = 0;
    virtual HRESULT __stdcall get_CodePage(UINT *puiCodePage) = 0;
    virtual HRESULT __stdcall put_CodePage(UINT uiCodePage) = 0;
    virtual HRESULT __stdcall StartSimpleUpdateProjectConfiguration(IVsHierarchy *pIVsHierarchyToBuild, IVsHierarchy *pIVsHierarchyDependent, const wchar_t *pszDependentConfigurationCanonicalName, DWORD dwFlags, DWORD dwDefQueryResults, BOOL fSuppressUI) = 0;
    virtual HRESULT __stdcall get_StartupProject(IVsHierarchy **ppHierarchy) = 0;
    virtual HRESULT __stdcall set_StartupProject(IVsHierarchy *pHierarchy) = 0;
    virtual HRESULT __stdcall GetProjectDependencies(IVsHierarchy *pHier, ULONG celt, IVsHierarchy *rgpHier[], ULONG *pcActual) = 0;
  };
#pragma endregion

  // ================================ IVsSolutionBuildManager2 ================================
#pragma region
  const GUID IID_IVsSolutionBuildManager2 = { 0x80353f58, 0xf2a3, 0x47b8, { 0xb2, 0xdf, 0x04, 0x75, 0xe0, 0x7b, 0xb1, 0xc6 } };
  // Pass    IID_IVsSolutionBuildManager  for service IID

  struct __declspec(novtable) IVsSolutionBuildManager2 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionBuildManager ==
    virtual HRESULT __stdcall DebugLaunch(VSDBGLAUNCHFLAGS grfLaunch) = 0;
    virtual HRESULT __stdcall StartSimpleUpdateSolutionConfiguration(DWORD dwFlags, DWORD dwDefQueryResults, BOOL fSuppressUI) = 0;
    virtual HRESULT __stdcall AdviseUpdateSolutionEvents(IVsUpdateSolutionEvents *pIVsUpdateSolutionEvents, DWORD *pdwCookie) = 0;
    virtual HRESULT __stdcall UnadviseUpdateSolutionEvents(DWORD dwCookie) = 0;
    virtual HRESULT __stdcall UpdateSolutionConfigurationIsActive(BOOL *pfIsActive) = 0;
    virtual HRESULT __stdcall CanCancelUpdateSolutionConfiguration(BOOL *pfCanCancel) = 0;
    virtual HRESULT __stdcall CancelUpdateSolutionConfiguration() = 0;
    virtual HRESULT __stdcall QueryDebugLaunch(VSDBGLAUNCHFLAGS grfLaunch, BOOL *pfCanLaunch) = 0;
    virtual HRESULT __stdcall QueryBuildManagerBusy(BOOL *pfBuildManagerBusy) = 0;
    virtual HRESULT __stdcall FindActiveProjectCfg(IVsHierarchy *pvReserved1, const wchar_t *pvReserved2, IVsHierarchy *pIVsHierarchy_RequestedProject, IVsProjectCfg **ppIVsProjectCfg_Active) = 0;
    virtual HRESULT __stdcall get_IsDebug(BOOL *pfIsDebug) = 0;
    virtual HRESULT __stdcall put_IsDebug(BOOL fIsDebug) = 0;
    virtual HRESULT __stdcall get_CodePage(UINT *puiCodePage) = 0;
    virtual HRESULT __stdcall put_CodePage(UINT uiCodePage) = 0;
    virtual HRESULT __stdcall StartSimpleUpdateProjectConfiguration(IVsHierarchy *pIVsHierarchyToBuild, IVsHierarchy *pIVsHierarchyDependent, const wchar_t *pszDependentConfigurationCanonicalName, DWORD dwFlags, DWORD dwDefQueryResults, BOOL fSuppressUI) = 0;
    virtual HRESULT __stdcall get_StartupProject(IVsHierarchy **ppHierarchy) = 0;
    virtual HRESULT __stdcall set_StartupProject(IVsHierarchy *pHierarchy) = 0;
    virtual HRESULT __stdcall GetProjectDependencies(IVsHierarchy *pHier, ULONG celt, IVsHierarchy *rgpHier[], ULONG *pcActual) = 0;

    // == IVsSolutionBuildManager2 ==
    virtual HRESULT __stdcall StartUpdateProjectConfigurations(UINT cProjs, IVsHierarchy *rgpHierProjs[], DWORD dwFlags, BOOL fSuppressUI) = 0;
    virtual HRESULT __stdcall CalculateProjectDependencies() = 0;
    virtual HRESULT __stdcall QueryProjectDependency(IVsHierarchy *pHier, IVsHierarchy *pHierDependentOn, BOOL *pfIsDependentOn) = 0;
    virtual HRESULT __stdcall SaveDocumentsBeforeBuild(IVsHierarchy *pHier, DWORD itemid, DWORD docCookie) = 0;
    virtual HRESULT __stdcall StartUpdateSpecificProjectConfigurations(UINT cProjs, IVsHierarchy *rgpHier[], IVsCfg *rgpCfg[], DWORD rgdwCleanFlags[], DWORD rgdwBuildFlags[], DWORD rgdwDeployFlags[], DWORD dwFlags, BOOL fSuppressUI) = 0;
  };
#pragma endregion

  // ================================ IVsSolutionBuildManager3 ================================
#pragma region
  const GUID IID_IVsSolutionBuildManager3 = { 0xb6ea87ed, 0xc498, 0x4484, { 0x81, 0xac, 0x0b, 0xed, 0x18, 0x7e, 0x28, 0xe6 } };
  // Pass    IID_IVsSolutionBuildManager  for service IID

  // Provides access to IVsUpdateSolutionEvents3 events.
  struct __declspec(novtable) IVsSolutionBuildManager3 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionBuildManager3 ==

    // Establishes client notification of solution events.
    virtual HRESULT __stdcall AdviseUpdateSolutionEvents3(IVsUpdateSolutionEvents3 *pIVsUpdateSolutionEvents3, DWORD *pdwCookie) = 0;

    // Removes the caller from the list of listeners for IVsUpdateSolutionEvents3 events.
    virtual HRESULT __stdcall UnadviseUpdateSolutionEvents3(DWORD dwCookie) = 0;

    // Determines if projects are up to date.
    virtual HRESULT __stdcall AreProjectsUpToDate(DWORD dwOptions) = 0;

    // Determines whether the hierarchy has changed since last design time expression evaluation.
    virtual HRESULT __stdcall HasHierarchyChangedSinceLastDTEE() = 0;

    // Determines if the build manager is busy.
    virtual HRESULT __stdcall QueryBuildManagerBusyEx(DWORD *pdwBuildManagerOperation) = 0;
  };
#pragma endregion

  // ================================ IVsSolutionBuildManager4 ================================
#pragma region
  const GUID IID_IVsSolutionBuildManager4 = { 0x2c07342b, 0xba98, 0x4235, { 0x98, 0x3c, 0x86, 0x38, 0x39, 0x1a, 0x42, 0x0a } };
  struct __declspec(novtable) IVsSolutionBuildManager4 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionBuildManager4 ==
    virtual HRESULT __stdcall UpdateProjectDependencies(IVsHierarchy *pHier) = 0;
  };
#pragma endregion

  // ================================ IVsSolutionBuildManager5 ================================
#pragma region
  const GUID IID_IVsSolutionBuildManager5 = { 0x75d64352, 0xc2f9, 0x4bf8, { 0x9c, 0x89, 0x57, 0xcf, 0xe5, 0x48, 0xbf, 0x75 } };
  struct __declspec(novtable) IVsSolutionBuildManager5 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionBuildManager5 ==
    virtual HRESULT __stdcall AdviseUpdateSolutionEvents4(IVsUpdateSolutionEvents4 *pIVsUpdateSolutionEvents4, DWORD *pdwCookie) = 0;
    virtual HRESULT __stdcall UnadviseUpdateSolutionEvents4(DWORD dwCookie) = 0;
    virtual HRESULT __stdcall AdviseUpdateSolutionEventsAsync(IVsUpdateSolutionEventsAsync *pIVsUpdateSolutionEventsAsync, DWORD *pdwCookie) = 0;
    virtual HRESULT __stdcall UnadviseUpdateSolutionEventsAsync(DWORD dwCookie) = 0;
    virtual HRESULT __stdcall FindActiveProjectCfgName(const GUID &rguidProjectID, BSTR *pbstrProjectCfgCanonicalName) = 0;
  };
#pragma endregion

  // ================================ IVsSolutionBuildManager6 ================================
#pragma region
  const GUID IID_IVsSolutionBuildManager6 = { 0x61aa4ea9, 0x018f, 0x4394, { 0xad, 0x31, 0x1e, 0x76, 0xd1, 0xbf, 0x80, 0xc8 } };
  struct __declspec(novtable) IVsSolutionBuildManager6 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionBuildManager6 ==
    virtual HRESULT __stdcall AdviseUpdateSolutionEventsEx(const GUID &guidActivityId, IUnknown *pSink, DWORD *pdwCookie) = 0;
  };
#pragma endregion

  // ================================ IVsSolution ================================
#pragma region
  const GUID IID_IVsSolution = { 0x7f7cd0db, 0x91ef, 0x49dc, { 0x9f, 0xa9, 0x02, 0xd1, 0x28, 0x51, 0x5d, 0xd4 } };

  enum VSENUMPROJFLAGS {
    EPF_LOADEDINSOLUTION = 0x1,
    EPF_UNLOADEDINSOLUTION = 0x2,
    EPF_ALLINSOLUTION = (EPF_LOADEDINSOLUTION | EPF_UNLOADEDINSOLUTION),
    EPF_MATCHTYPE = 0x4,
    EPF_VIRTUALVISIBLEPROJECT = 0x8,
    EPF_VIRTUALNONVISIBLEPROJECT = 0x10,
    EPF_ALLVIRTUAL = (EPF_VIRTUALVISIBLEPROJECT | EPF_VIRTUALNONVISIBLEPROJECT),
    EPF_ALLPROJECTS = (EPF_ALLINSOLUTION | EPF_ALLVIRTUAL)
  };

  enum VSSLNSAVEOPTIONS {
    SLNSAVEOPT_SaveIfDirty = 0,
    SLNSAVEOPT_PromptSave = 0x1,
    SLNSAVEOPT_SkipDocs = 0x2,
    SLNSAVEOPT_SkipProj = 0x4,
    SLNSAVEOPT_SkipSolution = 0x8,
    SLNSAVEOPT_SkipUserOptFile = 0x10,
    SLNSAVEOPT_NoSave = 0x1e,
    SLNSAVEOPT_ForceSave = 0x20,
    SLNSAVEOPT_DocClose = 0x40
  };

  enum VSSLNCLOSEOPTIONS {
    SLNCLOSEOPT_SLNSAVEOPT_MASK = 0xffff,
    SLNCLOSEOPT_UnloadProject = 0x10000,
    SLNCLOSEOPT_DeleteProject = 0x20000
  };

  enum VSUPDATEPROJREFREASON {
    UPR_NoUpdate = 0,
    UPR_ProjectRenamed = 1,
    UPR_ProjectUsedInNewSolution = 2,
    UPR_ItemRenamed = 3,
    UPR_SolutionLocationChanged = 4
  };

  enum VSADDVPFLAGS {
    ADDVP_AddToProjectWindow = 0x1,
    ADDVP_ExcludeFromBuild = 0x2,
    ADDVP_ExcludeFromDebugLaunch = 0x4,
    ADDVP_ExcludeFromDeploy = 0x8,
    ADDVP_ExcludeFromSCC = 0x10,
    ADDVP_ExcludeFromEnumOutputs = 0x20,
    ADDVP_ExcludeFromCfgUI = 0x40
  };

  enum VSPROPID {
    VSPROPID_LAST = -8000,
    VSPROPID_SolutionDirectory = -8000,
    VSPROPID_SolutionFileName = -8001,
    VSPROPID_UserOptionsFileName = -8002,
    VSPROPID_SolutionBaseName = -8003,
    VSPROPID_IsSolutionDirty = -8004,
    VSPROPID_IsSolutionOpen = -8005,
    VSPROPID_ProjectCount = -8006,
    VSPROPID_RegisteredProjExtns = -8007,
    VSPROPID_OpenProjectFilter = -8008,
    VSPROPID_FileDefaultCodePage = -8009,
    VSPROPID_SolutionFileNameBeingLoaded = -8010,
    VSPROPID_SolutionNodeCaption = -8011,
    VSPROPID_IsSolutionOpening = -8013,
    VSPROPID_IsSolutionSaveAsRequired = -8014,
    VSPROPID_CountOfProjectsBeingLoaded = -8015,
    VSPROPID_SolutionPropertyPages = -8016,
    VSPROPID_FIRST = -8016
  };

  enum VSSLNOPENOPTIONS {
    SLNOPENOPT_Silent = 0x1,
    SLNOPENOPT_AddToCurrent = 0x2,
    SLNOPENOPT_DontConvertSLN = 0x4
  };

  enum VSCREATESOLUTIONFLAGS {
    CSF_SILENT = 0x1,
    CSF_OVERWRITE = 0x2,
    CSF_TEMPORARY = 0x4,
    CSF_DELAYNOTIFY = 0x8
  };

  enum VSREMOVEVPFLAGS {
    REMOVEVP_DontCloseHierarchy = 0x1,
    REMOVEVP_DontSaveHierarchy = 0x2
  };

  enum VSGETPROJFILESFLAGS {
    GPFF_SKIPUNLOADEDPROJECTS = 0x1
  };

  struct __declspec(novtable) IVsSolution {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolution ==
    virtual HRESULT __stdcall GetProjectEnum(VSENUMPROJFLAGS grfEnumFlags, const GUID &rguidEnumOnlyThisType, IEnumHierarchies **ppEnum) = 0;
    virtual HRESULT __stdcall CreateProject(const GUID &rguidProjectType, const wchar_t *lpszMoniker, const wchar_t *lpszLocation, const wchar_t *lpszName, VSCREATEPROJFLAGS grfCreateFlags, const IID &iidProject, void **ppProject) = 0;
    virtual HRESULT __stdcall GenerateUniqueProjectName(const wchar_t *lpszRoot, BSTR *pbstrProjectName) = 0;
    virtual HRESULT __stdcall GetProjectOfGuid(const GUID &rguidProjectID, IVsHierarchy **ppHierarchy) = 0;
    virtual HRESULT __stdcall GetGuidOfProject(IVsHierarchy *pHierarchy, GUID *pguidProjectID) = 0;
    virtual HRESULT __stdcall GetSolutionInfo(BSTR *pbstrSolutionDirectory, BSTR *pbstrSolutionFile, BSTR *pbstrUserOptsFile) = 0;
    virtual HRESULT __stdcall AdviseSolutionEvents(IVsSolutionEvents *pSink, DWORD *pdwCookie) = 0;
    virtual HRESULT __stdcall UnadviseSolutionEvents(DWORD dwCookie) = 0;
    virtual HRESULT __stdcall SaveSolutionElement(VSSLNSAVEOPTIONS grfSaveOpts, IVsHierarchy *pHier, DWORD docCookie) = 0;
    virtual HRESULT __stdcall CloseSolutionElement(VSSLNCLOSEOPTIONS grfCloseOpts, IVsHierarchy *pHier, DWORD docCookie) = 0;
    virtual HRESULT __stdcall GetProjectOfProjref(const wchar_t *pszProjref, IVsHierarchy **ppHierarchy, BSTR *pbstrUpdatedProjref, VSUPDATEPROJREFREASON *puprUpdateReason) = 0;
    virtual HRESULT __stdcall GetProjrefOfProject(IVsHierarchy *pHierarchy, BSTR *pbstrProjref) = 0;
    virtual HRESULT __stdcall GetProjectInfoOfProjref(const wchar_t *pszProjref, VSHPROPID propid, VARIANT *pvar) = 0;
    virtual HRESULT __stdcall AddVirtualProject(IVsHierarchy *pHierarchy, VSADDVPFLAGS grfAddVPFlags) = 0;
    virtual HRESULT __stdcall GetItemOfProjref(const wchar_t *pszProjref, IVsHierarchy **ppHierarchy, DWORD *pitemid, BSTR *pbstrUpdatedProjref, VSUPDATEPROJREFREASON *puprUpdateReason) = 0;
    virtual HRESULT __stdcall GetProjrefOfItem(IVsHierarchy *pHierarchy, DWORD itemid, BSTR *pbstrProjref) = 0;
    virtual HRESULT __stdcall GetItemInfoOfProjref(const wchar_t *pszProjref, VSHPROPID propid, VARIANT *pvar) = 0;
    virtual HRESULT __stdcall GetProjectOfUniqueName(const wchar_t *pszUniqueName, IVsHierarchy **ppHierarchy) = 0;
    virtual HRESULT __stdcall GetUniqueNameOfProject(IVsHierarchy *pHierarchy, BSTR *pbstrUniqueName) = 0;
    virtual HRESULT __stdcall GetProperty(VSPROPID propid, VARIANT *pvar) = 0;
    virtual HRESULT __stdcall SetProperty(VSPROPID propid, VARIANT var) = 0;
    virtual HRESULT __stdcall OpenSolutionFile(VSSLNOPENOPTIONS grfOpenOpts, const wchar_t *pszFilename) = 0;
    virtual HRESULT __stdcall QueryEditSolutionFile(DWORD *pdwEditResult) = 0;
    virtual HRESULT __stdcall CreateSolution(const wchar_t *lpszLocation, const wchar_t *lpszName, VSCREATESOLUTIONFLAGS grfCreateFlags) = 0;
    virtual HRESULT __stdcall GetProjectFactory(DWORD dwReserved, GUID *pguidProjectType, const wchar_t *pszMkProject, IVsProjectFactory **ppProjectFactory) = 0;
    virtual HRESULT __stdcall GetProjectTypeGuid(DWORD dwReserved, const wchar_t *pszMkProject, GUID *pguidProjectType) = 0;
    virtual HRESULT __stdcall OpenSolutionViaDlg(const wchar_t *pszStartDirectory, BOOL fDefaultToAllProjectsFilter) = 0;
    virtual HRESULT __stdcall AddVirtualProjectEx(IVsHierarchy *pHierarchy, VSADDVPFLAGS grfAddVPFlags, const GUID &rguidProjectID) = 0;
    virtual HRESULT __stdcall QueryRenameProject(IVsProject *pProject, const wchar_t *pszMkOldName, const wchar_t *pszMkNewName, DWORD dwReserved, BOOL *pfRenameCanContinue) = 0;
    virtual HRESULT __stdcall OnAfterRenameProject(IVsProject *pProject, const wchar_t *pszMkOldName, const wchar_t *pszMkNewName, DWORD dwReserved) = 0;
    virtual HRESULT __stdcall RemoveVirtualProject(IVsHierarchy *pHierarchy, VSREMOVEVPFLAGS grfRemoveVPFlags) = 0;
    virtual HRESULT __stdcall CreateNewProjectViaDlg(const wchar_t *pszExpand, const wchar_t *pszSelect, DWORD dwReserved) = 0;
    virtual HRESULT __stdcall GetVirtualProjectFlags(IVsHierarchy *pHierarchy, VSADDVPFLAGS *pgrfAddVPFlags) = 0;
    virtual HRESULT __stdcall GenerateNextDefaultProjectName(const wchar_t *pszBaseName, const wchar_t *pszLocation, BSTR *pbstrProjectName) = 0;
    virtual HRESULT __stdcall GetProjectFilesInSolution(VSGETPROJFILESFLAGS grfGetOpts, ULONG cProjects, BSTR *rgbstrProjectNames, ULONG *pcProjectsFetched) = 0;
    virtual HRESULT __stdcall CanCreateNewProjectAtLocation(BOOL fCreateNewSolution, const wchar_t *pszFullProjectFilePath, BOOL *pfCanCreate) = 0;
  };
#pragma endregion

  // ================================ IVsSolution2 ================================
#pragma region
  const GUID IID_IVsSolution2 = { 0x95c6a090, 0xbb9e, 0x4bf2, { 0xb0, 0xbe, 0xf1, 0xd0, 0x4f, 0x0e, 0xce, 0xa3 } };
  struct __declspec(novtable) IVsSolution2 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolution ==
    virtual HRESULT __stdcall GetProjectEnum(VSENUMPROJFLAGS grfEnumFlags, const GUID &rguidEnumOnlyThisType, IEnumHierarchies **ppEnum) = 0;
    virtual HRESULT __stdcall CreateProject(const GUID &rguidProjectType, const wchar_t *lpszMoniker, const wchar_t *lpszLocation, const wchar_t *lpszName, VSCREATEPROJFLAGS grfCreateFlags, const IID &iidProject, void **ppProject) = 0;
    virtual HRESULT __stdcall GenerateUniqueProjectName(const wchar_t *lpszRoot, BSTR *pbstrProjectName) = 0;
    virtual HRESULT __stdcall GetProjectOfGuid(const GUID &rguidProjectID, IVsHierarchy **ppHierarchy) = 0;
    virtual HRESULT __stdcall GetGuidOfProject(IVsHierarchy *pHierarchy, GUID *pguidProjectID) = 0;
    virtual HRESULT __stdcall GetSolutionInfo(BSTR *pbstrSolutionDirectory, BSTR *pbstrSolutionFile, BSTR *pbstrUserOptsFile) = 0;
    virtual HRESULT __stdcall AdviseSolutionEvents(IVsSolutionEvents *pSink, DWORD *pdwCookie) = 0;
    virtual HRESULT __stdcall UnadviseSolutionEvents(DWORD dwCookie) = 0;
    virtual HRESULT __stdcall SaveSolutionElement(VSSLNSAVEOPTIONS grfSaveOpts, IVsHierarchy *pHier, DWORD docCookie) = 0;
    virtual HRESULT __stdcall CloseSolutionElement(VSSLNCLOSEOPTIONS grfCloseOpts, IVsHierarchy *pHier, DWORD docCookie) = 0;
    virtual HRESULT __stdcall GetProjectOfProjref(const wchar_t *pszProjref, IVsHierarchy **ppHierarchy, BSTR *pbstrUpdatedProjref, VSUPDATEPROJREFREASON *puprUpdateReason) = 0;
    virtual HRESULT __stdcall GetProjrefOfProject(IVsHierarchy *pHierarchy, BSTR *pbstrProjref) = 0;
    virtual HRESULT __stdcall GetProjectInfoOfProjref(const wchar_t *pszProjref, VSHPROPID propid, VARIANT *pvar) = 0;
    virtual HRESULT __stdcall AddVirtualProject(IVsHierarchy *pHierarchy, VSADDVPFLAGS grfAddVPFlags) = 0;
    virtual HRESULT __stdcall GetItemOfProjref(const wchar_t *pszProjref, IVsHierarchy **ppHierarchy, DWORD *pitemid, BSTR *pbstrUpdatedProjref, VSUPDATEPROJREFREASON *puprUpdateReason) = 0;
    virtual HRESULT __stdcall GetProjrefOfItem(IVsHierarchy *pHierarchy, DWORD itemid, BSTR *pbstrProjref) = 0;
    virtual HRESULT __stdcall GetItemInfoOfProjref(const wchar_t *pszProjref, VSHPROPID propid, VARIANT *pvar) = 0;
    virtual HRESULT __stdcall GetProjectOfUniqueName(const wchar_t *pszUniqueName, IVsHierarchy **ppHierarchy) = 0;
    virtual HRESULT __stdcall GetUniqueNameOfProject(IVsHierarchy *pHierarchy, BSTR *pbstrUniqueName) = 0;
    virtual HRESULT __stdcall GetProperty(VSPROPID propid, VARIANT *pvar) = 0;
    virtual HRESULT __stdcall SetProperty(VSPROPID propid, VARIANT var) = 0;
    virtual HRESULT __stdcall OpenSolutionFile(VSSLNOPENOPTIONS grfOpenOpts, const wchar_t *pszFilename) = 0;
    virtual HRESULT __stdcall QueryEditSolutionFile(DWORD *pdwEditResult) = 0;
    virtual HRESULT __stdcall CreateSolution(const wchar_t *lpszLocation, const wchar_t *lpszName, VSCREATESOLUTIONFLAGS grfCreateFlags) = 0;
    virtual HRESULT __stdcall GetProjectFactory(DWORD dwReserved, GUID *pguidProjectType, const wchar_t *pszMkProject, IVsProjectFactory **ppProjectFactory) = 0;
    virtual HRESULT __stdcall GetProjectTypeGuid(DWORD dwReserved, const wchar_t *pszMkProject, GUID *pguidProjectType) = 0;
    virtual HRESULT __stdcall OpenSolutionViaDlg(const wchar_t *pszStartDirectory, BOOL fDefaultToAllProjectsFilter) = 0;
    virtual HRESULT __stdcall AddVirtualProjectEx(IVsHierarchy *pHierarchy, VSADDVPFLAGS grfAddVPFlags, const GUID &rguidProjectID) = 0;
    virtual HRESULT __stdcall QueryRenameProject(IVsProject *pProject, const wchar_t *pszMkOldName, const wchar_t *pszMkNewName, DWORD dwReserved, BOOL *pfRenameCanContinue) = 0;
    virtual HRESULT __stdcall OnAfterRenameProject(IVsProject *pProject, const wchar_t *pszMkOldName, const wchar_t *pszMkNewName, DWORD dwReserved) = 0;
    virtual HRESULT __stdcall RemoveVirtualProject(IVsHierarchy *pHierarchy, VSREMOVEVPFLAGS grfRemoveVPFlags) = 0;
    virtual HRESULT __stdcall CreateNewProjectViaDlg(const wchar_t *pszExpand, const wchar_t *pszSelect, DWORD dwReserved) = 0;
    virtual HRESULT __stdcall GetVirtualProjectFlags(IVsHierarchy *pHierarchy, VSADDVPFLAGS *pgrfAddVPFlags) = 0;
    virtual HRESULT __stdcall GenerateNextDefaultProjectName(const wchar_t *pszBaseName, const wchar_t *pszLocation, BSTR *pbstrProjectName) = 0;
    virtual HRESULT __stdcall GetProjectFilesInSolution(VSGETPROJFILESFLAGS grfGetOpts, ULONG cProjects, BSTR *rgbstrProjectNames, ULONG *pcProjectsFetched) = 0;
    virtual HRESULT __stdcall CanCreateNewProjectAtLocation(BOOL fCreateNewSolution, const wchar_t *pszFullProjectFilePath, BOOL *pfCanCreate) = 0;

    // == IVsSolution2 ==
    virtual HRESULT __stdcall UpdateProjectFileLocation(IVsHierarchy *pHierarchy) = 0;
  };
#pragma endregion

  // ================================ IVsSolution3 ================================
#pragma region
  const GUID IID_IVsSolution3 = { 0x58dcf7bf, 0xf14e, 0x43ec, { 0xa7, 0xb2, 0x9f, 0x78, 0xed, 0xd0, 0x64, 0x18 } };

  enum VSCREATENEWPROJVIADLGEXFLAGS {
    VNPVDE_ALWAYSNEWSOLUTION = 0x1,
    VNPVDE_OVERRIDEBROWSEBUTTON = 0x2,
    VNPVDE_ALWAYSADDTOSOLUTION = 0x4,
    VNPVDE_ADDNESTEDTOSELECTION = 0x8,
    VNPVDE_USENEWWEBSITEDLG = 0x10
  };

  enum VSSAVEDEFERREDSAVEFLAGS {
    VSDSF_HIDEADDTOSOURCECONTROL = 0x1
  };

  struct __declspec(novtable) IVsSolution3 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolution3 ==
    virtual HRESULT __stdcall CreateNewProjectViaDlgEx(const wchar_t *pszDlgTitle, const wchar_t *pszTemplateDir, const wchar_t *pszExpand, const wchar_t *pszSelect, const wchar_t *pszHelpTopic, VSCREATENEWPROJVIADLGEXFLAGS cnpvdeFlags, IVsBrowseProjectLocation *pBrowse) = 0;
    virtual HRESULT __stdcall GetUniqueUINameOfProject(IVsHierarchy *pHierarchy, BSTR *pbstrUniqueName) = 0;
    virtual HRESULT __stdcall CheckForAndSaveDeferredSaveSolution(BOOL fCloseSolution, const wchar_t *pszMessage, const wchar_t *pszTitle, VSSAVEDEFERREDSAVEFLAGS grfFlags) = 0;
    virtual HRESULT __stdcall UpdateProjectFileLocationForUpgrade(const wchar_t *pszCurrentLocation, const wchar_t *pszUpgradedLocation) = 0;
  };
#pragma endregion

  // ================================ IVsSolution4 ================================
#pragma region
  const GUID IID_IVsSolution4 = { 0xd2fb5b25, 0xeaf0, 0x4be9, { 0x8e, 0x9b, 0xf2, 0xc6, 0x62, 0xab, 0x98, 0x26 } };

  enum VSBSLFLAGS {
    VSBSLFLAGS_None = 0,
    VSBSLFLAGS_NotCancelable = 0x1,
    VSBSLFLAGS_LoadBuildDependencies = 0x2,
    VSBSLFLAGS_ExpandProjectOnLoad = 0x4,
    VSBSLFLAGS_SelectProjectOnLoad = 0x8,
    VSBSLFLAGS_LoadAllPendingProjects = 0x10
  };

  enum VSProjectUnloadStatus {
    UNLOADSTATUS_UnloadedByUser = 0,
    UNLOADSTATUS_LoadPendingIfNeeded = 1,
    UNLOADSTATUS_StorageNotLoadable = 2,
    UNLOADSTATUS_StorageNotAvailable = 3,
    UNLOADSTATUS_UpgradeFailed = 4
  };

  struct __declspec(novtable) IVsSolution4 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolution4 ==
    virtual HRESULT __stdcall WriteUserOptsFile() = 0;
    virtual HRESULT __stdcall IsBackgroundSolutionLoadEnabled(VARIANT_BOOL *pfIsEnabled) = 0;
    virtual HRESULT __stdcall EnsureProjectsAreLoaded(DWORD cProjects, GUID guidProjects[], VSBSLFLAGS grfFlags) = 0;
    virtual HRESULT __stdcall EnsureProjectIsLoaded(const GUID &guidProject, VSBSLFLAGS grfFlags) = 0;
    virtual HRESULT __stdcall EnsureSolutionIsLoaded(VSBSLFLAGS grfFlags) = 0;
    virtual HRESULT __stdcall ReloadProject(const GUID &guidProjectID) = 0;
    virtual HRESULT __stdcall UnloadProject(const GUID &guidProjectID, VSProjectUnloadStatus dwUnloadStatus) = 0;
  };
#pragma endregion

  // ================================ IVsSolution5 ================================
#pragma region
  const GUID IID_IVsSolution5 = { 0x90570d49, 0x7b10, 0x4dcd, { 0xb9, 0xac, 0x53, 0x0d, 0x91, 0xf4, 0xeb, 0xb5 } };
  struct __declspec(novtable) IVsSolution5 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolution5 ==
    virtual HRESULT __stdcall ResolveFaultedProjects(DWORD cHierarchies, IVsHierarchy *rgHierarchies[], IVsPropertyBag *pProjectFaultResolutionContext, DWORD *pcResolved, DWORD *pcFailed) = 0;
    virtual HRESULT __stdcall GetGuidOfProjectFile(const wchar_t *pszProjectFile, GUID *pProjectGuid) = 0;
  };
#pragma endregion

  // ================================ IVsSolution6 ================================
#pragma region
  const GUID IID_IVsSolution6 = { 0x96cb263f, 0xeb15, 0x4f70, { 0xb7, 0x35, 0xad, 0x5a, 0xd7, 0xf6, 0xd3, 0x63 } };
  struct __declspec(novtable) IVsSolution6 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolution6 ==
    virtual HRESULT __stdcall SetProjectParent(IVsHierarchy *pProject, IVsHierarchy *pParent) = 0;
    virtual HRESULT __stdcall AddNewProjectFromTemplate(const wchar_t *szTemplatePath, SAFEARRAY *rgCustomParams, const wchar_t *szTargetFramework, const wchar_t *szDestinationFolder, const wchar_t *szProjectName, IVsHierarchy *pParent, IVsHierarchy **ppNewProj) = 0;
    virtual HRESULT __stdcall AddExistingProject(const wchar_t *szFullPath, IVsHierarchy *pParent, IVsHierarchy **ppNewProj) = 0;
    virtual HRESULT __stdcall BrowseForExistingProject(const wchar_t *szDialogTitle, const wchar_t *szStartupLocation, GUID preferedProjectType, BSTR *pbstrSelected) = 0;
  };
#pragma endregion

  // ================================ IVsSolution7 ================================
#pragma region
  const GUID IID_IVsSolution7 = { 0xd32b0c42, 0x8aee, 0x4772, { 0xb5, 0xc3, 0x04, 0x56, 0x5c, 0xda, 0x5a, 0x47 } };
  struct __declspec(novtable) IVsSolution7 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolution7 ==
    virtual HRESULT __stdcall OpenFolder(const wchar_t *folderPath) = 0;
    virtual HRESULT __stdcall CloseFolder(const wchar_t *folderPath) = 0;
    virtual HRESULT __stdcall IsSolutionLoadDeferred(VARIANT_BOOL *deferred) = 0;
    virtual HRESULT __stdcall IsDeferredProjectLoadAllowed(const wchar_t *projectFullPath, VARIANT_BOOL *deferredLoadAllowed) = 0;
  };
#pragma endregion

  // ================================ IVsSolution8 ================================
#pragma region
  const GUID IID_IVsSolution8 = { 0x51a3a58a, 0xe65e, 0x46f8, { 0xa6, 0x3c, 0x49, 0x66, 0x5a, 0x67, 0x50, 0x16 } };

  enum VSBatchProjectActionFlags {
    BPAF_CLOSE_FILES = 0x1,
    BPAF_PROMPT_SAVE = 0x2,
    BPAF_ALLOW_RELOAD_SOLUTION = 0x4,
    BPAF_IGNORE_SELFRELOAD_PROJECTS = 0x8
  };

  struct __declspec(novtable) IVsSolution8 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolution8 ==
    virtual HRESULT __stdcall AdviseSolutionEventsEx(const GUID &guidCallerId, IUnknown *pSink, DWORD *pdwCookie) = 0;
    virtual HRESULT __stdcall BatchProjectAction(VSBatchProjectAction action, VSBatchProjectActionFlags dwFlags, DWORD dwNumProjects, GUID rgProjects[], IVsBatchProjectActionContext **pContext) = 0;
    virtual HRESULT __stdcall IsBatchProjectActionBlocking(VSBatchProjectAction action, VSBatchProjectActionFlags dwFlags, DWORD dwNumProjects, GUID rgProjects[], BOOL *pfIsBlocking) = 0;
    virtual HRESULT __stdcall GetCurrentBatchProjectAction(IVsBatchProjectActionContext **pContext) = 0;
  };
#pragma endregion

  // ================================ IVsUIShell ================================
#pragma region
  const GUID IID_IVsUIShell = { 0xb61fc35b, 0xeebf, 0x4dec, { 0xbf, 0xf1, 0x28, 0xa2, 0xdd, 0x43, 0xc3, 0x8f } };

  enum VSFINDTOOLWIN {
    FTW_fForceCreate = 0x80000,
    FTW_fFindFirst = 0x800000,
    FTW_fFrameOnly = 0x1000000
  };

  enum VSCREATETOOLWIN {
    CTW_RESERVED_MASK = 0xffff,
    CTW_fInitNew = 0x10000,
    CTW_fActivateWithProject = 0x20000,
    CTW_fActivateWithDocument = 0x40000,
    CTW_fForceCreate = 0x80000,
    CTW_fHasBorder = 0x100000,
    CTW_fMultiInstance = 0x200000,
    CTW_fToolbarHost = 0x400000
  };

  enum OLEMSGBUTTON {
    OLEMSGBUTTON_OK = 0,
    OLEMSGBUTTON_OKCANCEL = 1,
    OLEMSGBUTTON_ABORTRETRYIGNORE = 2,
    OLEMSGBUTTON_YESNOCANCEL = 3,
    OLEMSGBUTTON_YESNO = 4,
    OLEMSGBUTTON_RETRYCANCEL = 5,
    OLEMSGBUTTON_YESALLNOCANCEL = 6
  };

  enum OLEMSGDEFBUTTON {
    OLEMSGDEFBUTTON_FIRST = 0,
    OLEMSGDEFBUTTON_SECOND = 1,
    OLEMSGDEFBUTTON_THIRD = 2,
    OLEMSGDEFBUTTON_FOURTH = 3
  };

  enum OLEMSGICON {
    OLEMSGICON_NOICON = 0,
    OLEMSGICON_CRITICAL = 1,
    OLEMSGICON_QUERY = 2,
    OLEMSGICON_WARNING = 3,
    OLEMSGICON_INFO = 4
  };

  enum VSCREATEDOCWIN {
    CDW_RDTFLAGS_MASK = 0xfffff,
    CDW_fDockable = 0x100000,
    CDW_fAltDocData = 0x200000,
    CDW_fCreateNewWindow = 0x400000
  };

  enum VSSAVEFLAGS {
    VSSAVE_Save = 0,
    VSSAVE_SaveAs = 1,
    VSSAVE_SilentSave = 2,
    VSSAVE_SaveCopyAs = 3
  };

  enum VSSYSCOLOR {
    VSCOLOR_LIGHT = -1,
    VSCOLOR_MEDIUM = -2,
    VSCOLOR_DARK = -3,
    VSCOLOR_LIGHTCAPTION = -4,
    VSCOLOR_LAST = -4
  };

  enum DBGMODE {
    DBGMODE_Design = 0,
    DBGMODE_Break = 0x1,
    DBGMODE_Run = 0x2,
    DBGMODE_Enc = 0x10000000,
    DBGMODE_EncMask = 0xf0000000
  };

  struct VSOPENFILENAMEW {
    DWORD lStructSize;
    HWND hwndOwner;
    LPCWSTR pwzDlgTitle;
    LPWSTR pwzFileName;
    DWORD nMaxFileName;
    LPCWSTR pwzInitialDir;
    LPCWSTR pwzFilter;
    DWORD nFilterIndex;
    DWORD nFileOffset;
    DWORD nFileExtension;
    DWORD dwHelpTopic;
    DWORD dwFlags;
  };

  struct VSSAVEFILENAMEW {
    DWORD lStructSize;
    HWND hwndOwner;
    LPCWSTR pwzDlgTitle;
    LPWSTR pwzFileName;
    DWORD nMaxFileName;
    LPCWSTR pwzInitialDir;
    LPCWSTR pwzFilter;
    DWORD nFilterIndex;
    DWORD nFileOffset;
    DWORD nFileExtension;
    DWORD dwHelpTopic;
    DWORD dwFlags;
    IVsSaveOptionsDlg *pSaveOpts;
  };

  struct VSBROWSEINFOW {
    DWORD lStructSize;
    HWND hwndOwner;
    LPCWSTR pwzDlgTitle;
    LPWSTR pwzDirName;
    DWORD nMaxDirName;
    LPCWSTR pwzInitialDir;
    DWORD dwHelpTopic;
    DWORD dwFlags;
  };

  enum RemoveBFDirection {
    RemovePrev = 0,
    RemoveNext = (RemovePrev + 1)
  };

  struct __declspec(novtable) IVsUIShell {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIShell ==
    virtual HRESULT __stdcall GetToolWindowEnum(IEnumWindowFrames **ppEnum) = 0;
    virtual HRESULT __stdcall GetDocumentWindowEnum(IEnumWindowFrames **ppEnum) = 0;
    virtual HRESULT __stdcall FindToolWindow(VSFINDTOOLWIN grfFTW, const GUID &rguidPersistenceSlot, IVsWindowFrame **ppWindowFrame) = 0;
    virtual HRESULT __stdcall CreateToolWindow(VSCREATETOOLWIN grfCTW, DWORD dwToolWindowId, IUnknown *punkTool, const IID &rclsidTool, const GUID &rguidPersistenceSlot, const GUID &rguidAutoActivate, IServiceProvider *pSP, const wchar_t *pszCaption, BOOL *pfDefaultPosition, IVsWindowFrame **ppWindowFrame) = 0;
    virtual HRESULT __stdcall CreateDocumentWindow(VSCREATEDOCWIN grfCDW, const wchar_t *pszMkDocument, IVsUIHierarchy *pUIH, DWORD itemid, IUnknown *punkDocView, IUnknown *punkDocData, const GUID &rguidEditorType, const wchar_t *pszPhysicalView, const GUID &rguidCmdUI, IServiceProvider *pSP, const wchar_t *pszOwnerCaption, const wchar_t *pszEditorCaption, BOOL *pfDefaultPosition, IVsWindowFrame **ppWindowFrame) = 0;
    virtual HRESULT __stdcall SetErrorInfo(HRESULT hr, const wchar_t *pszDescription, DWORD dwReserved, const wchar_t *pszHelpKeyword, const wchar_t *pszSource) = 0;
    virtual HRESULT __stdcall ReportErrorInfo(HRESULT hr) = 0;
    virtual HRESULT __stdcall GetDialogOwnerHwnd(HWND *phwnd) = 0;
    virtual HRESULT __stdcall EnableModeless(BOOL fEnable) = 0;
    virtual HRESULT __stdcall SaveDocDataToFile(VSSAVEFLAGS grfSave, IUnknown *pPersistFile, const wchar_t *pszUntitledPath, BSTR *pbstrDocumentNew, BOOL *pfCanceled) = 0;
    virtual HRESULT __stdcall SetupToolbar(HWND hwnd, IVsToolWindowToolbar *ptwt, IVsToolWindowToolbarHost **pptwth) = 0;
    virtual HRESULT __stdcall SetForegroundWindow() = 0;
    virtual HRESULT __stdcall TranslateAcceleratorAsACmd(LPMSG pMsg) = 0;
    virtual HRESULT __stdcall UpdateCommandUI(BOOL fImmediateUpdate) = 0;
    virtual HRESULT __stdcall UpdateDocDataIsDirtyFeedback(DWORD docCookie, BOOL fDirty) = 0;
    virtual HRESULT __stdcall RefreshPropertyBrowser(DISPID dispid) = 0;
    virtual HRESULT __stdcall SetWaitCursor() = 0;
    virtual HRESULT __stdcall PostExecCommand(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn) = 0;
    virtual HRESULT __stdcall ShowContextMenu(DWORD dwCompRole, const IID &rclsidActive, LONG nMenuId, POINTS *pos, IOleCommandTarget *pCmdTrgtActive) = 0;
    virtual HRESULT __stdcall ShowMessageBox(DWORD dwCompRole, const IID &rclsidComp, LPOLESTR pszTitle, LPOLESTR pszText, LPOLESTR pszHelpFile, DWORD dwHelpContextID, OLEMSGBUTTON msgbtn, OLEMSGDEFBUTTON msgdefbtn, OLEMSGICON msgicon, BOOL fSysAlert, LONG *pnResult) = 0;
    virtual HRESULT __stdcall SetMRUComboText(const GUID *pguidCmdGroup, DWORD dwCmdId, LPSTR lpszText, BOOL fAddToList) = 0;
    virtual HRESULT __stdcall SetToolbarVisibleInFullScreen(const GUID *pguidCmdGroup, DWORD dwToolbarId, BOOL fVisibleInFullScreen) = 0;
    virtual HRESULT __stdcall FindToolWindowEx(VSFINDTOOLWIN grfFTW, const GUID &rguidPersistenceSlot, DWORD dwToolWinId, IVsWindowFrame **ppWindowFrame) = 0;
    virtual HRESULT __stdcall GetAppName(BSTR *pbstrAppName) = 0;
    virtual HRESULT __stdcall GetVSSysColor(VSSYSCOLOR dwSysColIndex, DWORD *pdwRGBval) = 0;
    virtual HRESULT __stdcall SetMRUComboTextW(const GUID *pguidCmdGroup, DWORD dwCmdId, LPWSTR pwszText, BOOL fAddToList) = 0;
    virtual HRESULT __stdcall PostSetFocusMenuCommand(const GUID *pguidCmdGroup, DWORD nCmdId) = 0;
    virtual HRESULT __stdcall GetCurrentBFNavigationItem(IVsWindowFrame **ppWindowFrame, BSTR *pbstrData, IUnknown **ppunk) = 0;
    virtual HRESULT __stdcall AddNewBFNavigationItem(IVsWindowFrame *pWindowFrame, BSTR bstrData, IUnknown *punk, BOOL fReplaceCurrent) = 0;
    virtual HRESULT __stdcall OnModeChange(DBGMODE dbgmodeNew) = 0;
    virtual HRESULT __stdcall GetErrorInfo(BSTR *pbstrErrText) = 0;
    virtual HRESULT __stdcall GetOpenFileNameViaDlg(VSOPENFILENAMEW *pOpenFileName) = 0;
    virtual HRESULT __stdcall GetSaveFileNameViaDlg(VSSAVEFILENAMEW *pSaveFileName) = 0;
    virtual HRESULT __stdcall GetDirectoryViaBrowseDlg(VSBROWSEINFOW *pBrowse) = 0;
    virtual HRESULT __stdcall CenterDialogOnWindow(HWND hwndDialog, HWND hwndParent) = 0;
    virtual HRESULT __stdcall GetPreviousBFNavigationItem(IVsWindowFrame **ppWindowFrame, BSTR *pbstrData, IUnknown **ppunk) = 0;
    virtual HRESULT __stdcall GetNextBFNavigationItem(IVsWindowFrame **ppWindowFrame, BSTR *pbstrData, IUnknown **ppunk) = 0;
    virtual HRESULT __stdcall GetURLViaDlg(const wchar_t *pszDlgTitle, const wchar_t *pszStaticLabel, const wchar_t *pszHelpTopic, BSTR *pbstrURL) = 0;
    virtual HRESULT __stdcall RemoveAdjacentBFNavigationItem(RemoveBFDirection rdDir) = 0;
    virtual HRESULT __stdcall RemoveCurrentNavigationDupes(RemoveBFDirection rdDir) = 0;
  };
#pragma endregion

  // ================================ IVsUIShell2 ================================
#pragma region
  const GUID IID_IVsUIShell2 = { 0x4e6b6ef9, 0x8e3d, 0x4756, { 0x99, 0xe9, 0x11, 0x92, 0xba, 0xad, 0x54, 0x96 } };

  struct VSNSEBROWSEINFOW {
    DWORD lStructSize;
    LPCOLESTR pszNamespaceGUID;
    LPCOLESTR pszTrayDisplayName;
    LPCOLESTR pszProtocolPrefix;
    BOOL fOnlyShowNSEInTray;
  };

  struct VSSAVETREEITEM {
    VSRDTSAVEOPTIONS grfSave;
    DWORD docCookie;
    IVsHierarchy *pHier;
    DWORD itemid;
  };

  enum VSSYSCOLOREX {
    VSCOLOR_ACCENT_BORDER = -5,
    VSCOLOR_ACCENT_DARK = -6,
    VSCOLOR_ACCENT_LIGHT = -7,
    VSCOLOR_ACCENT_MEDIUM = -8,
    VSCOLOR_ACCENT_PALE = -9,
    VSCOLOR_COMMANDBAR_BORDER = -10,
    VSCOLOR_COMMANDBAR_DRAGHANDLE = -11,
    VSCOLOR_COMMANDBAR_DRAGHANDLE_SHADOW = -12,
    VSCOLOR_COMMANDBAR_GRADIENT_BEGIN = -13,
    VSCOLOR_COMMANDBAR_GRADIENT_END = -14,
    VSCOLOR_COMMANDBAR_GRADIENT_MIDDLE = -15,
    VSCOLOR_COMMANDBAR_HOVER = -16,
    VSCOLOR_COMMANDBAR_HOVEROVERSELECTED = -17,
    VSCOLOR_COMMANDBAR_HOVEROVERSELECTEDICON = -18,
    VSCOLOR_COMMANDBAR_HOVEROVERSELECTEDICON_BORDER = -19,
    VSCOLOR_COMMANDBAR_SELECTED = -20,
    VSCOLOR_COMMANDBAR_SHADOW = -21,
    VSCOLOR_COMMANDBAR_TEXT_ACTIVE = -22,
    VSCOLOR_COMMANDBAR_TEXT_HOVER = -23,
    VSCOLOR_COMMANDBAR_TEXT_INACTIVE = -24,
    VSCOLOR_COMMANDBAR_TEXT_SELECTED = -25,
    VSCOLOR_CONTROL_EDIT_HINTTEXT = -26,
    VSCOLOR_CONTROL_EDIT_REQUIRED_BACKGROUND = -27,
    VSCOLOR_CONTROL_EDIT_REQUIRED_HINTTEXT = -28,
    VSCOLOR_CONTROL_LINK_TEXT = -29,
    VSCOLOR_CONTROL_LINK_TEXT_HOVER = -30,
    VSCOLOR_CONTROL_LINK_TEXT_PRESSED = -31,
    VSCOLOR_CONTROL_OUTLINE = -32,
    VSCOLOR_DEBUGGER_DATATIP_ACTIVE_BACKGROUND = -33,
    VSCOLOR_DEBUGGER_DATATIP_ACTIVE_BORDER = -34,
    VSCOLOR_DEBUGGER_DATATIP_ACTIVE_HIGHLIGHT = -35,
    VSCOLOR_DEBUGGER_DATATIP_ACTIVE_HIGHLIGHTTEXT = -36,
    VSCOLOR_DEBUGGER_DATATIP_ACTIVE_SEPARATOR = -37,
    VSCOLOR_DEBUGGER_DATATIP_ACTIVE_TEXT = -38,
    VSCOLOR_DEBUGGER_DATATIP_INACTIVE_BACKGROUND = -39,
    VSCOLOR_DEBUGGER_DATATIP_INACTIVE_BORDER = -40,
    VSCOLOR_DEBUGGER_DATATIP_INACTIVE_HIGHLIGHT = -41,
    VSCOLOR_DEBUGGER_DATATIP_INACTIVE_HIGHLIGHTTEXT = -42,
    VSCOLOR_DEBUGGER_DATATIP_INACTIVE_SEPARATOR = -43,
    VSCOLOR_DEBUGGER_DATATIP_INACTIVE_TEXT = -44,
    VSCOLOR_DESIGNER_BACKGROUND = -45,
    VSCOLOR_DESIGNER_SELECTIONDOTS = -46,
    VSCOLOR_DESIGNER_TRAY = -47,
    VSCOLOR_DESIGNER_WATERMARK = -48,
    VSCOLOR_EDITOR_EXPANSION_BORDER = -49,
    VSCOLOR_EDITOR_EXPANSION_FILL = -50,
    VSCOLOR_EDITOR_EXPANSION_LINK = -51,
    VSCOLOR_EDITOR_EXPANSION_TEXT = -52,
    VSCOLOR_ENVIRONMENT_BACKGROUND = -53,
    VSCOLOR_ENVIRONMENT_BACKGROUND_GRADIENTBEGIN = -54,
    VSCOLOR_ENVIRONMENT_BACKGROUND_GRADIENTEND = -55,
    VSCOLOR_FILETAB_BORDER = -56,
    VSCOLOR_FILETAB_CHANNELBACKGROUND = -57,
    VSCOLOR_FILETAB_GRADIENTDARK = -58,
    VSCOLOR_FILETAB_GRADIENTLIGHT = -59,
    VSCOLOR_FILETAB_SELECTEDBACKGROUND = -60,
    VSCOLOR_FILETAB_SELECTEDBORDER = -61,
    VSCOLOR_FILETAB_SELECTEDTEXT = -62,
    VSCOLOR_FILETAB_TEXT = -63,
    VSCOLOR_FORMSMARTTAG_ACTIONTAG_BORDER = -64,
    VSCOLOR_FORMSMARTTAG_ACTIONTAG_FILL = -65,
    VSCOLOR_FORMSMARTTAG_OBJECTTAG_BORDER = -66,
    VSCOLOR_FORMSMARTTAG_OBJECTTAG_FILL = -67,
    VSCOLOR_GRID_HEADING_BACKGROUND = -68,
    VSCOLOR_GRID_HEADING_TEXT = -69,
    VSCOLOR_GRID_LINE = -70,
    VSCOLOR_HELP_HOWDOI_PANE_BACKGROUND = -71,
    VSCOLOR_HELP_HOWDOI_PANE_LINK = -72,
    VSCOLOR_HELP_HOWDOI_PANE_TEXT = -73,
    VSCOLOR_HELP_HOWDOI_TASK_BACKGROUND = -74,
    VSCOLOR_HELP_HOWDOI_TASK_LINK = -75,
    VSCOLOR_HELP_HOWDOI_TASK_TEXT = -76,
    VSCOLOR_HELP_SEARCH_FRAME_BACKGROUND = -77,
    VSCOLOR_HELP_SEARCH_FRAME_TEXT = -78,
    VSCOLOR_HELP_SEARCH_BORDER = -79,
    VSCOLOR_HELP_SEARCH_FITLER_TEXT = -80,
    VSCOLOR_HELP_SEARCH_FITLER_BACKGROUND = -81,
    VSCOLOR_HELP_SEARCH_FITLER_BORDER = -82,
    VSCOLOR_HELP_SEARCH_PROVIDER_UNSELECTED_BACKGROUND = -83,
    VSCOLOR_HELP_SEARCH_PROVIDER_UNSELECTED_TEXT = -84,
    VSCOLOR_HELP_SEARCH_PROVIDER_SELECTED_BACKGROUND = -85,
    VSCOLOR_HELP_SEARCH_PROVIDER_SELECTED_TEXT = -86,
    VSCOLOR_HELP_SEARCH_PROVIDER_ICON = -87,
    VSCOLOR_HELP_SEARCH_RESULT_LINK_SELECTED = -88,
    VSCOLOR_HELP_SEARCH_RESULT_LINK_UNSELECTED = -89,
    VSCOLOR_HELP_SEARCH_RESULT_SELECTED_BACKGROUND = -90,
    VSCOLOR_HELP_SEARCH_RESULT_SELECTED_TEXT = -91,
    VSCOLOR_HELP_SEARCH_BACKGROUND = -92,
    VSCOLOR_HELP_SEARCH_TEXT = -93,
    VSCOLOR_HELP_SEARCH_PANEL_RULES = -94,
    VSCOLOR_MDICLIENT_BORDER = -95,
    VSCOLOR_PANEL_BORDER = -96,
    VSCOLOR_PANEL_GRADIENTDARK = -97,
    VSCOLOR_PANEL_GRADIENTLIGHT = -98,
    VSCOLOR_PANEL_HOVEROVERCLOSE_BORDER = -99,
    VSCOLOR_PANEL_HOVEROVERCLOSE_FILL = -100,
    VSCOLOR_PANEL_HYPERLINK = -101,
    VSCOLOR_PANEL_HYPERLINK_HOVER = -102,
    VSCOLOR_PANEL_HYPERLINK_PRESSED = -103,
    VSCOLOR_PANEL_SEPARATOR = -104,
    VSCOLOR_PANEL_SUBGROUPSEPARATOR = -105,
    VSCOLOR_PANEL_TEXT = -106,
    VSCOLOR_PANEL_TITLEBAR = -107,
    VSCOLOR_PANEL_TITLEBAR_TEXT = -108,
    VSCOLOR_PANEL_TITLEBAR_UNSELECTED = -109,
    VSCOLOR_PROJECTDESIGNER_BACKGROUND_GRADIENTBEGIN = -110,
    VSCOLOR_PROJECTDESIGNER_BACKGROUND_GRADIENTEND = -111,
    VSCOLOR_PROJECTDESIGNER_BORDER_OUTSIDE = -112,
    VSCOLOR_PROJECTDESIGNER_BORDER_INSIDE = -113,
    VSCOLOR_PROJECTDESIGNER_CONTENTS_BACKGROUND = -114,
    VSCOLOR_PROJECTDESIGNER_TAB_BACKGROUND_GRADIENTBEGIN = -115,
    VSCOLOR_PROJECTDESIGNER_TAB_BACKGROUND_GRADIENTEND = -116,
    VSCOLOR_PROJECTDESIGNER_TAB_SELECTED_INSIDEBORDER = -117,
    VSCOLOR_PROJECTDESIGNER_TAB_SELECTED_BORDER = -118,
    VSCOLOR_PROJECTDESIGNER_TAB_SELECTED_HIGHLIGHT1 = -119,
    VSCOLOR_PROJECTDESIGNER_TAB_SELECTED_HIGHLIGHT2 = -120,
    VSCOLOR_PROJECTDESIGNER_TAB_SELECTED_BACKGROUND = -121,
    VSCOLOR_PROJECTDESIGNER_TAB_SEP_BOTTOM_GRADIENTBEGIN = -122,
    VSCOLOR_PROJECTDESIGNER_TAB_SEP_BOTTOM_GRADIENTEND = -123,
    VSCOLOR_PROJECTDESIGNER_TAB_SEP_TOP_GRADIENTBEGIN = -124,
    VSCOLOR_PROJECTDESIGNER_TAB_SEP_TOP_GRADIENTEND = -125,
    VSCOLOR_SCREENTIP_BORDER = -126,
    VSCOLOR_SCREENTIP_BACKGROUND = -127,
    VSCOLOR_SCREENTIP_TEXT = -128,
    VSCOLOR_SIDEBAR_BACKGROUND = -129,
    VSCOLOR_SIDEBAR_GRADIENTDARK = -130,
    VSCOLOR_SIDEBAR_GRADIENTLIGHT = -131,
    VSCOLOR_SIDEBAR_TEXT = -132,
    VSCOLOR_SMARTTAG_BORDER = -133,
    VSCOLOR_SMARTTAG_FILL = -134,
    VSCOLOR_SMARTTAG_HOVER_BORDER = -135,
    VSCOLOR_SMARTTAG_HOVER_FILL = -136,
    VSCOLOR_SMARTTAG_HOVER_TEXT = -137,
    VSCOLOR_SMARTTAG_TEXT = -138,
    VSCOLOR_SNAPLINES = -139,
    VSCOLOR_SNAPLINES_PADDING = -140,
    VSCOLOR_SNAPLINES_TEXTBASELINE = -141,
    VSCOLOR_SORT_BACKGROUND = -142,
    VSCOLOR_SORT_TEXT = -143,
    VSCOLOR_TASKLIST_GRIDLINES = -144,
    VSCOLOR_TITLEBAR_ACTIVE = -145,
    VSCOLOR_TITLEBAR_ACTIVE_GRADIENTBEGIN = -146,
    VSCOLOR_TITLEBAR_ACTIVE_GRADIENTEND = -147,
    VSCOLOR_TITLEBAR_ACTIVE_TEXT = -148,
    VSCOLOR_TITLEBAR_INACTIVE = -149,
    VSCOLOR_TITLEBAR_INACTIVE_GRADIENTBEGIN = -150,
    VSCOLOR_TITLEBAR_INACTIVE_GRADIENTEND = -151,
    VSCOLOR_TITLEBAR_INACTIVE_TEXT = -152,
    VSCOLOR_TOOLBOX_BACKGROUND = -153,
    VSCOLOR_TOOLBOX_DIVIDER = -154,
    VSCOLOR_TOOLBOX_GRADIENTDARK = -155,
    VSCOLOR_TOOLBOX_GRADIENTLIGHT = -156,
    VSCOLOR_TOOLBOX_HEADINGACCENT = -157,
    VSCOLOR_TOOLBOX_HEADINGBEGIN = -158,
    VSCOLOR_TOOLBOX_HEADINGEND = -159,
    VSCOLOR_TOOLBOX_ICON_HIGHLIGHT = -160,
    VSCOLOR_TOOLBOX_ICON_SHADOW = -161,
    VSCOLOR_TOOLWINDOW_BACKGROUND = -162,
    VSCOLOR_TOOLWINDOW_BORDER = -163,
    VSCOLOR_TOOLWINDOW_BUTTON_DOWN = -164,
    VSCOLOR_TOOLWINDOW_BUTTON_DOWN_BORDER = -165,
    VSCOLOR_TOOLWINDOW_BUTTON_HOVER_ACTIVE = -166,
    VSCOLOR_TOOLWINDOW_BUTTON_HOVER_ACTIVE_BORDER = -167,
    VSCOLOR_TOOLWINDOW_BUTTON_HOVER_INACTIVE = -168,
    VSCOLOR_TOOLWINDOW_BUTTON_HOVER_INACTIVE_BORDER = -169,
    VSCOLOR_TOOLWINDOW_TEXT = -170,
    VSCOLOR_TOOLWINDOW_TAB_SELECTEDTAB = -171,
    VSCOLOR_TOOLWINDOW_TAB_BORDER = -172,
    VSCOLOR_TOOLWINDOW_TAB_GRADIENTBEGIN = -173,
    VSCOLOR_TOOLWINDOW_TAB_GRADIENTEND = -174,
    VSCOLOR_TOOLWINDOW_TAB_TEXT = -175,
    VSCOLOR_TOOLWINDOW_TAB_SELECTEDTEXT = -176,
    VSCOLOR_WIZARD_ORIENTATIONPANEL_BACKGROUND = -177,
    VSCOLOR_WIZARD_ORIENTATIONPANEL_TEXT = -178,
    VSCOLOR_LASTEX = -178
  };

  enum GRADIENTTYPE {
    VSGRADIENT_FILETAB = 1,
    VSGRADIENT_PANEL_BACKGROUND = 2,
    VSGRADIENT_SHELLBACKGROUND = 3,
    VSGRADIENT_TOOLBOX_HEADING = 4,
    VSGRADIENT_TOOLTAB = 5,
    VSGRADIENT_TOOLWIN_ACTIVE_TITLE_BAR = 6,
    VSGRADIENT_TOOLWIN_INACTIVE_TITLE_BAR = 7,
    VSGRADIENT_TOOLWIN_BACKGROUND = 8
  };

  enum VSCURSORTYPE {
    VSCURSOR_APPSTARTING = 1,
    VSCURSOR_COLUMNSPLIT_EW = 2,
    VSCURSOR_COLUMNSPLIT_NS = 3,
    VSCURSOR_CONTROL_COPY = 4,
    VSCURSOR_CONTROL_DELETE = 5,
    VSCURSOR_CONTROL_MOVE = 6,
    VSCURSOR_CROSS = 7,
    VSCURSOR_DRAGDOCUMENT_MOVE = 8,
    VSCURSOR_DRAGDOCUMENT_NOEFFECT = 9,
    VSCURSOR_DRAGSCRAP_COPY = 10,
    VSCURSOR_DRAGSCRAP_MOVE = 11,
    VSCURSOR_DRAGSCRAP_SCROLL = 12,
    VSCURSOR_HAND = 13,
    VSCURSOR_IBEAM = 14,
    VSCURSOR_ISEARCH = 15,
    VSCURSOR_ISEARCH_UP = 16,
    VSCURSOR_MACRO_RECORD_NO = 17,
    VSCURSOR_NO = 18,
    VSCURSOR_NOMOVE_2D = 19,
    VSCURSOR_NOMOVE_HORIZ = 20,
    VSCURSOR_NOMOVE_VERT = 21,
    VSCURSOR_PAN_EAST = 22,
    VSCURSOR_PAN_NE = 23,
    VSCURSOR_PAN_NORTH = 24,
    VSCURSOR_PAN_NW = 25,
    VSCURSOR_PAN_SE = 26,
    VSCURSOR_PAN_SOUTH = 27,
    VSCURSOR_PAN_SW = 28,
    VSCURSOR_PAN_WEST = 29,
    VSCURSOR_POINTER = 30,
    VSCURSOR_POINTER_REVERSE = 31,
    VSCURSOR_SIZE_NS = 32,
    VSCURSOR_SIZE_EW = 33,
    VSCURSOR_SIZE_NWSE = 34,
    VSCURSOR_SIZE_NESW = 35,
    VSCURSOR_SIZE_ALL = 36,
    VSCURSOR_SPLIT_EW = 37,
    VSCURSOR_SPLIT_NS = 38,
    VSCURSOR_UPARROW = 39,
    VSCURSOR_WAIT = 40
  };

  enum BWI_IMAGE_POS {
    BWI_IMAGE_POS_LEFT = 0,
    BWI_IMAGE_POS_RIGHT = 0x1,
    BWI_IMAGE_ONLY = 0x2
  };

  struct __declspec(novtable) IVsUIShell2 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIShell2 ==
    virtual HRESULT __stdcall GetOpenFileNameViaDlgEx(VSOPENFILENAMEW *pOpenFileName, const wchar_t *pszHelpTopic) = 0;
    virtual HRESULT __stdcall GetSaveFileNameViaDlgEx(VSSAVEFILENAMEW *pSaveFileName, const wchar_t *pszHelpTopic) = 0;
    virtual HRESULT __stdcall GetDirectoryViaBrowseDlgEx(VSBROWSEINFOW *pBrowse, const wchar_t *pszHelpTopic, const wchar_t *pszOpenButtonLabel, const wchar_t *pszCeilingDir, VSNSEBROWSEINFOW *pNSEBrowseInfo) = 0;
    virtual HRESULT __stdcall SaveItemsViaDlg(UINT cItems, VSSAVETREEITEM rgSaveItems[]) = 0;
    virtual HRESULT __stdcall GetVSSysColorEx(VSSYSCOLOREX dwSysColIndex, DWORD *pdwRGBval) = 0;
    virtual HRESULT __stdcall CreateGradient(GRADIENTTYPE gradientType, IVsGradient **pGradient) = 0;
    virtual HRESULT __stdcall GetVSCursor(VSCURSORTYPE cursor, HCURSOR *phIcon) = 0;
    virtual HRESULT __stdcall IsAutoRecoverSavingCheckpoints(BOOL *pfARSaving) = 0;
    virtual HRESULT __stdcall VsDialogBoxParam(HINSTANCE hinst, DWORD dwId, DLGPROC lpDialogFunc, LPARAM lp) = 0;
    virtual HRESULT __stdcall CreateIconImageButton(HWND hwnd, HICON hicon, BWI_IMAGE_POS bwiPos, IVsImageButton **ppImageButton) = 0;
    virtual HRESULT __stdcall CreateGlyphImageButton(HWND hwnd, WCHAR chGlyph, int xShift, int yShift, BWI_IMAGE_POS bwiPos, IVsImageButton **ppImageButton) = 0;
  };
#pragma endregion

  // ================================ IVsUIShell3 ================================
#pragma region
  const GUID IID_IVsUIShell3 = { 0x1c763d26, 0x637c, 0x46f8, { 0xa5, 0x5c, 0x6e, 0xcc, 0x84, 0xdf, 0x4e, 0x4f } };
  struct __declspec(novtable) IVsUIShell3 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIShell3 ==
    virtual HRESULT __stdcall ReportErrorInfo2(HRESULT hr, VARIANT_BOOL Suppress) = 0;
    virtual HRESULT __stdcall SuppressRestart(VARIANT_BOOL Suppress) = 0;
  };
#pragma endregion

  // ================================ IVsUIShell4 ================================
#pragma region
  const GUID IID_IVsUIShell4 = { 0xc59cda92, 0xd99d, 0x42da, { 0xb2, 0x21, 0x8e, 0x36, 0xb8, 0xdc, 0x47, 0x8e } };

  enum WINDOWFRAMETYPEFLAGS {
    WINDOWFRAMETYPE_Document = 0x1,
    WINDOWFRAMETYPE_Tool = 0x2,
    WINDOWFRAMETYPE_All = 0xffffff,
    WINDOWFRAMETYPE_Uninitialized = 0x80000000,
    WINDOWFRAMETYPE_AllStates = 0xff000000
  };

  struct __declspec(novtable) IVsUIShell4 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIShell4 ==
    virtual HRESULT __stdcall GetWindowEnum(WINDOWFRAMETYPEFLAGS type, IEnumWindowFrames **ppEnum) = 0;
    virtual HRESULT __stdcall SetupToolbar2(HWND hwnd, IVsToolWindowToolbar *ptwt, IOleCommandTarget *pCmdTarget, IVsToolWindowToolbarHost **pptwth) = 0;
    virtual HRESULT __stdcall SetupToolbar3(IVsWindowFrame *pFrame, IVsToolWindowToolbarHost **pptwth) = 0;
    virtual HRESULT __stdcall CreateToolbarTray(IOleCommandTarget *pCmdTarget, IVsToolbarTrayHost **ppToolbarTrayHost) = 0;
  };
#pragma endregion

  // ================================ IVsUIShell5 ================================
#pragma region
  const GUID IID_IVsUIShell5 = { 0x2b70ea30, 0x51f2, 0x48bb, { 0xab, 0xa8, 0x05, 0x19, 0x46, 0xa3, 0x72, 0x83 } };

  enum THEMEDCOLORTYPE {
    TCT_Background = 0,
    TCT_Foreground = (TCT_Background + 1)
  };

  struct __declspec(novtable) IVsUIShell5 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIShell5 ==
    virtual HRESULT __stdcall GetOpenFileNameViaDlgEx2(VSOPENFILENAMEW *openFileName, const wchar_t *helpTopic, const wchar_t *openButtonLabel) = 0;
    virtual HRESULT __stdcall GetThemedColor(const GUID &colorCategory, const wchar_t *colorName, THEMEDCOLORTYPE colorType, DWORD *colorRgba) = 0;
    virtual HRESULT __stdcall GetKeyBindingScope(const GUID &keyBindingScope, BSTR *pbstrName) = 0;
    virtual HRESULT __stdcall EnumKeyBindingScopes(IVsEnumGuids **ppEnum) = 0;
    virtual HRESULT __stdcall ThemeWindow(HWND hwnd, VARIANT_BOOL *pfThemeApplied) = 0;
    virtual HRESULT __stdcall CreateThemedImageList(HANDLE hImageList, COLORREF crBackground, HANDLE *phThemedImageList) = 0;
    virtual HRESULT __stdcall ThemeDIBits(DWORD dwBitmapLength, BYTE pBitmap[], DWORD dwPixelWidth, DWORD dwPixelHeight, VARIANT_BOOL fIsTopDownBitmap, COLORREF crBackground) = 0;
  };
#pragma endregion

  // ================================ IVsUIShell6 ================================
#pragma region
  const GUID IID_IVsUIShell6 = { 0x7033d7ed, 0x0e98, 0x4c91, { 0x98, 0x81, 0x1d, 0xd8, 0x48, 0x91, 0xd3, 0x78 } };
  struct __declspec(novtable) IVsUIShell6 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIShell5 ==
    virtual HRESULT __stdcall GetOpenFileNameViaDlgEx2(VSOPENFILENAMEW *openFileName, const wchar_t *helpTopic, const wchar_t *openButtonLabel) = 0;
    virtual HRESULT __stdcall GetThemedColor(const GUID &colorCategory, const wchar_t *colorName, THEMEDCOLORTYPE colorType, DWORD *colorRgba) = 0;
    virtual HRESULT __stdcall GetKeyBindingScope(const GUID &keyBindingScope, BSTR *pbstrName) = 0;
    virtual HRESULT __stdcall EnumKeyBindingScopes(IVsEnumGuids **ppEnum) = 0;
    virtual HRESULT __stdcall ThemeWindow(HWND hwnd, VARIANT_BOOL *pfThemeApplied) = 0;
    virtual HRESULT __stdcall CreateThemedImageList(HANDLE hImageList, COLORREF crBackground, HANDLE *phThemedImageList) = 0;
    virtual HRESULT __stdcall ThemeDIBits(DWORD dwBitmapLength, BYTE pBitmap[], DWORD dwPixelWidth, DWORD dwPixelHeight, VARIANT_BOOL fIsTopDownBitmap, COLORREF crBackground) = 0;

    // == IVsUIShell6 ==
    virtual HRESULT __stdcall SetFixedThemeColors(HWND hwnd) = 0;
  };
#pragma endregion

  // ================================ IVsUIShell7 ================================
#pragma region
  const GUID IID_IVsUIShell7 = { 0xc5ffe6b0, 0xc6e0, 0x45a4, { 0x9f, 0xe2, 0xda, 0x5d, 0xeb, 0x84, 0xe5, 0x9d } };
  struct __declspec(novtable) IVsUIShell7 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIShell7 ==
    virtual HRESULT __stdcall AdviseWindowFrameEvents(IVsWindowFrameEvents *eventSink, DWORD *cookie) = 0;
    virtual HRESULT __stdcall UnadviseWindowFrameEvents(DWORD cookie) = 0;
  };
#pragma endregion

  // ================================ IVsBuildPropertyStorage ================================
#pragma region
  const GUID IID_IVsBuildPropertyStorage = { 0xe7355fdf, 0xa118, 0x48f5, { 0x96, 0x55, 0x7e, 0xfd, 0x9d, 0x2d, 0xc3, 0x52 } };

  enum PersistStorageType {
    PST_PROJECT_FILE = 1,
    PST_USER_FILE = 2
  };

  struct __declspec(novtable) IVsBuildPropertyStorage {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsBuildPropertyStorage ==
    virtual HRESULT __stdcall GetPropertyValue(const wchar_t *pszPropName, const wchar_t *pszConfigName, PersistStorageType storage, BSTR *pbstrPropValue) = 0;
    virtual HRESULT __stdcall SetPropertyValue(const wchar_t *pszPropName, const wchar_t *pszConfigName, PersistStorageType storage, const wchar_t *pszPropValue) = 0;
    virtual HRESULT __stdcall RemoveProperty(const wchar_t *pszPropName, const wchar_t *pszConfigName, PersistStorageType storage) = 0;
    virtual HRESULT __stdcall GetItemAttribute(DWORD item, const wchar_t *pszAttributeName, BSTR *pbstrAttributeValue) = 0;
    virtual HRESULT __stdcall SetItemAttribute(DWORD item, const wchar_t *pszAttributeName, const wchar_t *pszAttributeValue) = 0;
  };
#pragma endregion

  // ================================ IVsBuildPropertyStorage2 ================================
#pragma region
  const GUID IID_IVsBuildPropertyStorage2 = { 0x3b175ac0, 0xf7e2, 0x4187, { 0x80, 0xa0, 0xa7, 0x3c, 0x39, 0x31, 0x3c, 0x49 } };
  struct __declspec(novtable) IVsBuildPropertyStorage2 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsBuildPropertyStorage2 ==
    virtual HRESULT __stdcall SetPropertyValueEx(const wchar_t *pszPropName, const wchar_t *pszPropertyGroupCondition, PersistStorageType storage, const wchar_t *pszPropValue) = 0;
  };
#pragma endregion

  // ================================ IVsBuildPropertyStorage3 ================================
#pragma region
  const GUID IID_IVsBuildPropertyStorage3 = { 0x9669894b, 0x8698, 0x4e4a, { 0xbf, 0x06, 0xae, 0xca, 0x45, 0x55, 0x9c, 0x36 } };
  struct __declspec(novtable) IVsBuildPropertyStorage3 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsBuildPropertyStorage3 ==
    
    // Subscribes to build property storage events.
    virtual HRESULT __stdcall AdviseEvents(IVsBuildPropertyStorageEvents *pSink, DWORD *pdwCookie) = 0;
    // Unsubscribes from build property storage events.
    virtual HRESULT __stdcall UnadviseEvents(DWORD dwCookie) = 0;
  };
#pragma endregion

  // ================================ IVsHierarchy ================================
#pragma region
  const GUID IID_IVsHierarchy = { 0x59b2d1d0, 0x5db0, 0x4f9f, { 0x96, 0x09, 0x13, 0xf0, 0x16, 0x85, 0x16, 0xd6 } };
  
  // A project is a tree of nodes (this "hierarchy"), each corresponding to a project item with associated properties.
  struct __declspec(novtable) IVsHierarchy {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsHierarchy ==
    virtual HRESULT __stdcall SetSite(IServiceProvider *pSP) = 0;
    virtual HRESULT __stdcall GetSite(IServiceProvider **ppSP) = 0;
    virtual HRESULT __stdcall QueryClose(BOOL *pfCanClose) = 0;
    virtual HRESULT __stdcall Close() = 0;
    virtual HRESULT __stdcall GetGuidProperty(DWORD itemid, VSHPROPID propid, GUID *pguid) = 0;
    virtual HRESULT __stdcall SetGuidProperty(DWORD itemid, VSHPROPID propid, const GUID &rguid) = 0;
    virtual HRESULT __stdcall GetProperty(DWORD itemid, VSHPROPID propid, VARIANT *pvar) = 0;
    virtual HRESULT __stdcall SetProperty(DWORD itemid, VSHPROPID propid, VARIANT var) = 0;
    virtual HRESULT __stdcall GetNestedHierarchy(DWORD itemid, const IID &iidHierarchyNested, void **ppHierarchyNested, DWORD *pitemidNested) = 0;
    virtual HRESULT __stdcall GetCanonicalName(DWORD itemid, BSTR *pbstrName) = 0;
    virtual HRESULT __stdcall ParseCanonicalName(const wchar_t *pszName, DWORD *pitemid) = 0;
    UNUSED_FUNC(0)
    virtual HRESULT __stdcall AdviseHierarchyEvents(IVsHierarchyEvents *pEventSink, DWORD *pdwCookie) = 0;
    virtual HRESULT __stdcall UnadviseHierarchyEvents(DWORD dwCookie) = 0;
    UNUSED_FUNC(1)
    UNUSED_FUNC(2)
    UNUSED_FUNC(3)
    UNUSED_FUNC(4)
  };
#pragma endregion

  // ================================ IVsHierarchy2 ================================
#pragma region
  const GUID IID_IVsHierarchy2 = { 0x41e82a90, 0x2488, 0x4595, { 0x96, 0x61, 0x39, 0x6a, 0x1c, 0x54, 0xb7, 0x2e } };
  struct __declspec(novtable) IVsHierarchy2 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsHierarchy2 ==
    virtual HRESULT __stdcall GetGuidProperties(DWORD itemid, int count, VSHPROPID propids[], GUID rgGuids[], int results[]) = 0;
    virtual HRESULT __stdcall GetProperties(DWORD itemid, int count, VSHPROPID propids[], VARIANT vars[], int results[]) = 0;
  };
#pragma endregion

  // ================================ IVsBuildPropertyStorageEvents ================================
#pragma region
  const GUID IID_IVsBuildPropertyStorageEvents = { 0x2c6c93fd, 0xc88f, 0x45ac, { 0xac, 0x2b, 0x39, 0xe9, 0x11, 0x76, 0xf8, 0x94 } };
  struct __declspec(novtable) IVsBuildPropertyStorageEvents {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsBuildPropertyStorageEvents ==
    virtual HRESULT __stdcall OnAfterItemsChanged(int cItems, const LPCOLESTR rgpszItemFullPaths[]) = 0;
  };
#pragma endregion

  // ================================ IVsHierarchyEvents ================================
#pragma region
  const GUID IID_IVsHierarchyEvents = { 0x6ddd8dc3, 0x32b2, 0x4bf1, { 0xa1, 0xe1, 0xb6, 0xda, 0x40, 0x52, 0x6d, 0x1e } };

  // Verified working: OnItemAdded, OnItemDeleted, OnInvalidateItems
  // OnPropertyChanged doesn't do shit 

  // Provides notification of changes to a hierarchy.
  struct __declspec(novtable) IVsHierarchyEvents {
    // == IUnknown methods ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsHierarchyEvents methods ==

    // Notifies the client when an item is added to the hierarchy.
    // itemidParent: The ID of the parent item.
    // itemidSiblingPrev: The ID of the previous sibling item.
    // itemidAdded: The ID of the added item.
    virtual HRESULT __stdcall OnItemAdded(DWORD itemidParent, DWORD itemidSiblingPrev, DWORD itemidAdded) = 0;

    // Notifies the client when items are appended to the hierarchy.
    // itemidParent: The ID of the parent item.
    virtual HRESULT __stdcall OnItemsAppended(DWORD itemidParent) = 0;

    // Notifies the client when an item is deleted from the hierarchy.
    // itemid: The ID of the deleted item.
    virtual HRESULT __stdcall OnItemDeleted(DWORD itemid) = 0;

    // Notifies the client when a property of an item in the hierarchy changes.
    // itemid: The ID of the item whose property changed.
    // propid: The ID of the property that changed.
    // flags: Additional flags indicating the nature of the change.
    virtual HRESULT __stdcall OnPropertyChanged(DWORD itemid, VSHPROPID propid, DWORD flags) = 0;

    // Notifies the client when items in the hierarchy need to be invalidated.
    // itemidParent: The ID of the parent item whose children need to be invalidated.
    virtual HRESULT __stdcall OnInvalidateItems(DWORD itemidParent) = 0;

    // Notifies the client when an icon in the hierarchy needs to be invalidated.
    // hicon: The handle to the icon to invalidate.
    virtual HRESULT __stdcall OnInvalidateIcon(HICON hicon) = 0;
  };
#pragma endregion

//
// Older Visual Studio SDK ("DTE" based)
//
#pragma region DTE

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
    virtual HRESULT __stdcall Progress(VARIANT_BOOL InProgress, BSTR Label, long AmountCompleted = 0, long Total = 0) = 0;
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
    virtual HRESULT __stdcall Save(BSTR FileName) = 0;
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
    virtual HRESULT __stdcall get_Solution_CLR(Solution **ppSolution) = 0; // ORIGINAL: get_Solution
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

    // == Unofficial utilites == 
    HRESULT get_Solution(_Solution **result) {
      IUnknown *sln_CLR = 0;
      HRESULT hr = get_Solution_CLR((Solution **)&sln_CLR);
      if (FAILED(hr)) {
        return hr;
      }

      // Actual sln
      return sln_CLR->QueryInterface(IID__Solution, (void **)result);
    }
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
    virtual HRESULT __stdcall Save(BSTR FileName) = 0;
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
    virtual HRESULT __stdcall get_DTE(DTE **ppDTE) = 0;
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

  // ================================ IServiceProvider ================================
#pragma region
  const GUID IID_IServiceProvider = { 0x6d5140c1, 0x7436, 0x11ce, { 0x80, 0x34, 0x00, 0xaa, 0x00, 0x60, 0x09, 0xfa } };

  struct IServiceProvider {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IServiceProvider ==
    virtual HRESULT __stdcall QueryService(const GUID &guidService, const IID &riid, void **ppvObject) = 0;

    // == Overloaded QueryService methods ==
    HRESULT QueryService(IVsSolution **result) {
      return QueryService(IID_IVsSolution, IID_IVsSolution, (void **)result);
    }

    HRESULT QueryService(IVsSolutionBuildManager **result) {
      return QueryService(IID_IVsSolutionBuildManager, IID_IVsSolutionBuildManager, (void **)result);
    }

    HRESULT QueryService(IVsUIShell **result) {
      return QueryService(IID_IVsUIShell, IID_IVsUIShell, (void **)result);
    }

    HRESULT QueryService(_DTE **result) {
      return QueryService(IID__DTE, IID__DTE, (void **)result);
    }
  };
#pragma endregion

#pragma endregion
}


#undef UNUSED_FUNC