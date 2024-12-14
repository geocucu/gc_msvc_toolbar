#pragma once
#include <vsshell.h>
#include <vsshell110.h>


namespace vsix {
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

  // ================================ IVsUpdateSolutionEvents ================================
  const GUID IID_IVsUpdateSolutionEvents = { 0xa9f86308, 0x5ea7, 0x485d, { 0xba, 0xb8, 0xe8, 0x98, 0x9c, 0x3c, 0xfb, 0xdc } };
  struct __declspec(novtable) IVsUpdateSolutionEvents {
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
  };

  // ================================ IVsUpdateSolutionEvents2 ================================
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

  // ================================ IVsUpdateSolutionEvents3 ================================
  const GUID IID_IVsUpdateSolutionEvents3 = { 0x40025c28, 0x3303, 0x42ca, { 0xbe, 0xd8, 0x0f, 0x3b, 0xd8, 0x56, 0xbd, 0x6d } };
  struct __declspec(novtable) IVsUpdateSolutionEvents3 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUpdateSolutionEvents3 ==
    virtual HRESULT __stdcall OnBeforeActiveSolutionCfgChange(IVsCfg *pOldActiveSlnCfg, IVsCfg *pNewActiveSlnCfg) = 0;
    virtual HRESULT __stdcall OnAfterActiveSolutionCfgChange(IVsCfg *pOldActiveSlnCfg, IVsCfg *pNewActiveSlnCfg) = 0;
  };

  // ================================ IVsUpdateSolutionEvents4 ================================
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

  // ================================ IVsUpdateSolutionEvents5 ================================
  const GUID IID_IVsUpdateSolutionEvents5 = { 0x95498691, 0xcb06, 0x4bc1, { 0x8a, 0x83, 0xf8, 0x4c, 0x6d, 0x56, 0x5e, 0x21 } };
  struct __declspec(novtable) IVsUpdateSolutionEvents5 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUpdateSolutionEvents5 ==
    virtual HRESULT __stdcall UpdateSolution_QueryDelayBuildAction(DWORD dwAction, IVsTask **pDelayTask) = 0;
  };

  // ================================ IVsSolutionEvents ================================
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

  // ================================ IVsSolutionEvents2 ================================
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

  // ================================ IVsSolutionEvents3 ================================
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

  // ================================ IVsSolutionEvents4 ================================
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

  // ================================ IVsSolutionEvents5 ================================
  const GUID IID_IVsSolutionEvents5 = { 0xaf530689, 0x9987, 0x48be, { 0xaf, 0x20, 0xd9, 0x39, 0x2a, 0x9c, 0x67, 0xff } };
  struct __declspec(novtable) IVsSolutionEvents5 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionEvents5 ==
    virtual HRESULT __stdcall OnBeforeOpenProject(const GUID &guidProjectID, const GUID &guidProjectType, const wchar_t *pszFileName) = 0;
  };

  // ================================ IVsSolutionEvents6 ================================
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

  // ================================ IVsSolutionEvents7 ================================
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

  // ================================ IVsSolutionEvents8 ================================
  const GUID IID_IVsSolutionEvents8 = { 0x36b55a6c, 0x8da0, 0x428f, { 0x82, 0x8c, 0x80, 0xde, 0x91, 0x05, 0xf9, 0xc1 } };
  struct __declspec(novtable) IVsSolutionEvents8 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionEvents8 ==
    virtual HRESULT __stdcall OnAfterRenameSolution(const wchar_t *old_name, const wchar_t *new_name) = 0;
  };

  // ================================ IVsSolutionBuildManager ================================
  const GUID IID_IVsSolutionBuildManager = { 0x93e969d6, 0x1aa0, 0x455f, { 0xb2, 0x08, 0x6e, 0xd3, 0xc8, 0x2b, 0x5c, 0x58 } };
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

  // ================================ IVsSolutionBuildManager2 ================================
  const GUID IID_IVsSolutionBuildManager2 = { 0x80353f58, 0xf2a3, 0x47b8, { 0xb2, 0xdf, 0x04, 0x75, 0xe0, 0x7b, 0xb1, 0xc6 } };
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

  // ================================ IVsSolutionBuildManager3 ================================
  const GUID IID_IVsSolutionBuildManager3 = { 0xb6ea87ed, 0xc498, 0x4484, { 0x81, 0xac, 0x0b, 0xed, 0x18, 0x7e, 0x28, 0xe6 } };
  struct __declspec(novtable) IVsSolutionBuildManager3 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionBuildManager3 ==
    virtual HRESULT __stdcall AdviseUpdateSolutionEvents3(IVsUpdateSolutionEvents3 *pIVsUpdateSolutionEvents3, DWORD *pdwCookie) = 0;
    virtual HRESULT __stdcall UnadviseUpdateSolutionEvents3(DWORD dwCookie) = 0;
    virtual HRESULT __stdcall AreProjectsUpToDate(DWORD dwOptions) = 0;
    virtual HRESULT __stdcall HasHierarchyChangedSinceLastDTEE() = 0;
    virtual HRESULT __stdcall QueryBuildManagerBusyEx(DWORD *pdwBuildManagerOperation) = 0;
  };

  // ================================ IVsSolutionBuildManager4 ================================
  const GUID IID_IVsSolutionBuildManager4 = { 0x2c07342b, 0xba98, 0x4235, { 0x98, 0x3c, 0x86, 0x38, 0x39, 0x1a, 0x42, 0x0a } };
  struct __declspec(novtable) IVsSolutionBuildManager4 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionBuildManager4 ==
    virtual HRESULT __stdcall UpdateProjectDependencies(IVsHierarchy *pHier) = 0;
  };

  // ================================ IVsSolutionBuildManager5 ================================
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

  // ================================ IVsSolutionBuildManager6 ================================
  const GUID IID_IVsSolutionBuildManager6 = { 0x61aa4ea9, 0x018f, 0x4394, { 0xad, 0x31, 0x1e, 0x76, 0xd1, 0xbf, 0x80, 0xc8 } };
  struct __declspec(novtable) IVsSolutionBuildManager6 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsSolutionBuildManager6 ==
    virtual HRESULT __stdcall AdviseUpdateSolutionEventsEx(const GUID &guidActivityId, IUnknown *pSink, DWORD *pdwCookie) = 0;
  };

  // ================================ IVsSolution ================================
  const GUID IID_IVsSolution = { 0x7f7cd0db, 0x91ef, 0x49dc, { 0x9f, 0xa9, 0x02, 0xd1, 0x28, 0x51, 0x5d, 0xd4 } };
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

  // ================================ IVsSolution2 ================================
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

  // ================================ IVsSolution3 ================================
  const GUID IID_IVsSolution3 = { 0x58dcf7bf, 0xf14e, 0x43ec, { 0xa7, 0xb2, 0x9f, 0x78, 0xed, 0xd0, 0x64, 0x18 } };
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

  // ================================ IVsSolution4 ================================
  const GUID IID_IVsSolution4 = { 0xd2fb5b25, 0xeaf0, 0x4be9, { 0x8e, 0x9b, 0xf2, 0xc6, 0x62, 0xab, 0x98, 0x26 } };
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

  // ================================ IVsSolution5 ================================
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

  // ================================ IVsSolution6 ================================
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

  // ================================ IVsSolution7 ================================
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

  // ================================ IVsSolution8 ================================
  const GUID IID_IVsSolution8 = { 0x51a3a58a, 0xe65e, 0x46f8, { 0xa6, 0x3c, 0x49, 0x66, 0x5a, 0x67, 0x50, 0x16 } };
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

  // ================================ IVsUIShell ================================
  const GUID IID_IVsUIShell = { 0xb61fc35b, 0xeebf, 0x4dec, { 0xbf, 0xf1, 0x28, 0xa2, 0xdd, 0x43, 0xc3, 0x8f } };
  struct __declspec(novtable) IVsUIShell {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIShell ==
    virtual HRESULT __stdcall GetToolWindowEnum(IEnumWindowFrames **ppEnum) = 0;
    virtual HRESULT __stdcall GetDocumentWindowEnum(IEnumWindowFrames **ppEnum) = 0;
    virtual HRESULT __stdcall FindToolWindow(VSFINDTOOLWIN grfFTW, const GUID &rguidPersistenceSlot, IVsWindowFrame **ppWindowFrame) = 0;
    virtual HRESULT __stdcall CreateToolWindow(VSCREATETOOLWIN grfCTW, DWORD dwToolWindowId, IUnknown *punkTool, REFCLSID rclsidTool, const GUID &rguidPersistenceSlot, const GUID &rguidAutoActivate, IServiceProvider *pSP, const wchar_t *pszCaption, BOOL *pfDefaultPosition, IVsWindowFrame **ppWindowFrame) = 0;
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
    virtual HRESULT __stdcall ShowContextMenu(DWORD dwCompRole, REFCLSID rclsidActive, LONG nMenuId, REFPOINTS pos, IOleCommandTarget *pCmdTrgtActive) = 0;
    virtual HRESULT __stdcall ShowMessageBox(DWORD dwCompRole, REFCLSID rclsidComp, LPOLESTR pszTitle, LPOLESTR pszText, LPOLESTR pszHelpFile, DWORD dwHelpContextID, OLEMSGBUTTON msgbtn, OLEMSGDEFBUTTON msgdefbtn, OLEMSGICON msgicon, BOOL fSysAlert, LONG *pnResult) = 0;
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

  // ================================ IVsUIShell2 ================================
  const GUID IID_IVsUIShell2 = { 0x4e6b6ef9, 0x8e3d, 0x4756, { 0x99, 0xe9, 0x11, 0x92, 0xba, 0xad, 0x54, 0x96 } };
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

  // ================================ IVsUIShell3 ================================
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

  // ================================ IVsUIShell4 ================================
  const GUID IID_IVsUIShell4 = { 0xc59cda92, 0xd99d, 0x42da, { 0xb2, 0x21, 0x8e, 0x36, 0xb8, 0xdc, 0x47, 0x8e } };
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

  // ================================ IVsUIShell5 ================================
  const GUID IID_IVsUIShell5 = { 0x2b70ea30, 0x51f2, 0x48bb, { 0xab, 0xa8, 0x05, 0x19, 0x46, 0xa3, 0x72, 0x83 } };
  struct __declspec(novtable) IVsUIShell5 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIShell5 ==
    virtual HRESULT __stdcall GetOpenFileNameViaDlgEx2(VSOPENFILENAMEW *openFileName, const wchar_t *helpTopic, const wchar_t *openButtonLabel) = 0;
    virtual HRESULT __stdcall GetThemedColor(const GUID &colorCategory, const wchar_t *colorName, THEMEDCOLORTYPE colorType, VS_RGBA *colorRgba) = 0;
    virtual HRESULT __stdcall GetKeyBindingScope(const GUID &keyBindingScope, BSTR *pbstrName) = 0;
    virtual HRESULT __stdcall EnumKeyBindingScopes(IVsEnumGuids **ppEnum) = 0;
    virtual HRESULT __stdcall ThemeWindow(HWND hwnd, VARIANT_BOOL *pfThemeApplied) = 0;
    virtual HRESULT __stdcall CreateThemedImageList(HANDLE hImageList, COLORREF crBackground, HANDLE *phThemedImageList) = 0;
    virtual HRESULT __stdcall ThemeDIBits(DWORD dwBitmapLength, BYTE pBitmap[], DWORD dwPixelWidth, DWORD dwPixelHeight, VARIANT_BOOL fIsTopDownBitmap, COLORREF crBackground) = 0;
  };

  // ================================ IVsUIShell6 ================================
  const GUID IID_IVsUIShell6 = { 0x7033d7ed, 0x0e98, 0x4c91, { 0x98, 0x81, 0x1d, 0xd8, 0x48, 0x91, 0xd3, 0x78 } };
  struct __declspec(novtable) IVsUIShell6 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsUIShell5 ==
    virtual HRESULT __stdcall GetOpenFileNameViaDlgEx2(VSOPENFILENAMEW *openFileName, const wchar_t *helpTopic, const wchar_t *openButtonLabel) = 0;
    virtual HRESULT __stdcall GetThemedColor(const GUID &colorCategory, const wchar_t *colorName, THEMEDCOLORTYPE colorType, VS_RGBA *colorRgba) = 0;
    virtual HRESULT __stdcall GetKeyBindingScope(const GUID &keyBindingScope, BSTR *pbstrName) = 0;
    virtual HRESULT __stdcall EnumKeyBindingScopes(IVsEnumGuids **ppEnum) = 0;
    virtual HRESULT __stdcall ThemeWindow(HWND hwnd, VARIANT_BOOL *pfThemeApplied) = 0;
    virtual HRESULT __stdcall CreateThemedImageList(HANDLE hImageList, COLORREF crBackground, HANDLE *phThemedImageList) = 0;
    virtual HRESULT __stdcall ThemeDIBits(DWORD dwBitmapLength, BYTE pBitmap[], DWORD dwPixelWidth, DWORD dwPixelHeight, VARIANT_BOOL fIsTopDownBitmap, COLORREF crBackground) = 0;

    // == IVsUIShell6 ==
    virtual HRESULT __stdcall SetFixedThemeColors(HWND hwnd) = 0;
  };

  // ================================ IVsUIShell7 ================================
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

  // ================================ IVsBuildPropertyStorage ================================
  const GUID IID_IVsBuildPropertyStorage = { 0xe7355fdf, 0xa118, 0x48f5, { 0x96, 0x55, 0x7e, 0xfd, 0x9d, 0x2d, 0xc3, 0x52 } };
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

  // ================================ IVsBuildPropertyStorage2 ================================
  const GUID IID_IVsBuildPropertyStorage2 = { 0x3b175ac0, 0xf7e2, 0x4187, { 0x80, 0xa0, 0xa7, 0x3c, 0x39, 0x31, 0x3c, 0x49 } };
  struct __declspec(novtable) IVsBuildPropertyStorage2 {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsBuildPropertyStorage2 ==
    virtual HRESULT __stdcall SetPropertyValueEx(const wchar_t *pszPropName, const wchar_t *pszPropertyGroupCondition, PersistStorageType storage, const wchar_t *pszPropValue) = 0;
  };

  // ================================ IVsBuildPropertyStorage3 ================================
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

  // ================================ IVsHierarchy ================================
  // A project is a tree of nodes (this "hierarchy"), each corresponding to a project item with associated properties.
  const GUID IID_IVsHierarchy = { 0x59b2d1d0, 0x5db0, 0x4f9f, { 0x96, 0x09, 0x13, 0xf0, 0x16, 0x85, 0x16, 0xd6 } };
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
    virtual HRESULT __stdcall Unused0() = 0;
    virtual HRESULT __stdcall AdviseHierarchyEvents(IVsHierarchyEvents *pEventSink, DWORD *pdwCookie) = 0;
    virtual HRESULT __stdcall UnadviseHierarchyEvents(DWORD dwCookie) = 0;
    virtual HRESULT __stdcall Unused1() = 0;
    virtual HRESULT __stdcall Unused2() = 0;
    virtual HRESULT __stdcall Unused3() = 0;
    virtual HRESULT __stdcall Unused4() = 0;
  };

  // ================================ IVsHierarchy2 ================================
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

  // ================================ IVsBuildPropertyStorageEvents ================================
  const GUID IID_IVsBuildPropertyStorageEvents = { 0x2c6c93fd, 0xc88f, 0x45ac, { 0xac, 0x2b, 0x39, 0xe9, 0x11, 0x76, 0xf8, 0x94 } };
  struct __declspec(novtable) IVsBuildPropertyStorageEvents {
    // == IUnknown ==
    virtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppvObject) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;

    // == IVsBuildPropertyStorageEvents ==
    virtual HRESULT __stdcall OnAfterItemsChanged(int cItems, const LPCOLESTR rgpszItemFullPaths[]) = 0;
  };
}