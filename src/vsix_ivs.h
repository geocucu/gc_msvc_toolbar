#pragma once

//
// Newer Visual Studio SDK ("IVs*" based)
//

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
#include <VSShellInterfaces.h>
#include <vsshell.h>
#include <vsshell110.h>

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


namespace vsix {
#pragma region Forward Declarations
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
      VSCOOKIE *pdwCookie
    ) = 0;

    // Unadvises the client of changes to project documents.
    // dwCookie: The event cookie to unadvise.
    virtual HRESULT __stdcall UnadviseTrackProjectDocumentsEvents(
      VSCOOKIE dwCookie
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
    virtual HRESULT __stdcall EnableStatementCompletion(BOOL fEnable, CharIndex iStartIndex, CharIndex iEndIndex, IVsTextView *pTextView) = 0;

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

    virtual HRESULT __stdcall GetUIContextState(REFGUID uiContext, UIContextState *pState) = 0;

    virtual HRESULT __stdcall SetUIContextState(REFGUID uiContext, VARIANT_BOOL isActive) = 0;

    // Advises for change events for all UI contexts.
    // Safe to access from any thread, though do be aware that if this method is called off the UI thread while a 
    // context is actively being set on the UI thread, the registered callback may miss the change notification.
    // callback: The callback to invoke when a UI context value changes.
    // pCookie: A cookie representing your subscription.
    virtual HRESULT __stdcall AdviseUIContextEvents(IVsUIContextEvents *callback, VSCOOKIE *pCookie) = 0;

    virtual HRESULT __stdcall AdviseSpecificUIContextEvents(IVsUIContextEvents *callback, REFGUID uiContext, VSCOOKIE *pCookie) = 0;

    // Safe to access from any thread.
    virtual HRESULT __stdcall UnadviseUIContextEvents(VSCOOKIE cookie) = 0;
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
  // To monitor these events, implement the interface and use it as an argument to IVsSolutionBuildManager3.AdviseUpdateSolutionEvents3
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
#pragma endregion

  // ================================ IVsUIShell2 ================================
#pragma region
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
    virtual HRESULT __stdcall GetThemedColor(const GUID &colorCategory, const wchar_t *colorName, THEMEDCOLORTYPE colorType, VS_RGBA *colorRgba) = 0;
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
    virtual HRESULT __stdcall OnItemAdded(VSITEMID itemidParent, VSITEMID itemidSiblingPrev, VSITEMID itemidAdded) = 0;

    // Notifies the client when items are appended to the hierarchy.
    // itemidParent: The ID of the parent item.
    virtual HRESULT __stdcall OnItemsAppended(VSITEMID itemidParent) = 0;

    // Notifies the client when an item is deleted from the hierarchy.
    // itemid: The ID of the deleted item.
    virtual HRESULT __stdcall OnItemDeleted(VSITEMID itemid) = 0;

    // Notifies the client when a property of an item in the hierarchy changes.
    // itemid: The ID of the item whose property changed.
    // propid: The ID of the property that changed.
    // flags: Additional flags indicating the nature of the change.
    virtual HRESULT __stdcall OnPropertyChanged(VSITEMID itemid, VSHPROPID propid, DWORD flags) = 0;

    // Notifies the client when items in the hierarchy need to be invalidated.
    // itemidParent: The ID of the parent item whose children need to be invalidated.
    virtual HRESULT __stdcall OnInvalidateItems(VSITEMID itemidParent) = 0;

    // Notifies the client when an icon in the hierarchy needs to be invalidated.
    // hicon: The handle to the icon to invalidate.
    virtual HRESULT __stdcall OnInvalidateIcon(HICON hicon) = 0;
  };
#pragma endregion

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
}

#undef UNUSED_FUNC