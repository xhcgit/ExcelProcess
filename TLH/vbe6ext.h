﻿// Created by Microsoft (R) C/C++ Compiler Version 11.00.61030.0 (5be78d6d).
//
// c:\users\huance.xu\desktop\g1.5-excel\source\cps___win32_releaseu\vbe6ext.tlh
//
// C++ source equivalent of Win32 type library ..\\ExcelProcess\\VBE6EXT.OLB
// compiler-generated file created 12/16/16 at 15:47:30 - DO NOT EDIT!

//
// Cross-referenced type libraries:
//
//

#pragma once
#pragma pack(push, 8)

#include "..\comdef.h"

namespace VBIDE {

//
// Forward references and typedefs
//

struct __declspec(uuid("0002e157-0000-0000-c000-000000000046"))
/* LIBID */ __VBIDE;
struct __declspec(uuid("0002e158-0000-0000-c000-000000000046"))
/* dual interface */ Application;
enum vbextFileTypes;
struct __declspec(uuid("0002e166-0000-0000-c000-000000000046"))
/* dual interface */ VBE;
enum vbext_WindowType;
enum vbext_WindowState;
struct __declspec(uuid("0002e16b-0000-0000-c000-000000000046"))
/* dual interface */ Window;
struct __declspec(uuid("0002e16a-0000-0000-c000-000000000046"))
/* dual interface */ _Windows_old;
struct __declspec(uuid("f57b7ed0-d8ab-11d1-85df-00c04f98f42c"))
/* dual interface */ _Windows;
struct /* coclass */ Windows;
struct __declspec(uuid("0002e16c-0000-0000-c000-000000000046"))
/* dual interface */ _LinkedWindows;
struct /* coclass */ LinkedWindows;
struct __declspec(uuid("0002e167-0000-0000-c000-000000000046"))
/* dual interface */ Events;
struct __declspec(uuid("0002e113-0000-0000-c000-000000000046"))
/* interface */ _VBProjectsEvents;
struct __declspec(uuid("0002e103-0000-0000-c000-000000000046"))
/* dispinterface */ _dispVBProjectsEvents;
struct __declspec(uuid("0002e115-0000-0000-c000-000000000046"))
/* interface */ _VBComponentsEvents;
struct __declspec(uuid("0002e116-0000-0000-c000-000000000046"))
/* dispinterface */ _dispVBComponentsEvents;
struct __declspec(uuid("0002e11a-0000-0000-c000-000000000046"))
/* interface */ _ReferencesEvents;
struct __declspec(uuid("0002e118-0000-0000-c000-000000000046"))
/* dispinterface */ _dispReferencesEvents;
struct /* coclass */ ReferencesEvents;
struct __declspec(uuid("0002e130-0000-0000-c000-000000000046"))
/* interface */ _CommandBarControlEvents;
struct __declspec(uuid("0002e131-0000-0000-c000-000000000046"))
/* dispinterface */ _dispCommandBarControlEvents;
struct /* coclass */ CommandBarEvents;
struct __declspec(uuid("0002e159-0000-0000-c000-000000000046"))
/* dual interface */ _ProjectTemplate;
struct /* coclass */ ProjectTemplate;
enum vbext_ProjectType;
enum vbext_ProjectProtection;
enum vbext_VBAMode;
struct __declspec(uuid("0002e160-0000-0000-c000-000000000046"))
/* dual interface */ _VBProject_Old;
struct __declspec(uuid("eee00915-e393-11d1-bb03-00c04fb6c4a6"))
/* dual interface */ _VBProject;
struct /* coclass */ VBProject;
struct __declspec(uuid("0002e165-0000-0000-c000-000000000046"))
/* dual interface */ _VBProjects_Old;
struct __declspec(uuid("eee00919-e393-11d1-bb03-00c04fb6c4a6"))
/* dual interface */ _VBProjects;
struct /* coclass */ VBProjects;
struct __declspec(uuid("be39f3d4-1b13-11d0-887f-00a0c90f2744"))
/* dual interface */ SelectedComponents;
enum vbext_ComponentType;
struct __declspec(uuid("0002e161-0000-0000-c000-000000000046"))
/* dual interface */ _Components;
struct /* coclass */ Components;
struct __declspec(uuid("0002e162-0000-0000-c000-000000000046"))
/* dual interface */ _VBComponents_Old;
struct __declspec(uuid("eee0091c-e393-11d1-bb03-00c04fb6c4a6"))
/* dual interface */ _VBComponents;
struct /* coclass */ VBComponents;
struct __declspec(uuid("0002e163-0000-0000-c000-000000000046"))
/* dual interface */ _Component;
struct /* coclass */ Component;
struct __declspec(uuid("0002e164-0000-0000-c000-000000000046"))
/* dual interface */ _VBComponent_Old;
struct __declspec(uuid("eee00921-e393-11d1-bb03-00c04fb6c4a6"))
/* dual interface */ _VBComponent;
struct /* coclass */ VBComponent;
struct __declspec(uuid("0002e18c-0000-0000-c000-000000000046"))
/* dual interface */ Property;
struct __declspec(uuid("0002e188-0000-0000-c000-000000000046"))
/* dual interface */ _Properties;
struct /* coclass */ Properties;
struct __declspec(uuid("da936b62-ac8b-11d1-b6e5-00a0c90f2744"))
/* dual interface */ _AddIns;
struct /* coclass */ Addins;
struct __declspec(uuid("da936b64-ac8b-11d1-b6e5-00a0c90f2744"))
/* dual interface */ AddIn;
enum vbext_ProcKind;
struct __declspec(uuid("0002e16e-0000-0000-c000-000000000046"))
/* dual interface */ _CodeModule;
struct /* coclass */ CodeModule;
struct __declspec(uuid("0002e172-0000-0000-c000-000000000046"))
/* dual interface */ _CodePanes;
struct /* coclass */ CodePanes;
enum vbext_CodePaneview;
struct __declspec(uuid("0002e176-0000-0000-c000-000000000046"))
/* dual interface */ _CodePane;
struct /* coclass */ CodePane;
struct __declspec(uuid("0002e17a-0000-0000-c000-000000000046"))
/* dual interface */ _References;
enum vbext_RefKind;
struct __declspec(uuid("0002e17e-0000-0000-c000-000000000046"))
/* dual interface */ Reference;
struct __declspec(uuid("cdde3804-2064-11cf-867f-00aa005ff34a"))
/* dispinterface */ _dispReferences_Events;
struct /* coclass */ References;

//
// Smart pointer typedef declarations
//

_COM_SMARTPTR_TYPEDEF(Application, __uuidof(Application));
_COM_SMARTPTR_TYPEDEF(_VBProjectsEvents, __uuidof(_VBProjectsEvents));
_COM_SMARTPTR_TYPEDEF(_dispVBProjectsEvents, __uuidof(_dispVBProjectsEvents));
_COM_SMARTPTR_TYPEDEF(_VBComponentsEvents, __uuidof(_VBComponentsEvents));
_COM_SMARTPTR_TYPEDEF(_dispVBComponentsEvents, __uuidof(_dispVBComponentsEvents));
_COM_SMARTPTR_TYPEDEF(_ReferencesEvents, __uuidof(_ReferencesEvents));
_COM_SMARTPTR_TYPEDEF(_dispReferencesEvents, __uuidof(_dispReferencesEvents));
_COM_SMARTPTR_TYPEDEF(_CommandBarControlEvents, __uuidof(_CommandBarControlEvents));
_COM_SMARTPTR_TYPEDEF(_dispCommandBarControlEvents, __uuidof(_dispCommandBarControlEvents));
_COM_SMARTPTR_TYPEDEF(_ProjectTemplate, __uuidof(_ProjectTemplate));
_COM_SMARTPTR_TYPEDEF(Events, __uuidof(Events));
_COM_SMARTPTR_TYPEDEF(_Component, __uuidof(_Component));
_COM_SMARTPTR_TYPEDEF(SelectedComponents, __uuidof(SelectedComponents));
_COM_SMARTPTR_TYPEDEF(_dispReferences_Events, __uuidof(_dispReferences_Events));
_COM_SMARTPTR_TYPEDEF(VBE, __uuidof(VBE));
_COM_SMARTPTR_TYPEDEF(Window, __uuidof(Window));
_COM_SMARTPTR_TYPEDEF(_Windows_old, __uuidof(_Windows_old));
_COM_SMARTPTR_TYPEDEF(_LinkedWindows, __uuidof(_LinkedWindows));
_COM_SMARTPTR_TYPEDEF(_VBProject_Old, __uuidof(_VBProject_Old));
_COM_SMARTPTR_TYPEDEF(_VBProject, __uuidof(_VBProject));
_COM_SMARTPTR_TYPEDEF(_VBProjects_Old, __uuidof(_VBProjects_Old));
_COM_SMARTPTR_TYPEDEF(_VBProjects, __uuidof(_VBProjects));
_COM_SMARTPTR_TYPEDEF(_Components, __uuidof(_Components));
_COM_SMARTPTR_TYPEDEF(_VBComponents_Old, __uuidof(_VBComponents_Old));
_COM_SMARTPTR_TYPEDEF(_VBComponents, __uuidof(_VBComponents));
_COM_SMARTPTR_TYPEDEF(_VBComponent_Old, __uuidof(_VBComponent_Old));
_COM_SMARTPTR_TYPEDEF(_VBComponent, __uuidof(_VBComponent));
_COM_SMARTPTR_TYPEDEF(Property, __uuidof(Property));
_COM_SMARTPTR_TYPEDEF(_Properties, __uuidof(_Properties));
_COM_SMARTPTR_TYPEDEF(AddIn, __uuidof(AddIn));
_COM_SMARTPTR_TYPEDEF(_Windows, __uuidof(_Windows));
_COM_SMARTPTR_TYPEDEF(_AddIns, __uuidof(_AddIns));
_COM_SMARTPTR_TYPEDEF(_CodeModule, __uuidof(_CodeModule));
_COM_SMARTPTR_TYPEDEF(_CodePanes, __uuidof(_CodePanes));
_COM_SMARTPTR_TYPEDEF(_CodePane, __uuidof(_CodePane));
_COM_SMARTPTR_TYPEDEF(Reference, __uuidof(Reference));
_COM_SMARTPTR_TYPEDEF(_References, __uuidof(_References));

//
// Type library items
//

struct __declspec(uuid("0002e158-0000-0000-c000-000000000046"))
Application : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetVersion))
    _bstr_t Version;

    //
    // Wrapper methods for error-handling
    //

    _bstr_t GetVersion ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Version (
        /*[out,retval]*/ BSTR * lpbstrReturn ) = 0;
};

enum __declspec(uuid("06a03650-2369-11ce-bfdc-08002b2b8cda"))
vbextFileTypes
{
    vbextFileTypeForm = 0,
    vbextFileTypeModule = 1,
    vbextFileTypeClass = 2,
    vbextFileTypeProject = 3,
    vbextFileTypeExe = 4,
    vbextFileTypeFrx = 5,
    vbextFileTypeRes = 6,
    vbextFileTypeUserControl = 7,
    vbextFileTypePropertyPage = 8,
    vbextFileTypeDocObject = 9,
    vbextFileTypeBinary = 10,
    vbextFileTypeGroupProject = 11,
    vbextFileTypeDesigners = 12
};

enum __declspec(uuid("be39f3db-1b13-11d0-887f-00a0c90f2744"))
vbext_WindowType
{
    vbext_wt_CodeWindow = 0,
    vbext_wt_Designer = 1,
    vbext_wt_Browser = 2,
    vbext_wt_Watch = 3,
    vbext_wt_Locals = 4,
    vbext_wt_Immediate = 5,
    vbext_wt_ProjectWindow = 6,
    vbext_wt_PropertyWindow = 7,
    vbext_wt_Find = 8,
    vbext_wt_FindReplace = 9,
    vbext_wt_Toolbox = 10,
    vbext_wt_LinkedWindowFrame = 11,
    vbext_wt_MainWindow = 12,
    vbext_wt_ToolWindow = 15
};

enum __declspec(uuid("be39f3dc-1b13-11d0-887f-00a0c90f2744"))
vbext_WindowState
{
    vbext_ws_Normal = 0,
    vbext_ws_Minimize = 1,
    vbext_ws_Maximize = 2
};

struct __declspec(uuid("0002e185-0000-0000-c000-000000000046"))
Windows;
    // [ default ] interface _Windows

struct __declspec(uuid("0002e187-0000-0000-c000-000000000046"))
LinkedWindows;
    // [ default ] interface _LinkedWindows

struct __declspec(uuid("0002e113-0000-0000-c000-000000000046"))
_VBProjectsEvents : IUnknown
{};

struct __declspec(uuid("0002e103-0000-0000-c000-000000000046"))
_dispVBProjectsEvents : IDispatch
{
    //
    // Wrapper methods for error-handling
    //

    // Methods:
    HRESULT ItemAdded (
        struct _VBProject * VBProject );
    HRESULT ItemRemoved (
        struct _VBProject * VBProject );
    HRESULT ItemRenamed (
        struct _VBProject * VBProject,
        _bstr_t OldName );
    HRESULT ItemActivated (
        struct _VBProject * VBProject );
};

struct __declspec(uuid("0002e115-0000-0000-c000-000000000046"))
_VBComponentsEvents : IUnknown
{};

struct __declspec(uuid("0002e116-0000-0000-c000-000000000046"))
_dispVBComponentsEvents : IDispatch
{
    //
    // Wrapper methods for error-handling
    //

    // Methods:
    HRESULT ItemAdded (
        struct _VBComponent * VBComponent );
    HRESULT ItemRemoved (
        struct _VBComponent * VBComponent );
    HRESULT ItemRenamed (
        struct _VBComponent * VBComponent,
        _bstr_t OldName );
    HRESULT ItemSelected (
        struct _VBComponent * VBComponent );
    HRESULT ItemActivated (
        struct _VBComponent * VBComponent );
    HRESULT ItemReloaded (
        struct _VBComponent * VBComponent );
};

struct __declspec(uuid("0002e11a-0000-0000-c000-000000000046"))
_ReferencesEvents : IUnknown
{};

struct __declspec(uuid("0002e118-0000-0000-c000-000000000046"))
_dispReferencesEvents : IDispatch
{
    //
    // Wrapper methods for error-handling
    //

    // Methods:
    HRESULT ItemAdded (
        struct Reference * Reference );
    HRESULT ItemRemoved (
        struct Reference * Reference );
};

struct __declspec(uuid("0002e119-0000-0000-c000-000000000046"))
ReferencesEvents;
    // [ default ] interface _ReferencesEvents
    // [ default, source ] dispinterface _dispReferencesEvents

struct __declspec(uuid("0002e130-0000-0000-c000-000000000046"))
_CommandBarControlEvents : IUnknown
{};

struct __declspec(uuid("0002e131-0000-0000-c000-000000000046"))
_dispCommandBarControlEvents : IDispatch
{
    //
    // Wrapper methods for error-handling
    //

    // Methods:
    HRESULT Click (
        IDispatch * CommandBarControl,
        VARIANT_BOOL * handled,
        VARIANT_BOOL * CancelDefault );
};

struct __declspec(uuid("0002e132-0000-0000-c000-000000000046"))
CommandBarEvents;
    // [ default ] interface _CommandBarControlEvents
    // [ default, source ] dispinterface _dispCommandBarControlEvents

struct __declspec(uuid("0002e159-0000-0000-c000-000000000046"))
_ProjectTemplate : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetApplication))
    ApplicationPtr Application;
    __declspec(property(get=GetParent))
    ApplicationPtr Parent;

    //
    // Wrapper methods for error-handling
    //

    ApplicationPtr GetApplication ( );
    ApplicationPtr GetParent ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Application (
        /*[out,retval]*/ struct Application * * lppaReturn ) = 0;
      virtual HRESULT __stdcall get_Parent (
        /*[out,retval]*/ struct Application * * lppaReturn ) = 0;
};

struct __declspec(uuid("32cdf9e0-1602-11ce-bfdc-08002b2b8cda"))
ProjectTemplate;
    // [ default ] interface _ProjectTemplate

enum __declspec(uuid("ffcf3247-debf-11d1-baff-00c04fb6c4a6"))
vbext_ProjectType
{
    vbext_pt_HostProject = 100,
    vbext_pt_StandAlone = 101
};

enum __declspec(uuid("0002e129-0000-0000-c000-000000000046"))
vbext_ProjectProtection
{
    vbext_pp_none = 0,
    vbext_pp_locked = 1
};

enum __declspec(uuid("be39f3d2-1b13-11d0-887f-00a0c90f2744"))
vbext_VBAMode
{
    vbext_vm_Run = 0,
    vbext_vm_Break = 1,
    vbext_vm_Design = 2
};

struct __declspec(uuid("0002e169-0000-0000-c000-000000000046"))
VBProject;
    // [ default ] interface _VBProject

struct __declspec(uuid("0002e167-0000-0000-c000-000000000046"))
Events : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetReferencesEvents))
    _ReferencesEventsPtr ReferencesEvents[];
    __declspec(property(get=GetCommandBarEvents))
    _CommandBarControlEventsPtr CommandBarEvents[];

    //
    // Wrapper methods for error-handling
    //

    _ReferencesEventsPtr GetReferencesEvents (
        struct _VBProject * VBProject );
    _CommandBarControlEventsPtr GetCommandBarEvents (
        IDispatch * CommandBarControl );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ReferencesEvents (
        /*[in]*/ struct _VBProject * VBProject,
        /*[out,retval]*/ struct _ReferencesEvents * * prceNew ) = 0;
      virtual HRESULT __stdcall get_CommandBarEvents (
        /*[in]*/ IDispatch * CommandBarControl,
        /*[out,retval]*/ struct _CommandBarControlEvents * * prceNew ) = 0;
};

struct __declspec(uuid("0002e101-0000-0000-c000-000000000046"))
VBProjects;
    // [ default ] interface _VBProjects

enum __declspec(uuid("be39f3d5-1b13-11d0-887f-00a0c90f2744"))
vbext_ComponentType
{
    vbext_ct_StdModule = 1,
    vbext_ct_ClassModule = 2,
    vbext_ct_MSForm = 3,
    vbext_ct_ActiveXDesigner = 11,
    vbext_ct_Document = 100
};

struct __declspec(uuid("be39f3d6-1b13-11d0-887f-00a0c90f2744"))
Components;
    // [ default ] interface _Components

struct __declspec(uuid("be39f3d7-1b13-11d0-887f-00a0c90f2744"))
VBComponents;
    // [ default ] interface _VBComponents

struct __declspec(uuid("0002e163-0000-0000-c000-000000000046"))
_Component : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetApplication))
    ApplicationPtr Application;
    __declspec(property(get=GetParent))
    _ComponentsPtr Parent;
    __declspec(property(get=GetIsDirty,put=PutIsDirty))
    VARIANT_BOOL IsDirty;
    __declspec(property(get=GetName,put=PutName))
    _bstr_t Name;

    //
    // Wrapper methods for error-handling
    //

    ApplicationPtr GetApplication ( );
    _ComponentsPtr GetParent ( );
    VARIANT_BOOL GetIsDirty ( );
    void PutIsDirty (
        VARIANT_BOOL lpfReturn );
    _bstr_t GetName ( );
    void PutName (
        _bstr_t pbstrReturn );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Application (
        /*[out,retval]*/ struct Application * * lppaReturn ) = 0;
      virtual HRESULT __stdcall get_Parent (
        /*[out,retval]*/ struct _Components * * lppcReturn ) = 0;
      virtual HRESULT __stdcall get_IsDirty (
        /*[out,retval]*/ VARIANT_BOOL * lpfReturn ) = 0;
      virtual HRESULT __stdcall put_IsDirty (
        /*[in]*/ VARIANT_BOOL lpfReturn ) = 0;
      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * pbstrReturn ) = 0;
      virtual HRESULT __stdcall put_Name (
        /*[in]*/ BSTR pbstrReturn ) = 0;
};

struct __declspec(uuid("be39f3d8-1b13-11d0-887f-00a0c90f2744"))
Component;
    // [ default ] interface _Component

struct __declspec(uuid("be39f3d4-1b13-11d0-887f-00a0c90f2744"))
SelectedComponents : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetApplication))
    ApplicationPtr Application;
    __declspec(property(get=GetParent))
    _VBProjectPtr Parent;
    __declspec(property(get=GetCount))
    long Count;

    //
    // Wrapper methods for error-handling
    //

    _ComponentPtr Item (
        int index );
    ApplicationPtr GetApplication ( );
    _VBProjectPtr GetParent ( );
    long GetCount ( );
    IUnknownPtr _NewEnum ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall raw_Item (
        /*[in]*/ int index,
        /*[out,retval]*/ struct _Component * * lppcReturn ) = 0;
      virtual HRESULT __stdcall get_Application (
        /*[out,retval]*/ struct Application * * lppaReturn ) = 0;
      virtual HRESULT __stdcall get_Parent (
        /*[out,retval]*/ struct _VBProject * * lppptReturn ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * lplReturn ) = 0;
      virtual HRESULT __stdcall raw__NewEnum (
        /*[out,retval]*/ IUnknown * * lppiuReturn ) = 0;
};

struct __declspec(uuid("be39f3da-1b13-11d0-887f-00a0c90f2744"))
VBComponent;
    // [ default ] interface _VBComponent

struct __declspec(uuid("0002e18b-0000-0000-c000-000000000046"))
Properties;
    // [ default ] interface _Properties

struct __declspec(uuid("da936b63-ac8b-11d1-b6e5-00a0c90f2744"))
Addins;
    // [ default ] interface _AddIns

enum vbext_ProcKind
{
    vbext_pk_Proc = 0,
    vbext_pk_Let = 1,
    vbext_pk_Set = 2,
    vbext_pk_Get = 3
};

struct __declspec(uuid("0002e170-0000-0000-c000-000000000046"))
CodeModule;
    // [ default ] interface _CodeModule

struct __declspec(uuid("0002e174-0000-0000-c000-000000000046"))
CodePanes;
    // [ default ] interface _CodePanes

enum vbext_CodePaneview
{
    vbext_cv_ProcedureView = 0,
    vbext_cv_FullModuleView = 1
};

struct __declspec(uuid("0002e178-0000-0000-c000-000000000046"))
CodePane;
    // [ default ] interface _CodePane

enum vbext_RefKind
{
    vbext_rk_TypeLib = 0,
    vbext_rk_Project = 1
};

struct __declspec(uuid("cdde3804-2064-11cf-867f-00aa005ff34a"))
_dispReferences_Events : IDispatch
{
    //
    // Wrapper methods for error-handling
    //

    // Methods:
    HRESULT ItemAdded (
        struct Reference * Reference );
    HRESULT ItemRemoved (
        struct Reference * Reference );
};

struct __declspec(uuid("0002e17c-0000-0000-c000-000000000046"))
References;
    // [ default ] interface _References
    // [ default, source ] dispinterface _dispReferences_Events

struct __declspec(uuid("0002e166-0000-0000-c000-000000000046"))
VBE : Application
{
    //
    // Property data
    //

    __declspec(property(get=GetActiveVBProject,put=PutRefActiveVBProject))
    _VBProjectPtr ActiveVBProject;
    __declspec(property(get=GetSelectedVBComponent))
    _VBComponentPtr SelectedVBComponent;
    __declspec(property(get=GetVBProjects))
    _VBProjectsPtr VBProjects;
    __declspec(property(get=GetCommandBars))
    Office::_CommandBarsPtr CommandBars;
    __declspec(property(get=GetCodePanes))
    _CodePanesPtr CodePanes;
    __declspec(property(get=GetWindows))
    _WindowsPtr Windows;
    __declspec(property(get=GetEvents))
    EventsPtr Events;
    __declspec(property(get=GetMainWindow))
    WindowPtr MainWindow;
    __declspec(property(get=GetActiveWindow))
    WindowPtr ActiveWindow;
    __declspec(property(get=GetActiveCodePane,put=PutRefActiveCodePane))
    _CodePanePtr ActiveCodePane;
    __declspec(property(get=GetAddins))
    _AddInsPtr Addins;

    //
    // Wrapper methods for error-handling
    //

    _VBProjectsPtr GetVBProjects ( );
    Office::_CommandBarsPtr GetCommandBars ( );
    _CodePanesPtr GetCodePanes ( );
    _WindowsPtr GetWindows ( );
    EventsPtr GetEvents ( );
    _VBProjectPtr GetActiveVBProject ( );
    void PutRefActiveVBProject (
        struct _VBProject * lppptReturn );
    _VBComponentPtr GetSelectedVBComponent ( );
    WindowPtr GetMainWindow ( );
    WindowPtr GetActiveWindow ( );
    _CodePanePtr GetActiveCodePane ( );
    void PutRefActiveCodePane (
        struct _CodePane * ppCodePane );
    _AddInsPtr GetAddins ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_VBProjects (
        /*[out,retval]*/ struct _VBProjects * * lppptReturn ) = 0;
      virtual HRESULT __stdcall get_CommandBars (
        /*[out,retval]*/ struct Office::_CommandBars * * ppcbs ) = 0;
      virtual HRESULT __stdcall get_CodePanes (
        /*[out,retval]*/ struct _CodePanes * * ppCodePanes ) = 0;
      virtual HRESULT __stdcall get_Windows (
        /*[out,retval]*/ struct _Windows * * ppwnsVBWindows ) = 0;
      virtual HRESULT __stdcall get_Events (
        /*[out,retval]*/ struct Events * * ppevtEvents ) = 0;
      virtual HRESULT __stdcall get_ActiveVBProject (
        /*[out,retval]*/ struct _VBProject * * lppptReturn ) = 0;
      virtual HRESULT __stdcall putref_ActiveVBProject (
        /*[in]*/ struct _VBProject * lppptReturn ) = 0;
      virtual HRESULT __stdcall get_SelectedVBComponent (
        /*[out,retval]*/ struct _VBComponent * * lppscReturn ) = 0;
      virtual HRESULT __stdcall get_MainWindow (
        /*[out,retval]*/ struct Window * * ppwin ) = 0;
      virtual HRESULT __stdcall get_ActiveWindow (
        /*[out,retval]*/ struct Window * * ppwinActive ) = 0;
      virtual HRESULT __stdcall get_ActiveCodePane (
        /*[out,retval]*/ struct _CodePane * * ppCodePane ) = 0;
      virtual HRESULT __stdcall putref_ActiveCodePane (
        /*[in]*/ struct _CodePane * ppCodePane ) = 0;
      virtual HRESULT __stdcall get_Addins (
        /*[out,retval]*/ struct _AddIns * * lpppAddIns ) = 0;
};

struct __declspec(uuid("0002e16b-0000-0000-c000-000000000046"))
Window : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetVBE))
    VBEPtr VBE;
    __declspec(property(get=GetCollection))
    _WindowsPtr Collection;
    __declspec(property(get=GetCaption))
    _bstr_t Caption;
    __declspec(property(get=GetLeft,put=PutLeft))
    long Left;
    __declspec(property(get=GetTop,put=PutTop))
    long Top;
    __declspec(property(get=GetWidth,put=PutWidth))
    long Width;
    __declspec(property(get=GetVisible,put=PutVisible))
    VARIANT_BOOL Visible;
    __declspec(property(get=GetHeight,put=PutHeight))
    long Height;
    __declspec(property(get=GetWindowState,put=PutWindowState))
    enum vbext_WindowState WindowState;
    __declspec(property(get=GetType))
    enum vbext_WindowType Type;
    __declspec(property(get=GetLinkedWindows))
    _LinkedWindowsPtr LinkedWindows;
    __declspec(property(get=GetLinkedWindowFrame))
    WindowPtr LinkedWindowFrame;
    __declspec(property(get=GetHWnd))
    long HWnd;

    //
    // Wrapper methods for error-handling
    //

    VBEPtr GetVBE ( );
    _WindowsPtr GetCollection ( );
    HRESULT Close ( );
    _bstr_t GetCaption ( );
    VARIANT_BOOL GetVisible ( );
    void PutVisible (
        VARIANT_BOOL pfVisible );
    long GetLeft ( );
    void PutLeft (
        long plLeft );
    long GetTop ( );
    void PutTop (
        long plTop );
    long GetWidth ( );
    void PutWidth (
        long plWidth );
    long GetHeight ( );
    void PutHeight (
        long plHeight );
    enum vbext_WindowState GetWindowState ( );
    void PutWindowState (
        enum vbext_WindowState plWindowState );
    HRESULT SetFocus ( );
    enum vbext_WindowType GetType ( );
    HRESULT SetKind (
        enum vbext_WindowType eKind );
    _LinkedWindowsPtr GetLinkedWindows ( );
    WindowPtr GetLinkedWindowFrame ( );
    HRESULT Detach ( );
    HRESULT Attach (
        long lWindowHandle );
    long GetHWnd ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * lppaReturn ) = 0;
      virtual HRESULT __stdcall get_Collection (
        /*[out,retval]*/ struct _Windows * * lppaReturn ) = 0;
      virtual HRESULT __stdcall raw_Close ( ) = 0;
      virtual HRESULT __stdcall get_Caption (
        /*[out,retval]*/ BSTR * pbstrTitle ) = 0;
      virtual HRESULT __stdcall get_Visible (
        /*[out,retval]*/ VARIANT_BOOL * pfVisible ) = 0;
      virtual HRESULT __stdcall put_Visible (
        /*[in]*/ VARIANT_BOOL pfVisible ) = 0;
      virtual HRESULT __stdcall get_Left (
        /*[out,retval]*/ long * plLeft ) = 0;
      virtual HRESULT __stdcall put_Left (
        /*[in]*/ long plLeft ) = 0;
      virtual HRESULT __stdcall get_Top (
        /*[out,retval]*/ long * plTop ) = 0;
      virtual HRESULT __stdcall put_Top (
        /*[in]*/ long plTop ) = 0;
      virtual HRESULT __stdcall get_Width (
        /*[out,retval]*/ long * plWidth ) = 0;
      virtual HRESULT __stdcall put_Width (
        /*[in]*/ long plWidth ) = 0;
      virtual HRESULT __stdcall get_Height (
        /*[out,retval]*/ long * plHeight ) = 0;
      virtual HRESULT __stdcall put_Height (
        /*[in]*/ long plHeight ) = 0;
      virtual HRESULT __stdcall get_WindowState (
        /*[out,retval]*/ enum vbext_WindowState * plWindowState ) = 0;
      virtual HRESULT __stdcall put_WindowState (
        /*[in]*/ enum vbext_WindowState plWindowState ) = 0;
      virtual HRESULT __stdcall raw_SetFocus ( ) = 0;
      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum vbext_WindowType * pKind ) = 0;
      virtual HRESULT __stdcall raw_SetKind (
        /*[in]*/ enum vbext_WindowType eKind ) = 0;
      virtual HRESULT __stdcall get_LinkedWindows (
        /*[out,retval]*/ struct _LinkedWindows * * ppwnsCollection ) = 0;
      virtual HRESULT __stdcall get_LinkedWindowFrame (
        /*[out,retval]*/ struct Window * * ppwinFrame ) = 0;
      virtual HRESULT __stdcall raw_Detach ( ) = 0;
      virtual HRESULT __stdcall raw_Attach (
        /*[in]*/ long lWindowHandle ) = 0;
      virtual HRESULT __stdcall get_HWnd (
        /*[out,retval]*/ long * plWindowHandle ) = 0;
};

struct __declspec(uuid("0002e16a-0000-0000-c000-000000000046"))
_Windows_old : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetVBE))
    VBEPtr VBE;
    __declspec(property(get=GetParent))
    ApplicationPtr Parent;
    __declspec(property(get=GetCount))
    long Count;

    //
    // Wrapper methods for error-handling
    //

    VBEPtr GetVBE ( );
    ApplicationPtr GetParent ( );
    WindowPtr Item (
        const _variant_t & index );
    long GetCount ( );
    IUnknownPtr _NewEnum ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * lppaReturn ) = 0;
      virtual HRESULT __stdcall get_Parent (
        /*[out,retval]*/ struct Application * * lppptReturn ) = 0;
      virtual HRESULT __stdcall raw_Item (
        /*[in]*/ VARIANT index,
        /*[out,retval]*/ struct Window * * lppcReturn ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * lplReturn ) = 0;
      virtual HRESULT __stdcall raw__NewEnum (
        /*[out,retval]*/ IUnknown * * lppiuReturn ) = 0;
};

struct __declspec(uuid("0002e16c-0000-0000-c000-000000000046"))
_LinkedWindows : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetVBE))
    VBEPtr VBE;
    __declspec(property(get=GetParent))
    WindowPtr Parent;
    __declspec(property(get=GetCount))
    long Count;

    //
    // Wrapper methods for error-handling
    //

    VBEPtr GetVBE ( );
    WindowPtr GetParent ( );
    WindowPtr Item (
        const _variant_t & index );
    long GetCount ( );
    IUnknownPtr _NewEnum ( );
    HRESULT Remove (
        struct Window * Window );
    HRESULT Add (
        struct Window * Window );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * lppaReturn ) = 0;
      virtual HRESULT __stdcall get_Parent (
        /*[out,retval]*/ struct Window * * ppptReturn ) = 0;
      virtual HRESULT __stdcall raw_Item (
        /*[in]*/ VARIANT index,
        /*[out,retval]*/ struct Window * * lppcReturn ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * lplReturn ) = 0;
      virtual HRESULT __stdcall raw__NewEnum (
        /*[out,retval]*/ IUnknown * * lppiuReturn ) = 0;
      virtual HRESULT __stdcall raw_Remove (
        /*[in]*/ struct Window * Window ) = 0;
      virtual HRESULT __stdcall raw_Add (
        /*[in]*/ struct Window * Window ) = 0;
};

struct __declspec(uuid("0002e160-0000-0000-c000-000000000046"))
_VBProject_Old : _ProjectTemplate
{
    //
    // Property data
    //

    __declspec(property(get=GetProtection))
    enum vbext_ProjectProtection Protection;
    __declspec(property(get=GetSaved))
    VARIANT_BOOL Saved;
    __declspec(property(get=GetVBComponents))
    _VBComponentsPtr VBComponents;
    __declspec(property(get=GetHelpFile,put=PutHelpFile))
    _bstr_t HelpFile;
    __declspec(property(get=GetHelpContextID,put=PutHelpContextID))
    long HelpContextID;
    __declspec(property(get=GetDescription,put=PutDescription))
    _bstr_t Description;
    __declspec(property(get=GetMode))
    enum vbext_VBAMode Mode;
    __declspec(property(get=GetReferences))
    _ReferencesPtr References;
    __declspec(property(get=GetName,put=PutName))
    _bstr_t Name;
    __declspec(property(get=GetVBE))
    VBEPtr VBE;
    __declspec(property(get=GetCollection))
    _VBProjectsPtr Collection;

    //
    // Wrapper methods for error-handling
    //

    _bstr_t GetHelpFile ( );
    void PutHelpFile (
        _bstr_t lpbstrHelpFile );
    long GetHelpContextID ( );
    void PutHelpContextID (
        long lpdwContextID );
    _bstr_t GetDescription ( );
    void PutDescription (
        _bstr_t lpbstrDescription );
    enum vbext_VBAMode GetMode ( );
    _ReferencesPtr GetReferences ( );
    _bstr_t GetName ( );
    void PutName (
        _bstr_t lpbstrName );
    VBEPtr GetVBE ( );
    _VBProjectsPtr GetCollection ( );
    enum vbext_ProjectProtection GetProtection ( );
    VARIANT_BOOL GetSaved ( );
    _VBComponentsPtr GetVBComponents ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_HelpFile (
        /*[out,retval]*/ BSTR * lpbstrHelpFile ) = 0;
      virtual HRESULT __stdcall put_HelpFile (
        /*[in]*/ BSTR lpbstrHelpFile ) = 0;
      virtual HRESULT __stdcall get_HelpContextID (
        /*[out,retval]*/ long * lpdwContextID ) = 0;
      virtual HRESULT __stdcall put_HelpContextID (
        /*[in]*/ long lpdwContextID ) = 0;
      virtual HRESULT __stdcall get_Description (
        /*[out,retval]*/ BSTR * lpbstrDescription ) = 0;
      virtual HRESULT __stdcall put_Description (
        /*[in]*/ BSTR lpbstrDescription ) = 0;
      virtual HRESULT __stdcall get_Mode (
        /*[out,retval]*/ enum vbext_VBAMode * lpVbaMode ) = 0;
      virtual HRESULT __stdcall get_References (
        /*[out,retval]*/ struct _References * * lppReferences ) = 0;
      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * lpbstrName ) = 0;
      virtual HRESULT __stdcall put_Name (
        /*[in]*/ BSTR lpbstrName ) = 0;
      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * lppaReturn ) = 0;
      virtual HRESULT __stdcall get_Collection (
        /*[out,retval]*/ struct _VBProjects * * lppaReturn ) = 0;
      virtual HRESULT __stdcall get_Protection (
        /*[out,retval]*/ enum vbext_ProjectProtection * lpProtection ) = 0;
      virtual HRESULT __stdcall get_Saved (
        /*[out,retval]*/ VARIANT_BOOL * lpfReturn ) = 0;
      virtual HRESULT __stdcall get_VBComponents (
        /*[out,retval]*/ struct _VBComponents * * lppcReturn ) = 0;
};

struct __declspec(uuid("eee00915-e393-11d1-bb03-00c04fb6c4a6"))
_VBProject : _VBProject_Old
{
    //
    // Property data
    //

    __declspec(property(get=GetType))
    enum vbext_ProjectType Type;
    __declspec(property(get=GetFileName))
    _bstr_t FileName;
    __declspec(property(get=GetBuildFileName,put=PutBuildFileName))
    _bstr_t BuildFileName;

    //
    // Wrapper methods for error-handling
    //

    HRESULT SaveAs (
        _bstr_t FileName );
    HRESULT MakeCompiledFile ( );
    enum vbext_ProjectType GetType ( );
    _bstr_t GetFileName ( );
    _bstr_t GetBuildFileName ( );
    void PutBuildFileName (
        _bstr_t lpbstrBldFName );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall raw_SaveAs (
        /*[in]*/ BSTR FileName ) = 0;
      virtual HRESULT __stdcall raw_MakeCompiledFile ( ) = 0;
      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum vbext_ProjectType * lpkind ) = 0;
      virtual HRESULT __stdcall get_FileName (
        /*[out,retval]*/ BSTR * lpbstrReturn ) = 0;
      virtual HRESULT __stdcall get_BuildFileName (
        /*[out,retval]*/ BSTR * lpbstrBldFName ) = 0;
      virtual HRESULT __stdcall put_BuildFileName (
        /*[in]*/ BSTR lpbstrBldFName ) = 0;
};

struct __declspec(uuid("0002e165-0000-0000-c000-000000000046"))
_VBProjects_Old : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetParent))
    VBEPtr Parent;
    __declspec(property(get=GetCount))
    long Count;
    __declspec(property(get=GetVBE))
    VBEPtr VBE;

    //
    // Wrapper methods for error-handling
    //

    _VBProjectPtr Item (
        const _variant_t & index );
    VBEPtr GetVBE ( );
    VBEPtr GetParent ( );
    long GetCount ( );
    IUnknownPtr _NewEnum ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall raw_Item (
        /*[in]*/ VARIANT index,
        /*[out,retval]*/ struct _VBProject * * lppcReturn ) = 0;
      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * lppaReturn ) = 0;
      virtual HRESULT __stdcall get_Parent (
        /*[out,retval]*/ struct VBE * * lppaReturn ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * lplReturn ) = 0;
      virtual HRESULT __stdcall raw__NewEnum (
        /*[out,retval]*/ IUnknown * * lppiuReturn ) = 0;
};

struct __declspec(uuid("eee00919-e393-11d1-bb03-00c04fb6c4a6"))
_VBProjects : _VBProjects_Old
{
    //
    // Wrapper methods for error-handling
    //

    _VBProjectPtr Add (
        enum vbext_ProjectType Type );
    HRESULT Remove (
        struct _VBProject * lpc );
    _VBProjectPtr Open (
        _bstr_t bstrPath );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall raw_Add (
        /*[in]*/ enum vbext_ProjectType Type,
        /*[out,retval]*/ struct _VBProject * * lppcReturn ) = 0;
      virtual HRESULT __stdcall raw_Remove (
        /*[in]*/ struct _VBProject * lpc ) = 0;
      virtual HRESULT __stdcall raw_Open (
        /*[in]*/ BSTR bstrPath,
        /*[out,retval]*/ struct _VBProject * * lpc ) = 0;
};

struct __declspec(uuid("0002e161-0000-0000-c000-000000000046"))
_Components : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetApplication))
    ApplicationPtr Application;
    __declspec(property(get=GetParent))
    _VBProjectPtr Parent;
    __declspec(property(get=GetCount))
    long Count;
    __declspec(property(get=GetVBE))
    VBEPtr VBE;

    //
    // Wrapper methods for error-handling
    //

    _ComponentPtr Item (
        const _variant_t & index );
    ApplicationPtr GetApplication ( );
    _VBProjectPtr GetParent ( );
    long GetCount ( );
    IUnknownPtr _NewEnum ( );
    HRESULT Remove (
        struct _Component * Component );
    _ComponentPtr Add (
        enum vbext_ComponentType ComponentType );
    _ComponentPtr Import (
        _bstr_t FileName );
    VBEPtr GetVBE ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall raw_Item (
        /*[in]*/ VARIANT index,
        /*[out,retval]*/ struct _Component * * lppcReturn ) = 0;
      virtual HRESULT __stdcall get_Application (
        /*[out,retval]*/ struct Application * * lppaReturn ) = 0;
      virtual HRESULT __stdcall get_Parent (
        /*[out,retval]*/ struct _VBProject * * lppptReturn ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * lplReturn ) = 0;
      virtual HRESULT __stdcall raw__NewEnum (
        /*[out,retval]*/ IUnknown * * lppiuReturn ) = 0;
      virtual HRESULT __stdcall raw_Remove (
        /*[in]*/ struct _Component * Component ) = 0;
      virtual HRESULT __stdcall raw_Add (
        /*[in]*/ enum vbext_ComponentType ComponentType,
        /*[out,retval]*/ struct _Component * * lppComponent ) = 0;
      virtual HRESULT __stdcall raw_Import (
        /*[in]*/ BSTR FileName,
        /*[out,retval]*/ struct _Component * * lppComponent ) = 0;
      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * lppaReturn ) = 0;
};

struct __declspec(uuid("0002e162-0000-0000-c000-000000000046"))
_VBComponents_Old : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetParent))
    _VBProjectPtr Parent;
    __declspec(property(get=GetCount))
    long Count;
    __declspec(property(get=GetVBE))
    VBEPtr VBE;

    //
    // Wrapper methods for error-handling
    //

    _VBComponentPtr Item (
        const _variant_t & index );
    _VBProjectPtr GetParent ( );
    long GetCount ( );
    IUnknownPtr _NewEnum ( );
    HRESULT Remove (
        struct _VBComponent * VBComponent );
    _VBComponentPtr Add (
        enum vbext_ComponentType ComponentType );
    _VBComponentPtr Import (
        _bstr_t FileName );
    VBEPtr GetVBE ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall raw_Item (
        /*[in]*/ VARIANT index,
        /*[out,retval]*/ struct _VBComponent * * lppcReturn ) = 0;
      virtual HRESULT __stdcall get_Parent (
        /*[out,retval]*/ struct _VBProject * * lppptReturn ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * lplReturn ) = 0;
      virtual HRESULT __stdcall raw__NewEnum (
        /*[out,retval]*/ IUnknown * * lppiuReturn ) = 0;
      virtual HRESULT __stdcall raw_Remove (
        /*[in]*/ struct _VBComponent * VBComponent ) = 0;
      virtual HRESULT __stdcall raw_Add (
        /*[in]*/ enum vbext_ComponentType ComponentType,
        /*[out,retval]*/ struct _VBComponent * * lppComponent ) = 0;
      virtual HRESULT __stdcall raw_Import (
        /*[in]*/ BSTR FileName,
        /*[out,retval]*/ struct _VBComponent * * lppComponent ) = 0;
      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * lppaReturn ) = 0;
};

struct __declspec(uuid("eee0091c-e393-11d1-bb03-00c04fb6c4a6"))
_VBComponents : _VBComponents_Old
{
    //
    // Wrapper methods for error-handling
    //

    _VBComponentPtr AddCustom (
        _bstr_t ProgId );
    _VBComponentPtr AddMTDesigner (
        long index );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall raw_AddCustom (
        /*[in]*/ BSTR ProgId,
        /*[out,retval]*/ struct _VBComponent * * lppComponent ) = 0;
      virtual HRESULT __stdcall raw_AddMTDesigner (
        /*[in]*/ long index,
        /*[out,retval]*/ struct _VBComponent * * lppComponent ) = 0;
};

struct __declspec(uuid("0002e164-0000-0000-c000-000000000046"))
_VBComponent_Old : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetSaved))
    VARIANT_BOOL Saved;
    __declspec(property(get=GetName,put=PutName))
    _bstr_t Name;
    __declspec(property(get=GetDesigner))
    IDispatchPtr Designer;
    __declspec(property(get=GetCodeModule))
    _CodeModulePtr CodeModule;
    __declspec(property(get=GetType))
    enum vbext_ComponentType Type;
    __declspec(property(get=GetVBE))
    VBEPtr VBE;
    __declspec(property(get=GetCollection))
    _VBComponentsPtr Collection;
    __declspec(property(get=GetHasOpenDesigner))
    VARIANT_BOOL HasOpenDesigner;
    __declspec(property(get=GetProperties))
    _PropertiesPtr Properties;

    //
    // Wrapper methods for error-handling
    //

    VARIANT_BOOL GetSaved ( );
    _bstr_t GetName ( );
    void PutName (
        _bstr_t pbstrReturn );
    IDispatchPtr GetDesigner ( );
    _CodeModulePtr GetCodeModule ( );
    enum vbext_ComponentType GetType ( );
    HRESULT Export (
        _bstr_t FileName );
    VBEPtr GetVBE ( );
    _VBComponentsPtr GetCollection ( );
    VARIANT_BOOL GetHasOpenDesigner ( );
    _PropertiesPtr GetProperties ( );
    WindowPtr DesignerWindow ( );
    HRESULT Activate ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Saved (
        /*[out,retval]*/ VARIANT_BOOL * lpfReturn ) = 0;
      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * pbstrReturn ) = 0;
      virtual HRESULT __stdcall put_Name (
        /*[in]*/ BSTR pbstrReturn ) = 0;
      virtual HRESULT __stdcall get_Designer (
        /*[out,retval]*/ IDispatch * * ppDispatch ) = 0;
      virtual HRESULT __stdcall get_CodeModule (
        /*[out,retval]*/ struct _CodeModule * * ppVbaModule ) = 0;
      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum vbext_ComponentType * pKind ) = 0;
      virtual HRESULT __stdcall raw_Export (
        /*[in]*/ BSTR FileName ) = 0;
      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * lppaReturn ) = 0;
      virtual HRESULT __stdcall get_Collection (
        /*[out,retval]*/ struct _VBComponents * * lppcReturn ) = 0;
      virtual HRESULT __stdcall get_HasOpenDesigner (
        /*[out,retval]*/ VARIANT_BOOL * lpfReturn ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct _Properties * * lpppReturn ) = 0;
      virtual HRESULT __stdcall raw_DesignerWindow (
        /*[out,retval]*/ struct Window * * lppcReturn ) = 0;
      virtual HRESULT __stdcall raw_Activate ( ) = 0;
};

struct __declspec(uuid("eee00921-e393-11d1-bb03-00c04fb6c4a6"))
_VBComponent : _VBComponent_Old
{
    //
    // Property data
    //

    __declspec(property(get=GetDesignerID))
    _bstr_t DesignerID;

    //
    // Wrapper methods for error-handling
    //

    _bstr_t GetDesignerID ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_DesignerID (
        /*[out,retval]*/ BSTR * pbstrReturn ) = 0;
};

struct __declspec(uuid("0002e18c-0000-0000-c000-000000000046"))
Property : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetValue,put=PutValue))
    _variant_t Value;
    __declspec(property(get=GetApplication))
    ApplicationPtr Application;
    __declspec(property(get=GetParent))
    _PropertiesPtr Parent;
    __declspec(property(get=GetIndexedValue,put=PutIndexedValue))
    _variant_t IndexedValue[][][][];
    __declspec(property(get=GetNumIndices))
    short NumIndices;
    __declspec(property(get=GetName))
    _bstr_t Name;
    __declspec(property(get=GetVBE))
    VBEPtr VBE;
    __declspec(property(get=GetCollection))
    _PropertiesPtr Collection;
    __declspec(property(get=GetObject,put=PutRefObject))
    IUnknownPtr Object;

    //
    // Wrapper methods for error-handling
    //

    _variant_t GetValue ( );
    void PutValue (
        const _variant_t & lppvReturn );
    _variant_t GetIndexedValue (
        const _variant_t & Index1,
        const _variant_t & Index2 = vtMissing,
        const _variant_t & Index3 = vtMissing,
        const _variant_t & Index4 = vtMissing );
    void PutIndexedValue (
        const _variant_t & Index1,
        const _variant_t & Index2,
        const _variant_t & Index3 = vtMissing,
        const _variant_t & Index4 = vtMissing,
        const _variant_t & lppvReturn = vtMissing );
    short GetNumIndices ( );
    ApplicationPtr GetApplication ( );
    _PropertiesPtr GetParent ( );
    _bstr_t GetName ( );
    VBEPtr GetVBE ( );
    _PropertiesPtr GetCollection ( );
    IUnknownPtr GetObject ( );
    void PutRefObject (
        IUnknown * lppunk );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ VARIANT * lppvReturn ) = 0;
      virtual HRESULT __stdcall put_Value (
        /*[in]*/ VARIANT lppvReturn ) = 0;
      virtual HRESULT __stdcall get_IndexedValue (
        /*[in]*/ VARIANT Index1,
        /*[in]*/ VARIANT Index2,
        /*[in]*/ VARIANT Index3,
        /*[in]*/ VARIANT Index4,
        /*[out,retval]*/ VARIANT * lppvReturn ) = 0;
      virtual HRESULT __stdcall put_IndexedValue (
        /*[in]*/ VARIANT Index1,
        /*[in]*/ VARIANT Index2,
        /*[in]*/ VARIANT Index3 = vtMissing,
        /*[in]*/ VARIANT Index4 = vtMissing,
        /*[in]*/ VARIANT lppvReturn = vtMissing ) = 0;
      virtual HRESULT __stdcall get_NumIndices (
        /*[out,retval]*/ short * lpiRetVal ) = 0;
      virtual HRESULT __stdcall get_Application (
        /*[out,retval]*/ struct Application * * lpaReturn ) = 0;
      virtual HRESULT __stdcall get_Parent (
        /*[out,retval]*/ struct _Properties * * lpppReturn ) = 0;
      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * lpbstrReturn ) = 0;
      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * lpaReturn ) = 0;
      virtual HRESULT __stdcall get_Collection (
        /*[out,retval]*/ struct _Properties * * lpppReturn ) = 0;
      virtual HRESULT __stdcall get_Object (
        /*[out,retval]*/ IUnknown * * lppunk ) = 0;
      virtual HRESULT __stdcall putref_Object (
        /*[in]*/ IUnknown * lppunk ) = 0;
};

struct __declspec(uuid("0002e188-0000-0000-c000-000000000046"))
_Properties : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetApplication))
    ApplicationPtr Application;
    __declspec(property(get=GetParent))
    IDispatchPtr Parent;
    __declspec(property(get=GetCount))
    long Count;
    __declspec(property(get=GetVBE))
    VBEPtr VBE;

    //
    // Wrapper methods for error-handling
    //

    PropertyPtr Item (
        const _variant_t & index );
    ApplicationPtr GetApplication ( );
    IDispatchPtr GetParent ( );
    long GetCount ( );
    IUnknownPtr _NewEnum ( );
    VBEPtr GetVBE ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall raw_Item (
        /*[in]*/ VARIANT index,
        /*[out,retval]*/ struct Property * * lplppReturn ) = 0;
      virtual HRESULT __stdcall get_Application (
        /*[out,retval]*/ struct Application * * lppaReturn ) = 0;
      virtual HRESULT __stdcall get_Parent (
        /*[out,retval]*/ IDispatch * * lppidReturn ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * lplReturn ) = 0;
      virtual HRESULT __stdcall raw__NewEnum (
        /*[out,retval]*/ IUnknown * * lppiuReturn ) = 0;
      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * lppaReturn ) = 0;
};

struct __declspec(uuid("da936b64-ac8b-11d1-b6e5-00a0c90f2744"))
AddIn : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetDescription,put=PutDescription))
    _bstr_t Description;
    __declspec(property(get=GetVBE))
    VBEPtr VBE;
    __declspec(property(get=GetCollection))
    _AddInsPtr Collection;
    __declspec(property(get=GetProgId))
    _bstr_t ProgId;
    __declspec(property(get=GetGuid))
    _bstr_t Guid;
    __declspec(property(get=GetConnect,put=PutConnect))
    VARIANT_BOOL Connect;
    __declspec(property(get=GetObject,put=PutObject))
    IDispatchPtr Object;

    //
    // Wrapper methods for error-handling
    //

    _bstr_t GetDescription ( );
    void PutDescription (
        _bstr_t lpbstr );
    VBEPtr GetVBE ( );
    _AddInsPtr GetCollection ( );
    _bstr_t GetProgId ( );
    _bstr_t GetGuid ( );
    VARIANT_BOOL GetConnect ( );
    void PutConnect (
        VARIANT_BOOL lpfConnect );
    IDispatchPtr GetObject ( );
    void PutObject (
        IDispatch * lppdisp );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Description (
        /*[out,retval]*/ BSTR * lpbstr ) = 0;
      virtual HRESULT __stdcall put_Description (
        /*[in]*/ BSTR lpbstr ) = 0;
      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * lppVBE ) = 0;
      virtual HRESULT __stdcall get_Collection (
        /*[out,retval]*/ struct _AddIns * * lppaddins ) = 0;
      virtual HRESULT __stdcall get_ProgId (
        /*[out,retval]*/ BSTR * lpbstr ) = 0;
      virtual HRESULT __stdcall get_Guid (
        /*[out,retval]*/ BSTR * lpbstr ) = 0;
      virtual HRESULT __stdcall get_Connect (
        /*[out,retval]*/ VARIANT_BOOL * lpfConnect ) = 0;
      virtual HRESULT __stdcall put_Connect (
        /*[in]*/ VARIANT_BOOL lpfConnect ) = 0;
      virtual HRESULT __stdcall get_Object (
        /*[out,retval]*/ IDispatch * * lppdisp ) = 0;
      virtual HRESULT __stdcall put_Object (
        /*[in]*/ IDispatch * lppdisp ) = 0;
};

struct __declspec(uuid("f57b7ed0-d8ab-11d1-85df-00c04f98f42c"))
_Windows : _Windows_old
{
    //
    // Wrapper methods for error-handling
    //

    WindowPtr CreateToolWindow (
        struct AddIn * AddInInst,
        _bstr_t ProgId,
        _bstr_t Caption,
        _bstr_t GuidPosition,
        IDispatch * * DocObj );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall raw_CreateToolWindow (
        /*[in]*/ struct AddIn * AddInInst,
        /*[in]*/ BSTR ProgId,
        /*[in]*/ BSTR Caption,
        /*[in]*/ BSTR GuidPosition,
        /*[in,out]*/ IDispatch * * DocObj,
        /*[out,retval]*/ struct Window * * lppcReturn ) = 0;
};

struct __declspec(uuid("da936b62-ac8b-11d1-b6e5-00a0c90f2744"))
_AddIns : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetVBE))
    VBEPtr VBE;
    __declspec(property(get=GetParent))
    IDispatchPtr Parent;
    __declspec(property(get=GetCount))
    long Count;

    //
    // Wrapper methods for error-handling
    //

    AddInPtr Item (
        const _variant_t & index );
    VBEPtr GetVBE ( );
    IDispatchPtr GetParent ( );
    long GetCount ( );
    IUnknownPtr _NewEnum ( );
    HRESULT Update ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall raw_Item (
        /*[in]*/ VARIANT index,
        /*[out,retval]*/ struct AddIn * * lppaddin ) = 0;
      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * lppVBA ) = 0;
      virtual HRESULT __stdcall get_Parent (
        /*[out,retval]*/ IDispatch * * lppVBA ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * lplReturn ) = 0;
      virtual HRESULT __stdcall raw__NewEnum (
        /*[out,retval]*/ IUnknown * * lppiuReturn ) = 0;
      virtual HRESULT __stdcall raw_Update ( ) = 0;
};

struct __declspec(uuid("0002e16e-0000-0000-c000-000000000046"))
_CodeModule : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetParent))
    _VBComponentPtr Parent;
    __declspec(property(get=GetVBE))
    VBEPtr VBE;
    __declspec(property(get=GetName,put=PutName))
    _bstr_t Name;
    __declspec(property(get=GetLines))
    _bstr_t Lines[][];
    __declspec(property(get=GetCountOfLines))
    long CountOfLines;
    __declspec(property(get=GetProcStartLine))
    long ProcStartLine[][];
    __declspec(property(get=GetProcCountLines))
    long ProcCountLines[][];
    __declspec(property(get=GetProcBodyLine))
    long ProcBodyLine[][];
    __declspec(property(get=GetCountOfDeclarationLines))
    long CountOfDeclarationLines;
    __declspec(property(get=GetCodePane))
    _CodePanePtr CodePane;

    //
    // Wrapper methods for error-handling
    //

    _VBComponentPtr GetParent ( );
    VBEPtr GetVBE ( );
    _bstr_t GetName ( );
    void PutName (
        _bstr_t pbstrName );
    HRESULT AddFromString (
        _bstr_t String );
    HRESULT AddFromFile (
        _bstr_t FileName );
    _bstr_t GetLines (
        long StartLine,
        long Count );
    long GetCountOfLines ( );
    HRESULT InsertLines (
        long Line,
        _bstr_t String );
    HRESULT DeleteLines (
        long StartLine,
        long Count );
    HRESULT ReplaceLine (
        long Line,
        _bstr_t String );
    long GetProcStartLine (
        _bstr_t ProcName,
        enum vbext_ProcKind ProcKind );
    long GetProcCountLines (
        _bstr_t ProcName,
        enum vbext_ProcKind ProcKind );
    long GetProcBodyLine (
        _bstr_t ProcName,
        enum vbext_ProcKind ProcKind );
    _bstr_t GetProcOfLine (
        long Line,
        enum vbext_ProcKind * ProcKind );
    long GetCountOfDeclarationLines ( );
    long CreateEventProc (
        _bstr_t EventName,
        _bstr_t ObjectName );
    VARIANT_BOOL Find (
        _bstr_t Target,
        long * StartLine,
        long * StartColumn,
        long * EndLine,
        long * EndColumn,
        VARIANT_BOOL WholeWord,
        VARIANT_BOOL MatchCase,
        VARIANT_BOOL PatternSearch );
    _CodePanePtr GetCodePane ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Parent (
        /*[out,retval]*/ struct _VBComponent * * retval ) = 0;
      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * retval ) = 0;
      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * pbstrName ) = 0;
      virtual HRESULT __stdcall put_Name (
        /*[in]*/ BSTR pbstrName ) = 0;
      virtual HRESULT __stdcall raw_AddFromString (
        /*[in]*/ BSTR String ) = 0;
      virtual HRESULT __stdcall raw_AddFromFile (
        /*[in]*/ BSTR FileName ) = 0;
      virtual HRESULT __stdcall get_Lines (
        /*[in]*/ long StartLine,
        /*[in]*/ long Count,
        /*[out,retval]*/ BSTR * String ) = 0;
      virtual HRESULT __stdcall get_CountOfLines (
        /*[out,retval]*/ long * CountOfLines ) = 0;
      virtual HRESULT __stdcall raw_InsertLines (
        /*[in]*/ long Line,
        /*[in]*/ BSTR String ) = 0;
      virtual HRESULT __stdcall raw_DeleteLines (
        /*[in]*/ long StartLine,
        /*[in]*/ long Count ) = 0;
      virtual HRESULT __stdcall raw_ReplaceLine (
        /*[in]*/ long Line,
        /*[in]*/ BSTR String ) = 0;
      virtual HRESULT __stdcall get_ProcStartLine (
        /*[in]*/ BSTR ProcName,
        /*[in]*/ enum vbext_ProcKind ProcKind,
        /*[out,retval]*/ long * ProcStartLine ) = 0;
      virtual HRESULT __stdcall get_ProcCountLines (
        /*[in]*/ BSTR ProcName,
        /*[in]*/ enum vbext_ProcKind ProcKind,
        /*[out,retval]*/ long * ProcCountLines ) = 0;
      virtual HRESULT __stdcall get_ProcBodyLine (
        /*[in]*/ BSTR ProcName,
        /*[in]*/ enum vbext_ProcKind ProcKind,
        /*[out,retval]*/ long * ProcBodyLine ) = 0;
      virtual HRESULT __stdcall get_ProcOfLine (
        /*[in]*/ long Line,
        /*[out]*/ enum vbext_ProcKind * ProcKind,
        /*[out,retval]*/ BSTR * pbstrName ) = 0;
      virtual HRESULT __stdcall get_CountOfDeclarationLines (
        /*[out,retval]*/ long * pDeclCountOfLines ) = 0;
      virtual HRESULT __stdcall raw_CreateEventProc (
        /*[in]*/ BSTR EventName,
        /*[in]*/ BSTR ObjectName,
        /*[out,retval]*/ long * Line ) = 0;
      virtual HRESULT __stdcall raw_Find (
        /*[in]*/ BSTR Target,
        /*[in,out]*/ long * StartLine,
        /*[in,out]*/ long * StartColumn,
        /*[in,out]*/ long * EndLine,
        /*[in,out]*/ long * EndColumn,
        /*[in]*/ VARIANT_BOOL WholeWord,
        /*[in]*/ VARIANT_BOOL MatchCase,
        /*[in]*/ VARIANT_BOOL PatternSearch,
        /*[out,retval]*/ VARIANT_BOOL * pfFound ) = 0;
      virtual HRESULT __stdcall get_CodePane (
        /*[out,retval]*/ struct _CodePane * * CodePane ) = 0;
};

struct __declspec(uuid("0002e172-0000-0000-c000-000000000046"))
_CodePanes : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetParent))
    VBEPtr Parent;
    __declspec(property(get=GetVBE))
    VBEPtr VBE;
    __declspec(property(get=GetCount))
    long Count;
    __declspec(property(get=GetCurrent,put=PutCurrent))
    _CodePanePtr Current;

    //
    // Wrapper methods for error-handling
    //

    VBEPtr GetParent ( );
    VBEPtr GetVBE ( );
    _CodePanePtr Item (
        const _variant_t & index );
    long GetCount ( );
    IUnknownPtr _NewEnum ( );
    _CodePanePtr GetCurrent ( );
    void PutCurrent (
        struct _CodePane * CodePane );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Parent (
        /*[out,retval]*/ struct VBE * * retval ) = 0;
      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * retval ) = 0;
      virtual HRESULT __stdcall raw_Item (
        /*[in]*/ VARIANT index,
        /*[out,retval]*/ struct _CodePane * * CodePane ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * Count ) = 0;
      virtual HRESULT __stdcall raw__NewEnum (
        /*[out,retval]*/ IUnknown * * ppenum ) = 0;
      virtual HRESULT __stdcall get_Current (
        /*[out,retval]*/ struct _CodePane * * CodePane ) = 0;
      virtual HRESULT __stdcall put_Current (
        /*[in]*/ struct _CodePane * CodePane ) = 0;
};

struct __declspec(uuid("0002e176-0000-0000-c000-000000000046"))
_CodePane : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetCollection))
    _CodePanesPtr Collection;
    __declspec(property(get=GetVBE))
    VBEPtr VBE;
    __declspec(property(get=GetWindow))
    WindowPtr Window;
    __declspec(property(get=GetTopLine,put=PutTopLine))
    long TopLine;
    __declspec(property(get=GetCountOfVisibleLines))
    long CountOfVisibleLines;
    __declspec(property(get=GetCodeModule))
    _CodeModulePtr CodeModule;
    __declspec(property(get=GetCodePaneView))
    enum vbext_CodePaneview CodePaneView;

    //
    // Wrapper methods for error-handling
    //

    _CodePanesPtr GetCollection ( );
    VBEPtr GetVBE ( );
    WindowPtr GetWindow ( );
    HRESULT GetSelection (
        long * StartLine,
        long * StartColumn,
        long * EndLine,
        long * EndColumn );
    HRESULT SetSelection (
        long StartLine,
        long StartColumn,
        long EndLine,
        long EndColumn );
    long GetTopLine ( );
    void PutTopLine (
        long TopLine );
    long GetCountOfVisibleLines ( );
    _CodeModulePtr GetCodeModule ( );
    HRESULT Show ( );
    enum vbext_CodePaneview GetCodePaneView ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Collection (
        /*[out,retval]*/ struct _CodePanes * * retval ) = 0;
      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * retval ) = 0;
      virtual HRESULT __stdcall get_Window (
        /*[out,retval]*/ struct Window * * retval ) = 0;
      virtual HRESULT __stdcall raw_GetSelection (
        /*[out]*/ long * StartLine,
        /*[out]*/ long * StartColumn,
        /*[out]*/ long * EndLine,
        /*[out]*/ long * EndColumn ) = 0;
      virtual HRESULT __stdcall raw_SetSelection (
        /*[in]*/ long StartLine,
        /*[in]*/ long StartColumn,
        /*[in]*/ long EndLine,
        /*[in]*/ long EndColumn ) = 0;
      virtual HRESULT __stdcall get_TopLine (
        /*[out,retval]*/ long * TopLine ) = 0;
      virtual HRESULT __stdcall put_TopLine (
        /*[in]*/ long TopLine ) = 0;
      virtual HRESULT __stdcall get_CountOfVisibleLines (
        /*[out,retval]*/ long * CountOfVisibleLines ) = 0;
      virtual HRESULT __stdcall get_CodeModule (
        /*[out,retval]*/ struct _CodeModule * * CodeModule ) = 0;
      virtual HRESULT __stdcall raw_Show ( ) = 0;
      virtual HRESULT __stdcall get_CodePaneView (
        /*[out,retval]*/ enum vbext_CodePaneview * pCodePaneview ) = 0;
};

struct __declspec(uuid("0002e17e-0000-0000-c000-000000000046"))
Reference : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetCollection))
    _ReferencesPtr Collection;
    __declspec(property(get=GetVBE))
    VBEPtr VBE;
    __declspec(property(get=GetName))
    _bstr_t Name;
    __declspec(property(get=GetGuid))
    _bstr_t Guid;
    __declspec(property(get=GetMajor))
    long Major;
    __declspec(property(get=GetMinor))
    long Minor;
    __declspec(property(get=GetFullPath))
    _bstr_t FullPath;
    __declspec(property(get=GetBuiltIn))
    VARIANT_BOOL BuiltIn;
    __declspec(property(get=GetIsBroken))
    VARIANT_BOOL IsBroken;
    __declspec(property(get=GetType))
    enum vbext_RefKind Type;
    __declspec(property(get=GetDescription))
    _bstr_t Description;

    //
    // Wrapper methods for error-handling
    //

    _ReferencesPtr GetCollection ( );
    VBEPtr GetVBE ( );
    _bstr_t GetName ( );
    _bstr_t GetGuid ( );
    long GetMajor ( );
    long GetMinor ( );
    _bstr_t GetFullPath ( );
    VARIANT_BOOL GetBuiltIn ( );
    VARIANT_BOOL GetIsBroken ( );
    enum vbext_RefKind GetType ( );
    _bstr_t GetDescription ( );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Collection (
        /*[out,retval]*/ struct _References * * retval ) = 0;
      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * lppaReturn ) = 0;
      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * pbstrName ) = 0;
      virtual HRESULT __stdcall get_Guid (
        /*[out,retval]*/ BSTR * pbstrGuid ) = 0;
      virtual HRESULT __stdcall get_Major (
        /*[out,retval]*/ long * pMajor ) = 0;
      virtual HRESULT __stdcall get_Minor (
        /*[out,retval]*/ long * pMinor ) = 0;
      virtual HRESULT __stdcall get_FullPath (
        /*[out,retval]*/ BSTR * pbstrLocation ) = 0;
      virtual HRESULT __stdcall get_BuiltIn (
        /*[out,retval]*/ VARIANT_BOOL * pfIsDefault ) = 0;
      virtual HRESULT __stdcall get_IsBroken (
        /*[out,retval]*/ VARIANT_BOOL * pfIsBroken ) = 0;
      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum vbext_RefKind * pKind ) = 0;
      virtual HRESULT __stdcall get_Description (
        /*[out,retval]*/ BSTR * pbstrName ) = 0;
};

struct __declspec(uuid("0002e17a-0000-0000-c000-000000000046"))
_References : IDispatch
{
    //
    // Property data
    //

    __declspec(property(get=GetParent))
    _VBProjectPtr Parent;
    __declspec(property(get=GetVBE))
    VBEPtr VBE;
    __declspec(property(get=GetCount))
    long Count;

    //
    // Wrapper methods for error-handling
    //

    _VBProjectPtr GetParent ( );
    VBEPtr GetVBE ( );
    ReferencePtr Item (
        const _variant_t & index );
    long GetCount ( );
    IUnknownPtr _NewEnum ( );
    ReferencePtr AddFromGuid (
        _bstr_t Guid,
        long Major,
        long Minor );
    ReferencePtr AddFromFile (
        _bstr_t FileName );
    HRESULT Remove (
        struct Reference * Reference );

    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Parent (
        /*[out,retval]*/ struct _VBProject * * retval ) = 0;
      virtual HRESULT __stdcall get_VBE (
        /*[out,retval]*/ struct VBE * * retval ) = 0;
      virtual HRESULT __stdcall raw_Item (
        /*[in]*/ VARIANT index,
        /*[out,retval]*/ struct Reference * * Reference ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * Count ) = 0;
      virtual HRESULT __stdcall raw__NewEnum (
        /*[out,retval]*/ IUnknown * * ppenum ) = 0;
      virtual HRESULT __stdcall raw_AddFromGuid (
        /*[in]*/ BSTR Guid,
        /*[in]*/ long Major,
        /*[in]*/ long Minor,
        /*[out,retval]*/ struct Reference * * Reference ) = 0;
      virtual HRESULT __stdcall raw_AddFromFile (
        /*[in]*/ BSTR FileName,
        /*[out,retval]*/ struct Reference * * Reference ) = 0;
      virtual HRESULT __stdcall raw_Remove (
        /*[in]*/ struct Reference * Reference ) = 0;
};

//
// Wrapper method implementations
//

//#include "c:\users\huance.xu\desktop\g1.5-excel\source\cps___win32_releaseu\vbe6ext.tli"
#include "..\TLI\vbe6ext.tli"

} // namespace VBIDE

#pragma pack(pop)
