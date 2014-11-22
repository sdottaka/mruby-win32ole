#include "win32ole.h"

/*
 * Document-class: WIN32OLE_EVENT
 *
 *   <code>WIN32OLE_EVENT</code> objects controls OLE event.
 */

typedef struct {
    struct IEventSinkVtbl * lpVtbl;
} IEventSink, *PEVENTSINK;

typedef struct IEventSinkVtbl IEventSinkVtbl;

struct IEventSinkVtbl {
    STDMETHOD(QueryInterface)(
        PEVENTSINK,
        REFIID,
        LPVOID *);
    STDMETHOD_(ULONG, AddRef)(PEVENTSINK);
    STDMETHOD_(ULONG, Release)(PEVENTSINK);

    STDMETHOD(GetTypeInfoCount)(
        PEVENTSINK,
        UINT *);
    STDMETHOD(GetTypeInfo)(
        PEVENTSINK,
        UINT,
        LCID,
        ITypeInfo **);
    STDMETHOD(GetIDsOfNames)(
        PEVENTSINK,
        REFIID,
        OLECHAR **,
        UINT,
        LCID,
        DISPID *);
    STDMETHOD(Invoke)(
        PEVENTSINK,
        DISPID,
        REFIID,
        LCID,
        WORD,
        DISPPARAMS *,
        VARIANT *,
        EXCEPINFO *,
        UINT *);
};

typedef struct tagIEVENTSINKOBJ {
    IEventSinkVtbl *lpVtbl;
    DWORD m_cRef;
    IID m_iid;
    long  m_event_id;
    ITypeInfo *pTypeInfo;
    mrb_state *mrb;
}IEVENTSINKOBJ, *PIEVENTSINKOBJ;

struct oleeventdata {
    DWORD dwCookie;
    IConnectionPoint *pConnectionPoint;
    IDispatch *pDispatch;
    long event_id;
};

#define ary_ole_event (mrb_obj_iv_get(mrb, (struct RObject *)mrb->object_class, mrb_intern_lit(mrb, "_win32ole_ary_ole_event")))
#define id_events (mrb_intern_lit(mrb, "events"))

static BOOL g_IsEventSinkVtblInitialized = FALSE;

static IEventSinkVtbl vtEventSink;

void EVENTSINK_Destructor(PIEVENTSINKOBJ);
static void ole_val2ptr_variant(mrb_state *mrb, mrb_value val, VARIANT *var);
static void hash2ptr_dispparams(mrb_state *mrb, mrb_value hash, ITypeInfo *pTypeInfo, DISPID dispid, DISPPARAMS *pdispparams);
static mrb_value hash2result(mrb_state *mrb, mrb_value hash);
static void ary2ptr_dispparams(mrb_state *mrb, mrb_value ary, DISPPARAMS *pdispparams);
/*
static mrb_value exec_callback(mrb_state *mrb, mrb_value arg);
static mrb_value rescue_callback(mrb_state *mrb, mrb_value arg);
*/
static HRESULT find_iid(mrb_state *mrb, mrb_value ole, const char *pitf, IID *piid, ITypeInfo **ppTypeInfo);
static HRESULT find_coclass(ITypeInfo *pTypeInfo, TYPEATTR *pTypeAttr, ITypeInfo **pTypeInfo2, TYPEATTR **pTypeAttr2);
static HRESULT find_default_source_from_typeinfo(ITypeInfo *pTypeInfo, TYPEATTR *pTypeAttr, ITypeInfo **ppTypeInfo);
static HRESULT find_default_source(mrb_state *mrb, mrb_value ole, IID *piid, ITypeInfo **ppTypeInfo);
static long ole_search_event_at(mrb_state *mrb, mrb_value ary, mrb_value ev);
static mrb_value ole_search_event(mrb_state *mrb, mrb_value ary, mrb_value ev, BOOL  *is_default);
static mrb_sym ole_search_handler_method(mrb_state *mrb, mrb_value handler, mrb_value ev, BOOL *is_default_handler);
static void ole_delete_event(mrb_state *mrb, mrb_value ary, mrb_value ev);
static void ole_event_free(mrb_state *mrb, void *ptr);
static mrb_value ev_advise(mrb_state *mrb, mrb_int argc, mrb_value *argv, mrb_value self);
static mrb_value fev_initialize(mrb_state *mrb, mrb_value self);
static void ole_msg_loop(void);
static mrb_value fev_s_msg_loop(mrb_state *mrb, mrb_value klass);
static void add_event_call_back(mrb_state *mrb, mrb_value obj, mrb_value event, mrb_value data);
static mrb_value ev_on_event(mrb_state *mrb, mrb_int argc, mrb_value *argv, mrb_value self, mrb_value is_ary_arg, mrb_value block);
static mrb_value fev_on_event(mrb_state *mrb, mrb_value self);
static mrb_value fev_on_event_with_outargs(mrb_state *mrb, mrb_value self);
static mrb_value fev_off_event(mrb_state *mrb, mrb_value self);
static mrb_value fev_unadvise(mrb_state *mrb, mrb_value self);
static mrb_value fev_set_handler(mrb_state *mrb, mrb_value self);
static mrb_value fev_get_handler(mrb_state *mrb, mrb_value self);
static mrb_value evs_push(mrb_state *mrb, mrb_value ev);
static mrb_value evs_delete(mrb_state *mrb, long i);
static mrb_value evs_entry(mrb_state *mrb, long i);
static long  evs_length(mrb_state *mrb);

static const mrb_data_type oleevent_datatype = {
    "win32ole_event",
    ole_event_free
};

STDMETHODIMP EVENTSINK_Invoke(
    PEVENTSINK pEventSink,
    DISPID dispid,
    REFIID riid,
    LCID lcid,
    WORD wFlags,
    DISPPARAMS *pdispparams,
    VARIANT *pvarResult,
    EXCEPINFO *pexcepinfo,
    UINT *puArgErr
    ) {

    HRESULT hr;
    BSTR bstr;
    unsigned int count;
    unsigned int i;
    ITypeInfo *pTypeInfo;
    VARIANT *pvar;
    mrb_value ary, obj, event, args, outargv, ev, result;
    mrb_value handler = mrb_nil_value();
/*  mrb_value arg[3]; */
    mrb_sym   mid;
    mrb_value is_outarg = mrb_false_value();
    BOOL is_default_handler = FALSE;
/*  int state; */
    mrb_state *mrb;
    PIEVENTSINKOBJ pEV = (PIEVENTSINKOBJ)pEventSink;
    mrb = pEV->mrb;

    pTypeInfo = pEV->pTypeInfo;
    obj = evs_entry(mrb, pEV->m_event_id);
    if (!mrb_obj_is_kind_of(mrb, obj, C_WIN32OLE_EVENT)) {
        return NOERROR;
    }

    ary = mrb_iv_get(mrb, obj, id_events);
    if (mrb_nil_p(ary) || !mrb_array_p(ary)) {
        return NOERROR;
    }
    hr = pTypeInfo->lpVtbl->GetNames(pTypeInfo, dispid,
                                     &bstr, 1, &count);
    if (FAILED(hr)) {
        return NOERROR;
    }
    ev = WC2VSTR(mrb, bstr);
    event = ole_search_event(mrb, ary, ev, &is_default_handler);
    if (mrb_array_p(event)) {
        handler = mrb_ary_entry(event, 0);
        mid = mrb_intern_lit(mrb, "call");
        is_outarg = mrb_ary_entry(event, 3);
    } else {
        handler = mrb_iv_get(mrb, obj, mrb_intern_lit(mrb, "handler"));
        if (mrb_nil_p(handler)) {
            return NOERROR;
        }
        mid = ole_search_handler_method(mrb, handler, ev, &is_default_handler);
    }
    if (mrb_nil_p(handler) || mid == (mrb_sym)-1) {
        return NOERROR;
    }

    args = mrb_ary_new(mrb);
    if (is_default_handler) {
        mrb_ary_push(mrb, args, ev);
    }

    /* make argument of event handler */
    for (i = 0; i < pdispparams->cArgs; ++i) {
        pvar = &pdispparams->rgvarg[pdispparams->cArgs-i-1];
        mrb_ary_push(mrb, args, ole_variant2val(mrb, pvar));
    }
    outargv = mrb_nil_value();
    if (mrb_bool(is_outarg)) {
	outargv = mrb_ary_new(mrb);
        mrb_ary_push(mrb, args, outargv);
    }

    /*
     * if exception raised in event callback,
     * then you receive cfp consistency error.
     * to avoid this error we use begin rescue end.
     * and the exception raised then error message print
     * and exit ruby process by Win32OLE itself.
     */
    /*
     FIXME:
    arg[0] = handler;
    arg[1] = mid;
    arg[2] = args;
    result = rb_protect(exec_callback, (mrb_value)arg, &state);
    if (state != 0) {
        rescue_callback(mrb, mrb_nil_value());
    }
    */
    result = mrb_funcall_argv(mrb, handler, mid, RARRAY_LEN(args), RARRAY_PTR(args));

    if(mrb_hash_p(result)) {
        hash2ptr_dispparams(mrb, result, pTypeInfo, dispid, pdispparams);
        result = hash2result(mrb, result);
    }else if (mrb_bool(is_outarg) && mrb_array_p(outargv)) {
        ary2ptr_dispparams(mrb, outargv, pdispparams);
    }

    if (pvarResult) {
        VariantInit(pvarResult);
        ole_val2variant(mrb, result, pvarResult);
    }

    return NOERROR;
}

STDMETHODIMP
EVENTSINK_QueryInterface(
    PEVENTSINK pEV,
    REFIID     iid,
    LPVOID*    ppv
    ) {
    if (IsEqualIID(iid, &IID_IUnknown) ||
        IsEqualIID(iid, &IID_IDispatch) ||
        IsEqualIID(iid, &((PIEVENTSINKOBJ)pEV)->m_iid)) {
        *ppv = pEV;
    }
    else {
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    ((LPUNKNOWN)*ppv)->lpVtbl->AddRef((LPUNKNOWN)*ppv);
    return NOERROR;
}

STDMETHODIMP_(ULONG)
EVENTSINK_AddRef(
    PEVENTSINK pEV
    ){
    PIEVENTSINKOBJ pEVObj = (PIEVENTSINKOBJ)pEV;
    return ++pEVObj->m_cRef;
}

STDMETHODIMP_(ULONG) EVENTSINK_Release(
    PEVENTSINK pEV
    ) {
    PIEVENTSINKOBJ pEVObj = (PIEVENTSINKOBJ)pEV;
    --pEVObj->m_cRef;
    if(pEVObj->m_cRef != 0)
        return pEVObj->m_cRef;
    EVENTSINK_Destructor(pEVObj);
    return 0;
}

STDMETHODIMP EVENTSINK_GetTypeInfoCount(
    PEVENTSINK pEV,
    UINT *pct
    ) {
    *pct = 0;
    return NOERROR;
}

STDMETHODIMP EVENTSINK_GetTypeInfo(
    PEVENTSINK pEV,
    UINT info,
    LCID lcid,
    ITypeInfo **pInfo
    ) {
    *pInfo = NULL;
    return DISP_E_BADINDEX;
}

STDMETHODIMP EVENTSINK_GetIDsOfNames(
    PEVENTSINK pEventSink,
    REFIID riid,
    OLECHAR **szNames,
    UINT cNames,
    LCID lcid,
    DISPID *pDispID
    ) {
    ITypeInfo *pTypeInfo;
    PIEVENTSINKOBJ pEV = (PIEVENTSINKOBJ)pEventSink;
    pTypeInfo = pEV->pTypeInfo;
    if (pTypeInfo) {
        return pTypeInfo->lpVtbl->GetIDsOfNames(pTypeInfo, szNames, cNames, pDispID);
    }
    return DISP_E_UNKNOWNNAME;
}

PIEVENTSINKOBJ
EVENTSINK_Constructor(mrb_state *mrb)
{
    PIEVENTSINKOBJ pEv;
    if (!g_IsEventSinkVtblInitialized) {
        vtEventSink.QueryInterface=EVENTSINK_QueryInterface;
        vtEventSink.AddRef = EVENTSINK_AddRef;
        vtEventSink.Release = EVENTSINK_Release;
        vtEventSink.Invoke = EVENTSINK_Invoke;
        vtEventSink.GetIDsOfNames = EVENTSINK_GetIDsOfNames;
        vtEventSink.GetTypeInfoCount = EVENTSINK_GetTypeInfoCount;
        vtEventSink.GetTypeInfo = EVENTSINK_GetTypeInfo;

        g_IsEventSinkVtblInitialized = TRUE;
    }
    pEv = ALLOC_N(IEVENTSINKOBJ, 1);
    if(pEv == NULL) return NULL;
    pEv->lpVtbl = &vtEventSink;
    pEv->m_cRef = 0;
    pEv->m_event_id = 0;
    pEv->pTypeInfo = NULL;
    pEv->mrb = mrb;
    return pEv;
}

void
EVENTSINK_Destructor(
    PIEVENTSINKOBJ pEVObj
    ) {
    if(pEVObj != NULL) {
        OLE_RELEASE(pEVObj->pTypeInfo);
        free(pEVObj);
        pEVObj = NULL;
    }
}

static void
ole_val2ptr_variant(mrb_state *mrb, mrb_value val, VARIANT *var)
{
    switch (mrb_type(val)) {
    case MRB_TT_STRING:
        if (V_VT(var) == (VT_BSTR | VT_BYREF)) {
            *V_BSTRREF(var) = ole_vstr2wc(mrb, val);
        }
        break;
    case MRB_TT_FIXNUM:
        switch(V_VT(var)) {
        case (VT_UI1 | VT_BYREF) :
            *V_UI1REF(var) = NUM2CHR(val);
            break;
        case (VT_I2 | VT_BYREF) :
            *V_I2REF(var) = (short)NUM2INT(val);
            break;
        case (VT_I4 | VT_BYREF) :
            *V_I4REF(var) = NUM2INT(val);
            break;
        case (VT_R4 | VT_BYREF) :
            *V_R4REF(var) = (float)NUM2INT(val);
            break;
        case (VT_R8 | VT_BYREF) :
            *V_R8REF(var) = NUM2INT(val);
            break;
        default:
            break;
        }
        break;
    case MRB_TT_FLOAT:
        switch(V_VT(var)) {
        case (VT_I2 | VT_BYREF) :
            *V_I2REF(var) = (short)NUM2INT(val);
            break;
        case (VT_I4 | VT_BYREF) :
            *V_I4REF(var) = NUM2INT(val);
            break;
        case (VT_R4 | VT_BYREF) :
            *V_R4REF(var) = (float)NUM2DBL(val);
            break;
        case (VT_R8 | VT_BYREF) :
            *V_R8REF(var) = NUM2DBL(val);
            break;
        default:
            break;
        }
        break;
/*  FIXME:
    case MRB_TT_BIGNUM:
        if (V_VT(var) == (VT_R8 | VT_BYREF)) {
            *V_R8REF(var) = rb_big2dbl(val);
        }
        break;
*/
    case MRB_TT_TRUE:
        if (V_VT(var) == (VT_BOOL | VT_BYREF)) {
            *V_BOOLREF(var) = VARIANT_TRUE;
        }
        break;
    case MRB_TT_FALSE:
        if (V_VT(var) == (VT_BOOL | VT_BYREF)) {
            *V_BOOLREF(var) = VARIANT_FALSE;
        }
        break;
    default:
        break;
    }
}

static void
hash2ptr_dispparams(mrb_state *mrb, mrb_value hash, ITypeInfo *pTypeInfo, DISPID dispid, DISPPARAMS *pdispparams)
{
    BSTR *bstrs;
    HRESULT hr;
    UINT len, i;
    VARIANT *pvar;
    mrb_value val;
    mrb_value key;
    len = 0;
    bstrs = ALLOCA_N(BSTR, pdispparams->cArgs + 1);
    hr = pTypeInfo->lpVtbl->GetNames(pTypeInfo, dispid,
                                     bstrs, pdispparams->cArgs + 1,
                                     &len);
    if (FAILED(hr))
	return;

    for (i = 0; i < len - 1; i++) {
	key = WC2VSTR(mrb, bstrs[i + 1]);
        val = mrb_hash_get(mrb, hash, INT2FIX(i));
	if (mrb_nil_p(val))
	    val = mrb_hash_get(mrb, hash, key);
	if (mrb_nil_p(val))
	    val = mrb_hash_get(mrb, hash, mrb_str_intern(mrb, key));
        pvar = &pdispparams->rgvarg[pdispparams->cArgs-i-1];
        ole_val2ptr_variant(mrb, val, pvar);
    }
}

static mrb_value
hash2result(mrb_state *mrb, mrb_value hash)
{
    mrb_value ret = mrb_nil_value();
    ret = mrb_hash_get(mrb, hash, mrb_str_new_lit(mrb, "return"));
    if (mrb_nil_p(ret))
	ret = mrb_hash_get(mrb, hash, mrb_str_intern(mrb, mrb_str_new_lit(mrb, "return")));
    return ret;
}

static void
ary2ptr_dispparams(mrb_state *mrb, mrb_value ary, DISPPARAMS *pdispparams)
{
    int i;
    mrb_value v;
    VARIANT *pvar;
    for(i = 0; i < RARRAY_LEN(ary) && (unsigned int) i < pdispparams->cArgs; i++) {
        v = mrb_ary_entry(ary, i);
        pvar = &pdispparams->rgvarg[pdispparams->cArgs-i-1];
        ole_val2ptr_variant(mrb, v, pvar);
    }
}

/*

FIXME:

static mrb_value
exec_callback(mrb_state *mrb, mrb_value arg)
{
    mrb_value *parg = (mrb_value *)arg;
    mrb_value handler = parg[0];
    mrb_value mid = parg[1];
    mrb_value args = parg[2];
    return rb_apply(handler, mid, args);
}

static mrb_value
rescue_callback(mrb_state *mrb, mrb_value arg)
{

    mrb_value error;
    mrb_value e = rb_errinfo();
    mrb_value bt = mrb_funcall(mrb, e, "backtrace", 0);
    mrb_value msg = mrb_funcall(mrb, e, "message", 0);
    bt = mrb_ary_entry(bt, 0);
    error = mrb_format(mrb, "%S: %S (%S)\n", bt, msg, mrb_obj_classname(mrb, e));
    rb_write_error(mrb_string_value_ptr(mrb, error));
    rb_backtrace();
    ruby_finalize();
    exit(-1);

    return mrb_nil_value();
}
*/

static HRESULT
find_iid(mrb_state *mrb, mrb_value ole, const char *pitf, IID *piid, ITypeInfo **ppTypeInfo)
{
    HRESULT hr;
    IDispatch *pDispatch;
    ITypeInfo *pTypeInfo;
    ITypeLib *pTypeLib;
    TYPEATTR *pTypeAttr;
    HREFTYPE RefType;
    ITypeInfo *pImplTypeInfo;
    TYPEATTR *pImplTypeAttr;

    struct oledata *pole;
    unsigned int index;
    unsigned int count;
    int type;
    BSTR bstr;
    char *pstr;

    BOOL is_found = FALSE;
    LCID    lcid = cWIN32OLE_lcid;

    OLEData_Get_Struct(mrb, ole, pole);

    pDispatch = pole->pDispatch;

    hr = pDispatch->lpVtbl->GetTypeInfo(pDispatch, 0, lcid, &pTypeInfo);
    if (FAILED(hr))
        return hr;

    hr = pTypeInfo->lpVtbl->GetContainingTypeLib(pTypeInfo,
                                                 &pTypeLib,
                                                 &index);
    OLE_RELEASE(pTypeInfo);
    if (FAILED(hr))
        return hr;

    if (!pitf) {
        hr = pTypeLib->lpVtbl->GetTypeInfoOfGuid(pTypeLib,
                                                 piid,
                                                 ppTypeInfo);
        OLE_RELEASE(pTypeLib);
        return hr;
    }
    count = pTypeLib->lpVtbl->GetTypeInfoCount(pTypeLib);
    for (index = 0; index < count; index++) {
        hr = pTypeLib->lpVtbl->GetTypeInfo(pTypeLib,
                                           index,
                                           &pTypeInfo);
        if (FAILED(hr))
            break;
        hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);

        if(FAILED(hr)) {
            OLE_RELEASE(pTypeInfo);
            break;
        }
        if(pTypeAttr->typekind == TKIND_COCLASS) {
            for (type = 0; type < pTypeAttr->cImplTypes; type++) {
                hr = pTypeInfo->lpVtbl->GetRefTypeOfImplType(pTypeInfo,
                                                             type,
                                                             &RefType);
                if (FAILED(hr))
                    break;
                hr = pTypeInfo->lpVtbl->GetRefTypeInfo(pTypeInfo,
                                                       RefType,
                                                       &pImplTypeInfo);
                if (FAILED(hr))
                    break;

                hr = pImplTypeInfo->lpVtbl->GetDocumentation(pImplTypeInfo,
                                                             -1,
                                                             &bstr,
                                                             NULL, NULL, NULL);
                if (FAILED(hr)) {
                    OLE_RELEASE(pImplTypeInfo);
                    break;
                }
                pstr = ole_wc2mb(mrb, bstr);
                if (strcmp(pitf, pstr) == 0) {
                    hr = pImplTypeInfo->lpVtbl->GetTypeAttr(pImplTypeInfo,
                                                            &pImplTypeAttr);
                    if (SUCCEEDED(hr)) {
                        is_found = TRUE;
                        *piid = pImplTypeAttr->guid;
                        if (ppTypeInfo) {
                            *ppTypeInfo = pImplTypeInfo;
                            (*ppTypeInfo)->lpVtbl->AddRef((*ppTypeInfo));
                        }
                        pImplTypeInfo->lpVtbl->ReleaseTypeAttr(pImplTypeInfo,
                                                               pImplTypeAttr);
                    }
                }
                free(pstr);
                OLE_RELEASE(pImplTypeInfo);
                if (is_found || FAILED(hr))
                    break;
            }
        }

        OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
        OLE_RELEASE(pTypeInfo);
        if (is_found || FAILED(hr))
            break;
    }
    OLE_RELEASE(pTypeLib);
    if(!is_found)
        return E_NOINTERFACE;
    return hr;
}

static HRESULT
find_coclass(
    ITypeInfo *pTypeInfo,
    TYPEATTR *pTypeAttr,
    ITypeInfo **pCOTypeInfo,
    TYPEATTR **pCOTypeAttr)
{
    HRESULT hr = E_NOINTERFACE;
    ITypeLib *pTypeLib;
    int count;
    BOOL found = FALSE;
    ITypeInfo *pTypeInfo2 = NULL;
    TYPEATTR *pTypeAttr2 = NULL;
    int flags;
    int i,j;
    HREFTYPE href;
    ITypeInfo *pRefTypeInfo;
    TYPEATTR *pRefTypeAttr;

    hr = pTypeInfo->lpVtbl->GetContainingTypeLib(pTypeInfo, &pTypeLib, NULL);
    if (FAILED(hr)) {
	return hr;
    }
    count = pTypeLib->lpVtbl->GetTypeInfoCount(pTypeLib);
    for (i = 0; i < count && !found; i++) {
        hr = pTypeLib->lpVtbl->GetTypeInfo(pTypeLib, i, &pTypeInfo2);
        if (FAILED(hr))
            continue;
        hr = OLE_GET_TYPEATTR(pTypeInfo2, &pTypeAttr2);
        if (FAILED(hr)) {
            OLE_RELEASE(pTypeInfo2);
            continue;
        }
        if (pTypeAttr2->typekind != TKIND_COCLASS) {
            OLE_RELEASE_TYPEATTR(pTypeInfo2, pTypeAttr2);
            OLE_RELEASE(pTypeInfo2);
            continue;
        }
        for (j = 0; j < pTypeAttr2->cImplTypes && !found; j++) {
            hr = pTypeInfo2->lpVtbl->GetImplTypeFlags(pTypeInfo2, j, &flags);
            if (FAILED(hr))
                continue;
            if (!(flags & IMPLTYPEFLAG_FDEFAULT))
                continue;
            hr = pTypeInfo2->lpVtbl->GetRefTypeOfImplType(pTypeInfo2, j, &href);
            if (FAILED(hr))
                continue;
            hr = pTypeInfo2->lpVtbl->GetRefTypeInfo(pTypeInfo2, href, &pRefTypeInfo);
            if (FAILED(hr))
                continue;
            hr = OLE_GET_TYPEATTR(pRefTypeInfo, &pRefTypeAttr);
            if (FAILED(hr))  {
                OLE_RELEASE(pRefTypeInfo);
                continue;
            }
            if (IsEqualGUID(&(pTypeAttr->guid), &(pRefTypeAttr->guid))) {
                found = TRUE;
            }
        }
        if (!found) {
            OLE_RELEASE_TYPEATTR(pTypeInfo2, pTypeAttr2);
            OLE_RELEASE(pTypeInfo2);
        }
    }
    OLE_RELEASE(pTypeLib);
    if (found) {
        *pCOTypeInfo = pTypeInfo2;
        *pCOTypeAttr = pTypeAttr2;
        hr = S_OK;
    } else {
        hr = E_NOINTERFACE;
    }
    return hr;
}

static HRESULT
find_default_source_from_typeinfo(
    ITypeInfo *pTypeInfo,
    TYPEATTR *pTypeAttr,
    ITypeInfo **ppTypeInfo)
{
    int i = 0;
    HRESULT hr = E_NOINTERFACE;
    int flags;
    HREFTYPE hRefType;
    /* Enumerate all implemented types of the COCLASS */
    for (i = 0; i < pTypeAttr->cImplTypes; i++) {
        hr = pTypeInfo->lpVtbl->GetImplTypeFlags(pTypeInfo, i, &flags);
        if (FAILED(hr))
            continue;

        /*
           looking for the [default] [source]
           we just hope that it is a dispinterface :-)
        */
        if ((flags & IMPLTYPEFLAG_FDEFAULT) &&
            (flags & IMPLTYPEFLAG_FSOURCE)) {

            hr = pTypeInfo->lpVtbl->GetRefTypeOfImplType(pTypeInfo,
                                                         i, &hRefType);
            if (FAILED(hr))
                continue;
            hr = pTypeInfo->lpVtbl->GetRefTypeInfo(pTypeInfo,
                                                   hRefType, ppTypeInfo);
            if (SUCCEEDED(hr))
                break;
        }
    }
    return hr;
}

static HRESULT
find_default_source(mrb_state *mrb, mrb_value ole, IID *piid, ITypeInfo **ppTypeInfo)
{
    HRESULT hr;
    IProvideClassInfo2 *pProvideClassInfo2;
    IProvideClassInfo *pProvideClassInfo;
    void *p;

    IDispatch *pDispatch;
    ITypeInfo *pTypeInfo = NULL;
    ITypeInfo *pTypeInfo2 = NULL;
    TYPEATTR *pTypeAttr = NULL;
    TYPEATTR *pTypeAttr2 = NULL;

    struct oledata *pole;

    OLEData_Get_Struct(mrb, ole, pole);
    pDispatch = pole->pDispatch;
    hr = pDispatch->lpVtbl->QueryInterface(pDispatch,
                                           &IID_IProvideClassInfo2,
                                           &p);
    if (SUCCEEDED(hr)) {
        pProvideClassInfo2 = p;
        hr = pProvideClassInfo2->lpVtbl->GetGUID(pProvideClassInfo2,
                                                 GUIDKIND_DEFAULT_SOURCE_DISP_IID,
                                                 piid);
        OLE_RELEASE(pProvideClassInfo2);
        if (SUCCEEDED(hr)) {
            hr = find_iid(mrb, ole, NULL, piid, ppTypeInfo);
        }
    }
    if (SUCCEEDED(hr)) {
        return hr;
    }
    hr = pDispatch->lpVtbl->QueryInterface(pDispatch,
            &IID_IProvideClassInfo,
            &p);
    if (SUCCEEDED(hr)) {
        pProvideClassInfo = p;
        hr = pProvideClassInfo->lpVtbl->GetClassInfo(pProvideClassInfo,
                                                     &pTypeInfo);
        OLE_RELEASE(pProvideClassInfo);
    }
    if (FAILED(hr)) {
        hr = pDispatch->lpVtbl->GetTypeInfo(pDispatch, 0, cWIN32OLE_lcid, &pTypeInfo );
    }
    if (FAILED(hr))
        return hr;
    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if (FAILED(hr)) {
        OLE_RELEASE(pTypeInfo);
        return hr;
    }

    *ppTypeInfo = 0;
    hr = find_default_source_from_typeinfo(pTypeInfo, pTypeAttr, ppTypeInfo);
    if (!*ppTypeInfo) {
        hr = find_coclass(pTypeInfo, pTypeAttr, &pTypeInfo2, &pTypeAttr2);
        if (SUCCEEDED(hr)) {
            hr = find_default_source_from_typeinfo(pTypeInfo2, pTypeAttr2, ppTypeInfo);
            OLE_RELEASE_TYPEATTR(pTypeInfo2, pTypeAttr2);
            OLE_RELEASE(pTypeInfo2);
        }
    }
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
    OLE_RELEASE(pTypeInfo);
    /* Now that would be a bad surprise, if we didn't find it, wouldn't it? */
    if (!*ppTypeInfo) {
        if (SUCCEEDED(hr))
            hr = E_UNEXPECTED;
        return hr;
    }

    /* Determine IID of default source interface */
    hr = (*ppTypeInfo)->lpVtbl->GetTypeAttr(*ppTypeInfo, &pTypeAttr);
    if (SUCCEEDED(hr)) {
        *piid = pTypeAttr->guid;
        (*ppTypeInfo)->lpVtbl->ReleaseTypeAttr(*ppTypeInfo, pTypeAttr);
    }
    else
        OLE_RELEASE(*ppTypeInfo);

    return hr;
}

static long
ole_search_event_at(mrb_state *mrb, mrb_value ary, mrb_value ev)
{
    mrb_value event;
    mrb_value event_name;
    long i, len;
    long ret = -1;
    len = RARRAY_LEN(ary);
    for(i = 0; i < len; i++) {
        event = mrb_ary_entry(ary, i);
        event_name = mrb_ary_entry(event, 1);
        if(mrb_nil_p(event_name) && mrb_nil_p(ev)) {
            ret = i;
            break;
        }
        else if (mrb_string_p(ev) &&
                 mrb_string_p(event_name) &&
                 mrb_str_cmp(mrb, ev, event_name) == 0) {
            ret = i;
            break;
        }
    }
    return ret;
}

static mrb_value
ole_search_event(mrb_state *mrb, mrb_value ary, mrb_value ev, BOOL  *is_default)
{
    mrb_value event;
    mrb_value def_event;
    mrb_value event_name;
    int i, len;
    *is_default = FALSE;
    def_event = mrb_nil_value();
    len = RARRAY_LEN(ary);
    for(i = 0; i < len; i++) {
        event = mrb_ary_entry(ary, i);
        event_name = mrb_ary_entry(event, 1);
        if(mrb_nil_p(event_name)) {
            *is_default = TRUE;
            def_event = event;
        }
        else if (mrb_str_cmp(mrb, ev, event_name) == 0) {
            *is_default = FALSE;
            return event;
        }
    }
    return def_event;
}

static mrb_sym
ole_search_handler_method(mrb_state *mrb, mrb_value handler, mrb_value ev, BOOL *is_default_handler)
{
    mrb_sym mid;

    *is_default_handler = FALSE;
    mid = mrb_intern_str(mrb, mrb_format(mrb, "on%S", ev));
    if (mrb_respond_to(mrb, handler, mid)) {
        return mid;
    }
    mid = mrb_intern_lit(mrb, "method_missing");
    if (mrb_respond_to(mrb, handler, mid)) {
        *is_default_handler = TRUE;
        return mid;
    }
    return (mrb_sym)-1;
}

static void
ole_delete_event(mrb_state *mrb, mrb_value ary, mrb_value ev)
{
    long at = -1;
    at = ole_search_event_at(mrb, ary, ev);
    if (at >= 0) {
        mrb_funcall(mrb, ary, "delete_at", 1, mrb_fixnum_value(at));
    }
}


static void
ole_event_free(mrb_state *mrb, void *ptr)
{
    struct oleeventdata *poleev = (struct oleeventdata *)ptr;
    if (!ptr) return;
    if (ole_initialized())
    {
        if (poleev->pConnectionPoint) {
            poleev->pConnectionPoint->lpVtbl->Unadvise(poleev->pConnectionPoint, poleev->dwCookie);
            OLE_RELEASE(poleev->pConnectionPoint);
            poleev->pConnectionPoint = NULL;
        }
        OLE_RELEASE(poleev->pDispatch);
    }
    mrb_free(mrb, poleev);
}

static struct oleeventdata *
oleeventdata_alloc(mrb_state *mrb)
{
    struct oleeventdata *poleev = ALLOC(struct oleeventdata);
    poleev->dwCookie = 0;
    poleev->pConnectionPoint = NULL;
    poleev->event_id = 0;
    poleev->pDispatch = NULL;
    return poleev;
}

static mrb_value
ev_advise(mrb_state *mrb, mrb_int argc, mrb_value *argv, mrb_value self)
{

    mrb_value ole, itf;
    struct oledata *pole;
    const char *pitf;
    HRESULT hr;
    IID iid;
    ITypeInfo *pTypeInfo = 0;
    IDispatch *pDispatch;
    IConnectionPointContainer *pContainer;
    IConnectionPoint *pConnectionPoint;
    IEVENTSINKOBJ *pIEV;
    DWORD dwCookie;
    struct oleeventdata *poleev;
    void *p;

    ole = (argc > 0) ? argv[0] : mrb_nil_value();
    itf = (argc > 1) ? argv[1] : mrb_nil_value();

    if (!mrb_obj_is_kind_of(mrb, ole, C_WIN32OLE)) {
        mrb_raise(mrb, E_TYPE_ERROR, "1st parameter must be WIN32OLE object");
    }

    if(!mrb_nil_p(itf)) {
        pitf = mrb_string_value_ptr(mrb, itf);
        /*
        if (rb_safe_level() > 0 && OBJ_TAINTED(itf)) {
            rb_raise(rb_eSecurityError, "insecure event creation - `%s'",
                     StringValuePtr(mrb, itf));
        }
        */
        hr = find_iid(mrb, ole, pitf, &iid, &pTypeInfo);
    }
    else {
        hr = find_default_source(mrb, ole, &iid, &pTypeInfo);
    }
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_RUNTIME_ERROR, "interface not found");
    }

    OLEData_Get_Struct(mrb, ole, pole);
    pDispatch = pole->pDispatch;
    hr = pDispatch->lpVtbl->QueryInterface(pDispatch,
                                           &IID_IConnectionPointContainer,
                                           &p);
    if (FAILED(hr)) {
        OLE_RELEASE(pTypeInfo);
        ole_raise(mrb, hr, E_RUNTIME_ERROR,
                  "failed to query IConnectionPointContainer");
    }
    pContainer = p;

    hr = pContainer->lpVtbl->FindConnectionPoint(pContainer,
                                                 &iid,
                                                 &pConnectionPoint);
    OLE_RELEASE(pContainer);
    if (FAILED(hr)) {
        OLE_RELEASE(pTypeInfo);
        ole_raise(mrb, hr, E_RUNTIME_ERROR, "failed to query IConnectionPoint");
    }
    pIEV = EVENTSINK_Constructor(mrb);
    pIEV->m_iid = iid;
    hr = pConnectionPoint->lpVtbl->Advise(pConnectionPoint,
                                          (IUnknown*)pIEV,
                                          &dwCookie);
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_RUNTIME_ERROR, "Advise Error");
    }

    Data_Get_Struct(mrb, self, &oleevent_datatype, poleev);
    pIEV->m_event_id = evs_length(mrb);
    pIEV->pTypeInfo = pTypeInfo;
    poleev->dwCookie = dwCookie;
    poleev->pConnectionPoint = pConnectionPoint;
    poleev->event_id = pIEV->m_event_id;
    poleev->pDispatch = pDispatch;
    OLE_ADDREF(pDispatch);

    return self;
}

/*
 *  call-seq:
 *     WIN32OLE_EVENT.new(ole, event) #=> WIN32OLE_EVENT object.
 *
 *  Returns OLE event object.
 *  The first argument specifies WIN32OLE object.
 *  The second argument specifies OLE event name.
 *     ie = WIN32OLE.new('InternetExplorer.Application')
 *     ev = WIN32OLE_EVENT.new(ie, 'DWebBrowserEvents')
 */
static mrb_value
fev_initialize(mrb_state *mrb, mrb_value self)
{
    mrb_int argc;
    mrb_value *argv;
    struct oleeventdata *poleev = (struct oleeventdata *)DATA_PTR(self);
    if (poleev) {
        mrb_free(mrb, poleev);
    }
    mrb_data_init(self, NULL, &oleevent_datatype);

    mrb_get_args(mrb, "*", &argv, &argc);

    poleev = oleeventdata_alloc(mrb);
    mrb_data_init(self, poleev, &oleevent_datatype);

    ev_advise(mrb, argc, argv, self);
    evs_push(mrb, self);
    mrb_iv_set(mrb, self, id_events, mrb_ary_new(mrb));
    mrb_iv_set(mrb, self, mrb_intern_lit(mrb, "handler"), mrb_nil_value());
    return self;
}

static void
ole_msg_loop(void)
{
    MSG msg;
    while(PeekMessage(&msg,NULL,0,0,PM_REMOVE)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
}

/*
 *  call-seq:
 *     WIN32OLE_EVENT.message_loop
 *
 *  Translates and dispatches Windows message.
 */
static mrb_value
fev_s_msg_loop(mrb_state *mrb, mrb_value klass)
{
    ole_msg_loop();
    return mrb_nil_value();
}

static void
add_event_call_back(mrb_state *mrb, mrb_value obj, mrb_value event, mrb_value data)
{
    mrb_value events = mrb_iv_get(mrb, obj, id_events);
    if (mrb_nil_p(events) || !mrb_array_p(events)) {
        events = mrb_ary_new(mrb);
        mrb_iv_set(mrb, obj, id_events, events);
    }
    ole_delete_event(mrb, events, event);
    mrb_ary_push(mrb, events, data);
}

static mrb_value
ev_on_event(mrb_state *mrb, mrb_int argc, mrb_value *argv, mrb_value self, mrb_value is_ary_arg, mrb_value block)
{
    struct oleeventdata *poleev;
    mrb_value event, args, data;
    int i;
    Data_Get_Struct(mrb, self, &oleevent_datatype, poleev);
    if (poleev->pConnectionPoint == NULL) {
        mrb_raise(mrb, E_WIN32OLE_RUNTIME_ERROR, "IConnectionPoint not found. You must call advise at first.");
    }

    event = argc > 0 ? argv[0] : mrb_nil_value();
    args = mrb_ary_new(mrb);
    for (i = 0; i < argc - 1; ++i)
        mrb_ary_push(mrb, args, argv[i + 1]);
    if(!mrb_nil_p(event)) {
        if(!mrb_string_p(event) && !mrb_symbol_p(event)) {
            mrb_raise(mrb, E_TYPE_ERROR, "wrong argument type (expected String or Symbol)");
        }
        if (mrb_symbol_p(event)) {
            event = mrb_sym2str(mrb, mrb_symbol(event));
        }
    }
    data = mrb_ary_new_capa(mrb, 4);
    mrb_ary_set(mrb, data, 0, block);
    mrb_ary_set(mrb, data, 1, event);
    mrb_ary_set(mrb, data, 2, args);
    mrb_ary_set(mrb, data, 3, is_ary_arg);
    add_event_call_back(mrb, self, event, data);
    return mrb_nil_value();
}

/*
 *  call-seq:
 *     WIN32OLE_EVENT#on_event([event]){...}
 *
 *  Defines the callback event.
 *  If argument is omitted, this method defines the callback of all events.
 *  If you want to modify reference argument in callback, return hash in
 *  callback. If you want to return value to OLE server as result of callback
 *  use `return' or :return.
 *
 *    ie = WIN32OLE.new('InternetExplorer.Application')
 *    ev = WIN32OLE_EVENT.new(ie)
 *    ev.on_event("NavigateComplete") {|url| puts url}
 *    ev.on_event() {|ev, *args| puts "#{ev} fired"}
 *
 *    ev.on_event("BeforeNavigate2") {|*args|
 *      ...
 *      # set true to BeforeNavigate reference argument `Cancel'.
 *      # Cancel is 7-th argument of BeforeNavigate,
 *      # so you can use 6 as key of hash instead of 'Cancel'.
 *      # The argument is counted from 0.
 *      # The hash key of 0 means first argument.)
 *      {:Cancel => true}  # or {'Cancel' => true} or {6 => true}
 *    }
 *
 *    ev.on_event(...) {|*args|
 *      {:return => 1, :xxx => yyy}
 *    }
 */
static mrb_value
fev_on_event(mrb_state *mrb, mrb_value self)
{
    mrb_int argc;
    mrb_value *argv;
    mrb_value block;
    mrb_get_args(mrb, "*&", &argv, &argc, &block);
    return ev_on_event(mrb, argc, argv, self, mrb_false_value(), block);
}

/*
 *  call-seq:
 *     WIN32OLE_EVENT#on_event_with_outargs([event]){...}
 *
 *  Defines the callback of event.
 *  If you want modify argument in callback,
 *  you could use this method instead of WIN32OLE_EVENT#on_event.
 *
 *    ie = WIN32OLE.new('InternetExplorer.Application')
 *    ev = WIN32OLE_EVENT.new(ie)
 *    ev.on_event_with_outargs('BeforeNavigate2') {|*args|
 *      args.last[6] = true
 *    }
 */
static mrb_value
fev_on_event_with_outargs(mrb_state *mrb, mrb_value self)
{
    mrb_int argc;
    mrb_value *argv;
    mrb_value block;
    mrb_get_args(mrb, "*&", &argv, &argc, &block);
    return ev_on_event(mrb, argc, argv, self, mrb_true_value(), block);
}

/*
 *  call-seq:
 *     WIN32OLE_EVENT#off_event([event])
 *
 *  removes the callback of event.
 *
 *    ie = WIN32OLE.new('InternetExplorer.Application')
 *    ev = WIN32OLE_EVENT.new(ie)
 *    ev.on_event('BeforeNavigate2') {|*args|
 *      args.last[6] = true
 *    }
 *      ...
 *    ev.off_event('BeforeNavigate2')
 *      ...
 */
static mrb_value
fev_off_event(mrb_state *mrb, mrb_value self)
{
    mrb_value event = mrb_nil_value();
    mrb_value events;

    mrb_get_args(mrb, "|o", &event);

    if(!mrb_nil_p(event)) {
        if(!mrb_string_p(event) && !mrb_symbol_p(event)) {
            mrb_raise(mrb, E_TYPE_ERROR, "wrong argument type (expected String or Symbol)");
        }
        if (mrb_symbol_p(event)) {
            event = mrb_sym2str(mrb, mrb_symbol(event));
        }
    }
    events = mrb_iv_get(mrb, self, id_events);
    if (mrb_nil_p(events)) {
        return mrb_nil_value();
    }
    ole_delete_event(mrb, events, event);
    return mrb_nil_value();
}

/*
 *  call-seq:
 *     WIN32OLE_EVENT#unadvise -> nil
 *
 *  disconnects OLE server. If this method called, then the WIN32OLE_EVENT object
 *  does not receive the OLE server event any more.
 *  This method is trial implementation.
 *
 *      ie = WIN32OLE.new('InternetExplorer.Application')
 *      ev = WIN32OLE_EVENT.new(ie)
 *      ev.on_event() {...}
 *         ...
 *      ev.unadvise
 *
 */
static mrb_value
fev_unadvise(mrb_state *mrb, mrb_value self)
{
    struct oleeventdata *poleev;
    Data_Get_Struct(mrb, self, &oleevent_datatype, poleev);
    if (poleev->pConnectionPoint) {
        ole_msg_loop();
        evs_delete(mrb, poleev->event_id);
        poleev->pConnectionPoint->lpVtbl->Unadvise(poleev->pConnectionPoint, poleev->dwCookie);
        OLE_RELEASE(poleev->pConnectionPoint);
        poleev->pConnectionPoint = NULL;
    }
    OLE_FREE(poleev->pDispatch);
    return mrb_nil_value();
}

static mrb_value
evs_push(mrb_state *mrb, mrb_value ev)
{
    mrb_ary_push(mrb, ary_ole_event, ev);
    return ary_ole_event;
}

static mrb_value
evs_delete(mrb_state *mrb, long i)
{
    mrb_ary_set(mrb, ary_ole_event, i, mrb_nil_value());
    return mrb_nil_value();
}

static mrb_value
evs_entry(mrb_state *mrb, long i)
{
    return mrb_ary_entry(ary_ole_event, i);
}

static long
evs_length(mrb_state *mrb)
{
    return RARRAY_LEN(ary_ole_event);
}

/*
 *  call-seq:
 *     WIN32OLE_EVENT#handler=
 *
 *  sets event handler object. If handler object has onXXX
 *  method according to XXX event, then onXXX method is called
 *  when XXX event occurs.
 *
 *  If handler object has method_missing and there is no
 *  method according to the event, then method_missing
 *  called and 1-st argument is event name.
 *
 *  If handler object has onXXX method and there is block
 *  defined by WIN32OLE_EVENT#on_event('XXX'){},
 *  then block is executed but handler object method is not called
 *  when XXX event occurs.
 *
 *      class Handler
 *        def onStatusTextChange(text)
 *          puts "StatusTextChanged"
 *        end
 *        def onPropertyChange(prop)
 *          puts "PropertyChanged"
 *        end
 *        def method_missing(ev, *arg)
 *          puts "other event #{ev}"
 *        end
 *      end
 *
 *      handler = Handler.new
 *      ie = WIN32OLE.new('InternetExplorer.Application')
 *      ev = WIN32OLE_EVENT.new(ie)
 *      ev.on_event("StatusTextChange") {|*args|
 *        puts "this block executed."
 *        puts "handler.onStatusTextChange method is not called."
 *      }
 *      ev.handler = handler
 *
 */
static mrb_value
fev_set_handler(mrb_state *mrb, mrb_value self)
{
    mrb_value val;
    mrb_get_args(mrb, "o", &val);
    mrb_iv_set(mrb, self, mrb_intern_lit(mrb, "handler"), val);
    return val;
}

/*
 *  call-seq:
 *     WIN32OLE_EVENT#handler
 *
 *  returns handler object.
 *
 */
static mrb_value
fev_get_handler(mrb_state *mrb, mrb_value self)
{
    return mrb_iv_get(mrb, self, mrb_intern_lit(mrb, "handler"));
}

void
Init_win32ole_event(mrb_state *mrb)
{
    struct RClass *cWIN32OLE_EVENT;
    mrb_obj_iv_set(mrb, (struct RObject *)mrb->object_class, mrb_intern_lit(mrb, "_win32ole_ary_ole_event"), mrb_ary_new(mrb));
    cWIN32OLE_EVENT = mrb_define_class(mrb, "WIN32OLE_EVENT", mrb->object_class);
    mrb_define_class_method(mrb, cWIN32OLE_EVENT, "message_loop", fev_s_msg_loop, MRB_ARGS_NONE());
    MRB_SET_INSTANCE_TT(cWIN32OLE_EVENT, MRB_TT_DATA);
    mrb_define_method(mrb, cWIN32OLE_EVENT, "initialize", fev_initialize, MRB_ARGS_ANY());
    mrb_define_method(mrb, cWIN32OLE_EVENT, "on_event", fev_on_event, MRB_ARGS_ANY());
    mrb_define_method(mrb, cWIN32OLE_EVENT, "on_event_with_outargs", fev_on_event_with_outargs, MRB_ARGS_ANY());
    mrb_define_method(mrb, cWIN32OLE_EVENT, "off_event", fev_off_event, MRB_ARGS_ANY());
    mrb_define_method(mrb, cWIN32OLE_EVENT, "unadvise", fev_unadvise, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_EVENT, "handler=", fev_set_handler, MRB_ARGS_REQ(1));
    mrb_define_method(mrb, cWIN32OLE_EVENT, "handler", fev_get_handler, MRB_ARGS_NONE());
}
