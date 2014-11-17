/*
 *  (c) 1995 Microsoft Corporation. All rights reserved.
 *  Developed by ActiveWare Internet Corp., http://www.ActiveWare.com
 *
 *  Other modifications Copyright (c) 1997, 1998 by Gurusamy Sarathy
 *  <gsar@umich.edu> and Jan Dubois <jan.dubois@ibm.net>
 *
 *  You may distribute under the terms of either the GNU General Public
 *  License or the Artistic License, as specified in the README file
 *  of the Perl distribution.
 *
 */

/*
  modified for win32ole (ruby) by Masaki.Suketa <masaki.suketa@nifty.ne.jp>
 */

#include "win32ole.h"

/*
 * unfortunately IID_IMultiLanguage2 is not included in any libXXX.a
 * in Cygwin(mingw32).
 */
#if defined(__CYGWIN__) ||  defined(__MINGW32__)
#undef IID_IMultiLanguage2
const IID IID_IMultiLanguage2 = {0xDCCFC164, 0x2B38, 0x11d2, {0xB7, 0xEC, 0x00, 0xC0, 0x4F, 0x8F, 0x5D, 0x9A}};
#endif

#define WIN32OLE_VERSION "1.8.3"

typedef HRESULT (STDAPICALLTYPE FNCOCREATEINSTANCEEX)
    (REFCLSID, IUnknown*, DWORD, COSERVERINFO*, DWORD, MULTI_QI*);

typedef HWND (WINAPI FNHTMLHELP)(HWND hwndCaller, LPCSTR pszFile,
                                 UINT uCommand, DWORD dwData);
typedef BOOL (FNENUMSYSEMCODEPAGES) (CODEPAGE_ENUMPROC, DWORD);

#if defined(RB_THREAD_SPECIFIC) && (defined(__CYGWIN__) || defined(__MINGW32__))
static RB_THREAD_SPECIFIC BOOL g_ole_initialized;
# define g_ole_initialized_init() ((void)0)
# define g_ole_initialized_set(val) (g_ole_initialized = (val))
#else
static volatile DWORD g_ole_initialized_key = TLS_OUT_OF_INDEXES;
# define g_ole_initialized (BOOL)(intptr_t)TlsGetValue(g_ole_initialized_key)
# define g_ole_initialized_init() (g_ole_initialized_key = TlsAlloc())
# define g_ole_initialized_set(val) TlsSetValue(g_ole_initialized_key, (void*)(val))
#endif

typedef struct mrb_encoding
{
   char *name;
} mrb_encoding;

static mrb_encoding g_encs[] = 
{
	{"system"},
	{"UTF-8"}
};

static int g_encindex = -1;
#define mrb_enc_name(enc) (enc)->name
#define mrb_enc_str_new(mrb, str, len, enc) (mrb_str_new((mrb), (str), (len)))

static mrb_encoding *mrb_default_internal_encoding(mrb_state *mrb)
{
	if (g_encindex == -1)
	{
		mrb_value str = mrb_str_new_lit(mrb, "");
		if (mrb_respond_to(mrb, str, mrb_intern_lit(mrb, "codepoints"))) {
			g_encindex = 1; /* UTF-8 */
		} else {
			g_encindex = 0;
		}
	}
	return &g_encs[g_encindex];
}

static mrb_encoding *mrb_default_external_encoding(mrb_state *mrb)
{
	return mrb_default_internal_encoding(mrb);
}

static mrb_encoding *mrb_enc_get(mrb_state *mrb, mrb_value obj) 
{
	if (g_encindex == -1) {
		if (mrb_respond_to(mrb, obj, mrb_intern_lit(mrb, "codepoints"))) {
			g_encindex = 1; /* UTF-8 */
		} else {
			g_encindex = 0;
		}
	}
	return &g_encs[g_encindex];
}

static int mrb_enc_find_index(mrb_state *mrb, const char *enc_name)
{
	if (strcmp(enc_name, "UTF-8") == 0)
		return 1;
	return 0;
}

static int mrb_define_dummy_encoding(mrb_state *mrb, const char *enc_name)
{
	return 0;
}

static mrb_encoding *mrb_enc_from_index(mrb_state *mrb, int idx)
{
	if (idx >= 0 && idx <= 1)
		return &g_encs[idx];
	return NULL;
}

static BOOL g_uninitialize_hooked = FALSE;
static BOOL g_cp_installed = FALSE;
static BOOL g_lcid_installed = FALSE;
static HINSTANCE ghhctrl = NULL;
static HINSTANCE gole32 = NULL;
static FNCOCREATEINSTANCEEX *gCoCreateInstanceEx = NULL;
#define com_hash mrb_gv_get(mrb, mrb_intern_lit(mrb, "win32ole_com_hash"))
static IDispatchVtbl com_vtbl;
static UINT cWIN32OLE_cp = CP_ACP;
static mrb_encoding *cWIN32OLE_enc;
static UINT g_cp_to_check = CP_ACP;
static char g_lcid_to_check[8 + 1];
static VARTYPE g_nil_to = VT_ERROR;
#define enc2cp_table mrb_gv_get(mrb, mrb_intern_lit(mrb, "win32ole_enc2cp_table"))
static IMessageFilterVtbl message_filter;
static IMessageFilter imessage_filter = { &message_filter };
static IMessageFilter* previous_filter;

#if defined(HAVE_TYPE_IMULTILANGUAGE2)
static IMultiLanguage2 *pIMultiLanguage = NULL;
#elif defined(HAVE_TYPE_IMULTILANGUAGE)
static IMultiLanguage *pIMultiLanguage = NULL;
#else
#define pIMultiLanguage NULL /* dummy */
#endif

struct oleparam {
    DISPPARAMS dp;
    OLECHAR** pNamedArgs;
};

static HRESULT ( STDMETHODCALLTYPE QueryInterface )(IDispatch __RPC_FAR *, REFIID riid, void __RPC_FAR *__RPC_FAR *ppvObject);
static ULONG ( STDMETHODCALLTYPE AddRef )(IDispatch __RPC_FAR * This);
static ULONG ( STDMETHODCALLTYPE Release )(IDispatch __RPC_FAR * This);
static HRESULT ( STDMETHODCALLTYPE GetTypeInfoCount )(IDispatch __RPC_FAR * This, UINT __RPC_FAR *pctinfo);
static HRESULT ( STDMETHODCALLTYPE GetTypeInfo )(IDispatch __RPC_FAR * This, UINT iTInfo, LCID lcid, ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
static HRESULT ( STDMETHODCALLTYPE GetIDsOfNames )(IDispatch __RPC_FAR * This, REFIID riid, LPOLESTR __RPC_FAR *rgszNames, UINT cNames, LCID lcid, DISPID __RPC_FAR *rgDispId);
static HRESULT ( STDMETHODCALLTYPE Invoke )( IDispatch __RPC_FAR * This, DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS __RPC_FAR *pDispParams, VARIANT __RPC_FAR *pVarResult, EXCEPINFO __RPC_FAR *pExcepInfo, UINT __RPC_FAR *puArgErr);
static IDispatch* val2dispatch(mrb_state *mrb, mrb_value val);
static double rbtime2vtdate(mrb_state *mrb, mrb_value tmobj);
static mrb_value vtdate2rbtime(mrb_state *mrb, double date);
static mrb_encoding *ole_cp2encoding(mrb_state *mrb, UINT cp);
static UINT ole_encoding2cp(mrb_encoding *enc);
static mrb_noreturn void failed_load_conv51932(mrb_state *mrb);
#ifndef pIMultiLanguage
static void load_conv_function51932(mrb_state *mrb);
#endif
static UINT ole_init_cp(mrb_state *mrb);
static void ole_freeexceptinfo(EXCEPINFO *pExInfo);
static mrb_value ole_excepinfo2msg(mrb_state *mrb, EXCEPINFO *pExInfo);
static void ole_free(mrb_state *mrb, void *ptr);
static BSTR ole_mb2wc(mrb_state *mrb, char *pm, int len);
static mrb_value ole_ary_m_entry(mrb_state *mrb, mrb_value val, LONG *pid);
static mrb_value is_all_index_under(LONG *pid, long *pub, long dim);
static void * get_ptr_of_variant(VARIANT *pvar);
static void ole_set_safe_array(mrb_state *mrb, long n, SAFEARRAY *psa, LONG *pid, long *pub, mrb_value val, long dim,  VARTYPE vt);
static long dimension(mrb_value val);
static long ary_len_of_dim(mrb_value ary, long dim);
static mrb_value ole_set_member(mrb_state *mrb, mrb_value self, IDispatch *dispatch);
static struct oledata *oledata_alloc(mrb_state *mrb);
static mrb_value fole_s_allocate(mrb_state *mrb, struct RClass *klass);
static mrb_value create_win32ole_object(mrb_state *mrb, struct RClass *klass, IDispatch *pDispatch, int argc, mrb_value *argv);
static mrb_value ary_new_dim(mrb_state *mrb, mrb_value myary, LONG *pid, LONG *plb, LONG dim);
static void ary_store_dim(mrb_state *mrb, mrb_value myary, LONG *pid, LONG *plb, LONG dim, mrb_value val);
static void ole_const_load(mrb_state *mrb, ITypeLib *pTypeLib, mrb_value klass, mrb_value self);
static HRESULT clsid_from_remote(mrb_state *mrb, mrb_value host, mrb_value com, CLSID *pclsid);
static mrb_value ole_create_dcom(mrb_state *mrb, mrb_value self, mrb_value ole, mrb_value host, mrb_value others);
static mrb_value ole_bind_obj(mrb_state *mrb, mrb_value moniker, int argc, mrb_value *argv, mrb_value self);
static mrb_value fole_s_connect(mrb_state *mrb, mrb_value self);
static mrb_value fole_s_const_load(mrb_state *mrb, mrb_value self);
static ULONG reference_count(struct oledata * pole);
static mrb_value fole_s_reference_count(mrb_state *mrb, mrb_value self);
static mrb_value fole_s_free(mrb_state *mrb, mrb_value self);
static HWND ole_show_help(mrb_state *mrb, mrb_value helpfile, mrb_value helpcontext);
static mrb_value fole_s_show_help(mrb_state *mrb, mrb_value self);
static mrb_value fole_s_get_code_page(mrb_state *mrb, mrb_value self);
static BOOL CALLBACK installed_code_page_proc(LPSTR str);
static BOOL code_page_installed(UINT cp);
static mrb_value fole_s_set_code_page(mrb_state *mrb, mrb_value self);
static mrb_value fole_s_get_locale(mrb_state *mrb, mrb_value self);
static BOOL CALLBACK installed_lcid_proc(LPSTR str);
static BOOL lcid_installed(LCID lcid);
static mrb_value fole_s_set_locale(mrb_state *mrb, mrb_value self);
static mrb_value fole_s_create_guid(mrb_state *mrb, mrb_value self);
static mrb_value fole_s_ole_initialize(mrb_state *mrb, mrb_value self);
static mrb_value fole_s_ole_uninitialize(mrb_state *mrb, mrb_value self);
static mrb_value fole_initialize(mrb_state *mrb, mrb_value self);
static int hash2named_arg(mrb_state *mrb, mrb_value key, mrb_value val, struct oleparam* pOp);
static mrb_value set_argv(mrb_state *mrb, VARIANTARG* realargs, unsigned int beg, unsigned int end);
static mrb_value ole_invoke(mrb_state *mrb, int argc, mrb_value *argv, mrb_value self, USHORT wFlags, BOOL is_bracket);
static mrb_value fole_invoke(mrb_state *mrb, mrb_value self);
static mrb_value ole_invoke2(mrb_state *mrb, mrb_value self, mrb_value dispid, mrb_value args, mrb_value types, USHORT dispkind);
static mrb_value fole_invoke2(mrb_state *mrb, mrb_value self);
static mrb_value fole_getproperty2(mrb_state *mrb, mrb_value self);
static mrb_value fole_setproperty2(mrb_state *mrb, mrb_value self);
static mrb_value fole_setproperty_with_bracket(mrb_state *mrb, mrb_value self);
static mrb_value fole_setproperty(mrb_state *mrb, mrb_value self);
static mrb_value fole_getproperty_with_bracket(mrb_state *mrb, mrb_value self);
static mrb_value ole_propertyput(mrb_state *mrb, mrb_value self, mrb_value property, mrb_value value);
static mrb_value fole_free(mrb_state *mrb, mrb_value self);
static mrb_value ole_each_sub(mrb_state *mrb, mrb_value self);
static mrb_value ole_ienum_free(mrb_state *mrb, IEnumVARIANT *pEnum);
static mrb_value fole_each(mrb_state *mrb, mrb_value self);
static mrb_value fole_missing(mrb_state *mrb, mrb_value self);
static HRESULT typeinfo_from_ole(mrb_state *mrb, struct oledata *pole, ITypeInfo **ppti);
static mrb_value ole_methods(mrb_state *mrb, mrb_value self, int mask);
static mrb_value fole_methods(mrb_state *mrb, mrb_value self);
static mrb_value fole_get_methods(mrb_state *mrb, mrb_value self);
static mrb_value fole_put_methods(mrb_state *mrb, mrb_value self);
static mrb_value fole_func_methods(mrb_state *mrb, mrb_value self);
static mrb_value fole_type(mrb_state *mrb, mrb_value self);
static mrb_value fole_typelib(mrb_state *mrb, mrb_value self);
static mrb_value fole_query_interface(mrb_state *mrb, mrb_value self);
static mrb_value fole_respond_to(mrb_state *mrb, mrb_value self);
static mrb_value ole_usertype2val(mrb_state *mrb, ITypeInfo *pTypeInfo, TYPEDESC *pTypeDesc, mrb_value typedetails);
static mrb_value ole_ptrtype2val(mrb_state *mrb, ITypeInfo *pTypeInfo, TYPEDESC *pTypeDesc, mrb_value typedetails);
static mrb_value fole_method_help(mrb_state *mrb, mrb_value self);
static mrb_value fole_activex_initialize(mrb_state *mrb, mrb_value self);

static void init_enc2cp(mrb_state *mrb);
static void free_enc2cp(mrb_state *mrb);

const mrb_data_type ole_datatype = {
    "win32ole_ole",
    ole_free
};

static HRESULT (STDMETHODCALLTYPE mf_QueryInterface)(
    IMessageFilter __RPC_FAR * This,
    /* [in] */ REFIID riid,
    /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject)
{
    if (MEMCMP(riid, &IID_IUnknown, GUID, 1) == 0
        || MEMCMP(riid, &IID_IMessageFilter, GUID, 1) == 0)
    {
        *ppvObject = &message_filter;
        return S_OK;
    }
    return E_NOINTERFACE;
}

static ULONG (STDMETHODCALLTYPE mf_AddRef)(
    IMessageFilter __RPC_FAR * This)
{
    return 1;
}

static ULONG (STDMETHODCALLTYPE mf_Release)(
    IMessageFilter __RPC_FAR * This)
{
    return 1;
}

static DWORD (STDMETHODCALLTYPE mf_HandleInComingCall)(
    IMessageFilter __RPC_FAR * pThis,
    DWORD dwCallType,      //Type of incoming call
    HTASK threadIDCaller,  //Task handle calling this task
    DWORD dwTickCount,     //Elapsed tick count
    LPINTERFACEINFO lpInterfaceInfo //Pointer to INTERFACEINFO structure
    )
{
#ifdef DEBUG_MESSAGEFILTER
    printf("incoming %08X, %08X, %d\n", dwCallType, threadIDCaller, dwTickCount);
    fflush(stdout);
#endif
    switch (dwCallType)
    {
    case CALLTYPE_ASYNC:
    case CALLTYPE_TOPLEVEL_CALLPENDING:
    case CALLTYPE_ASYNC_CALLPENDING:
/* FIXME:
        if (rb_during_gc()) {
            return SERVERCALL_RETRYLATER;
        }
*/
        break;
    default:
        break;
    }
    if (previous_filter) {
        return previous_filter->lpVtbl->HandleInComingCall(previous_filter,
                                                   dwCallType,
                                                   threadIDCaller,
                                                   dwTickCount,
                                                   lpInterfaceInfo);
    }
    return SERVERCALL_ISHANDLED;
}

static DWORD (STDMETHODCALLTYPE mf_RetryRejectedCall)(
    IMessageFilter* pThis,
    HTASK threadIDCallee,  //Server task handle
    DWORD dwTickCount,     //Elapsed tick count
    DWORD dwRejectType     //Returned rejection message
    )
{
    if (previous_filter) {
        return previous_filter->lpVtbl->RetryRejectedCall(previous_filter,
                                                  threadIDCallee,
                                                  dwTickCount,
                                                  dwRejectType);
    }
    return 1000;
}

static DWORD (STDMETHODCALLTYPE mf_MessagePending)(
    IMessageFilter* pThis,
    HTASK threadIDCallee,  //Called applications task handle
    DWORD dwTickCount,     //Elapsed tick count
    DWORD dwPendingType    //Call type
    )
{
/* FIXME:
    if (rb_during_gc()) {
        return PENDINGMSG_WAITNOPROCESS;
    }
*/
    if (previous_filter) {
        return previous_filter->lpVtbl->MessagePending(previous_filter,
                                               threadIDCallee,
                                               dwTickCount,
                                               dwPendingType);
    }
    return PENDINGMSG_WAITNOPROCESS;
}

typedef struct _Win32OLEIDispatch
{
    IDispatch dispatch;
    ULONG refcount;
    mrb_value obj;
    mrb_state *mrb;
} Win32OLEIDispatch;

static HRESULT ( STDMETHODCALLTYPE QueryInterface )(
    IDispatch __RPC_FAR * This,
    /* [in] */ REFIID riid,
    /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject)
{
    if (MEMCMP(riid, &IID_IUnknown, GUID, 1) == 0
        || MEMCMP(riid, &IID_IDispatch, GUID, 1) == 0)
    {
        Win32OLEIDispatch* p = (Win32OLEIDispatch*)This;
        p->refcount++;
        *ppvObject = This;
        return S_OK;
    }
    return E_NOINTERFACE;
}

static ULONG ( STDMETHODCALLTYPE AddRef )(
    IDispatch __RPC_FAR * This)
{
    Win32OLEIDispatch* p = (Win32OLEIDispatch*)This;
    return ++(p->refcount);
}

static ULONG ( STDMETHODCALLTYPE Release )(
    IDispatch __RPC_FAR * This)
{
    Win32OLEIDispatch* p = (Win32OLEIDispatch*)This;
    mrb_state *mrb = p->mrb;
    ULONG u = --(p->refcount);
    if (u == 0) {
        mrb_hash_delete_key(mrb, com_hash, p->obj);
        free(p);
    }
    return u;
}

static HRESULT ( STDMETHODCALLTYPE GetTypeInfoCount )(
    IDispatch __RPC_FAR * This,
    /* [out] */ UINT __RPC_FAR *pctinfo)
{
    return E_NOTIMPL;
}

static HRESULT ( STDMETHODCALLTYPE GetTypeInfo )(
    IDispatch __RPC_FAR * This,
    /* [in] */ UINT iTInfo,
    /* [in] */ LCID lcid,
    /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo)
{
    return E_NOTIMPL;
}


static HRESULT ( STDMETHODCALLTYPE GetIDsOfNames )(
    IDispatch __RPC_FAR * This,
    /* [in] */ REFIID riid,
    /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
    /* [in] */ UINT cNames,
    /* [in] */ LCID lcid,
    /* [size_is][out] */ DISPID __RPC_FAR *rgDispId)
{
    Win32OLEIDispatch* p = (Win32OLEIDispatch*)This;
    mrb_state *mrb = p->mrb;
    char* psz = ole_wc2mb(mrb, *rgszNames); // support only one method
    mrb_sym nameid = mrb_intern_cstr(mrb, psz);
    free(psz);
    if ((mrb_sym)(DISPID)nameid != nameid) return E_NOINTERFACE;
    *rgDispId = (DISPID)nameid;
    return S_OK;
}

static /* [local] */ HRESULT ( STDMETHODCALLTYPE Invoke )(
    IDispatch __RPC_FAR * This,
    /* [in] */ DISPID dispIdMember,
    /* [in] */ REFIID riid,
    /* [in] */ LCID lcid,
    /* [in] */ WORD wFlags,
    /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
    /* [out] */ VARIANT __RPC_FAR *pVarResult,
    /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
    /* [out] */ UINT __RPC_FAR *puArgErr)
{
    mrb_value v;
    int i;
    int args = pDispParams->cArgs;
    Win32OLEIDispatch* p = (Win32OLEIDispatch*)This;
    mrb_value* parg = ALLOCA_N(mrb_value, args);
    mrb_sym mid = (mrb_sym)dispIdMember;
    mrb_state *mrb = p->mrb;
    for (i = 0; i < args; i++) {
        *(parg + i) = ole_variant2val(mrb, &pDispParams->rgvarg[args - i - 1]);
    }
    if (dispIdMember == DISPID_VALUE) {
        if (wFlags == DISPATCH_METHOD) {
            mid = mrb_intern_lit(mrb, "call");
        } else if (wFlags & DISPATCH_PROPERTYGET) {
            mid = mrb_intern_lit(mrb, "value");
        }
    }
    v = mrb_funcall_argv(mrb, p->obj, mid, args, parg);
    ole_val2variant(mrb, v, pVarResult);
    return S_OK;
}

BOOL
ole_initialized(void)
{
    return g_ole_initialized;
}

static IDispatch*
val2dispatch(mrb_state *mrb, mrb_value val)
{
    Win32OLEIDispatch* pdisp;
    mrb_value data = mrb_hash_fetch(mrb, com_hash, val, mrb_undef_value());
    if (!mrb_undef_p(data)) {
        pdisp = (Win32OLEIDispatch *)mrb_cptr(data);
        pdisp->refcount++;
    }
    else {
        pdisp = ALLOC(Win32OLEIDispatch);
        pdisp->dispatch.lpVtbl = &com_vtbl;
        pdisp->refcount = 1;
        pdisp->obj = val;
        mrb_hash_set(mrb, com_hash, val, mrb_cptr_value(mrb, pdisp));
    }
    return &pdisp->dispatch;
}

static double
rbtime2vtdate(mrb_state *mrb, mrb_value tmobj)
{
    SYSTEMTIME st;
    double t;
    double usec;

    st.wYear = FIX2INT(mrb_funcall(mrb, tmobj, "year", 0));
    st.wMonth = FIX2INT(mrb_funcall(mrb, tmobj, "month", 0));
    st.wDay = FIX2INT(mrb_funcall(mrb, tmobj, "mday", 0));
    st.wHour = FIX2INT(mrb_funcall(mrb, tmobj, "hour", 0));
    st.wMinute = FIX2INT(mrb_funcall(mrb, tmobj, "min", 0));
    st.wSecond = FIX2INT(mrb_funcall(mrb, tmobj, "sec", 0));
    st.wMilliseconds = 0;
    SystemTimeToVariantTime(&st, &t);

    /*
     * Unfortunately SystemTimeToVariantTime function always ignores the
     * wMilliseconds of SYSTEMTIME struct.
     * So, we need to calculate milliseconds by ourselves.
     */
    usec =  FIX2INT(mrb_funcall(mrb, tmobj, "usec", 0));
    usec /= 1000.0;
    usec /= (24.0 * 3600.0);
    usec /= 1000;
    return t + usec;
}

static mrb_value
vtdate2rbtime(mrb_state *mrb, double date)
{
    SYSTEMTIME st;
    mrb_value v;
    double msec;
    double sec;
    VariantTimeToSystemTime(date, &st);
    v = mrb_funcall(mrb, mrb_obj_value(mrb_class_get(mrb, "Time")), "new", 6,
		      INT2FIX(st.wYear),
		      INT2FIX(st.wMonth),
		      INT2FIX(st.wDay),
		      INT2FIX(st.wHour),
		      INT2FIX(st.wMinute),
		      INT2FIX(st.wSecond));
    st.wYear = FIX2INT(mrb_funcall(mrb, v, "year", 0));
    st.wMonth = FIX2INT(mrb_funcall(mrb, v, "month", 0));
    st.wDay = FIX2INT(mrb_funcall(mrb, v, "mday", 0));
    st.wHour = FIX2INT(mrb_funcall(mrb, v, "hour", 0));
    st.wMinute = FIX2INT(mrb_funcall(mrb, v, "min", 0));
    st.wSecond = FIX2INT(mrb_funcall(mrb, v, "sec", 0));
    st.wMilliseconds = 0;
    SystemTimeToVariantTime(&st, &sec);
    /*
     * Unfortunately VariantTimeToSystemTime always ignores the
     * wMilliseconds of SYSTEMTIME struct(The wMilliseconds is 0).
     * So, we need to calculate milliseconds by ourselves.
     */
    msec = date - sec;
    msec *= 24 * 60;
    msec -= floor(msec);
    msec *= 60;
    if (msec >= 59) {
        msec -= 60;
    }
    if (msec != 0) {
        return mrb_funcall(mrb, v, "+", 1, mrb_float_value(mrb, msec));
    }
    return v;
}

#define ENC_MACHING_CP(enc,encname,cp) if(strcasecmp(mrb_enc_name((enc)),(encname)) == 0) return cp

static UINT ole_encoding2cp(mrb_encoding *enc)
{
    /*
     * Is there any better solution to convert
     * Ruby encoding to Windows codepage???
     */
    ENC_MACHING_CP(enc, "Big5", 950);
    ENC_MACHING_CP(enc, "CP51932", 51932);
    ENC_MACHING_CP(enc, "CP850", 850);
    ENC_MACHING_CP(enc, "CP852", 852);
    ENC_MACHING_CP(enc, "CP855", 855);
    ENC_MACHING_CP(enc, "CP949", 949);
    ENC_MACHING_CP(enc, "EUC-JP", 20932);
    ENC_MACHING_CP(enc, "EUC-KR", 51949);
    ENC_MACHING_CP(enc, "EUC-TW", 51950);
    ENC_MACHING_CP(enc, "GB18030", 54936);
    ENC_MACHING_CP(enc, "GB2312", 20936);
    ENC_MACHING_CP(enc, "GBK", 936);
    ENC_MACHING_CP(enc, "IBM437", 437);
    ENC_MACHING_CP(enc, "IBM737", 737);
    ENC_MACHING_CP(enc, "IBM775", 775);
    ENC_MACHING_CP(enc, "IBM852", 852);
    ENC_MACHING_CP(enc, "IBM855", 855);
    ENC_MACHING_CP(enc, "IBM857", 857);
    ENC_MACHING_CP(enc, "IBM860", 860);
    ENC_MACHING_CP(enc, "IBM861", 861);
    ENC_MACHING_CP(enc, "IBM862", 862);
    ENC_MACHING_CP(enc, "IBM863", 863);
    ENC_MACHING_CP(enc, "IBM864", 864);
    ENC_MACHING_CP(enc, "IBM865", 865);
    ENC_MACHING_CP(enc, "IBM866", 866);
    ENC_MACHING_CP(enc, "IBM869", 869);
    ENC_MACHING_CP(enc, "ISO-2022-JP", 50220);
    ENC_MACHING_CP(enc, "ISO-8859-1", 28591);
    ENC_MACHING_CP(enc, "ISO-8859-15", 28605);
    ENC_MACHING_CP(enc, "ISO-8859-2", 28592);
    ENC_MACHING_CP(enc, "ISO-8859-3", 28593);
    ENC_MACHING_CP(enc, "ISO-8859-4", 28594);
    ENC_MACHING_CP(enc, "ISO-8859-5", 28595);
    ENC_MACHING_CP(enc, "ISO-8859-6", 28596);
    ENC_MACHING_CP(enc, "ISO-8859-7", 28597);
    ENC_MACHING_CP(enc, "ISO-8859-8", 28598);
    ENC_MACHING_CP(enc, "ISO-8859-9", 28599);
    ENC_MACHING_CP(enc, "KOI8-R", 20866);
    ENC_MACHING_CP(enc, "KOI8-U", 21866);
    ENC_MACHING_CP(enc, "Shift_JIS", 932);
    ENC_MACHING_CP(enc, "UTF-16BE", 1201);
    ENC_MACHING_CP(enc, "UTF-16LE", 1200);
    ENC_MACHING_CP(enc, "UTF-7", 65000);
    ENC_MACHING_CP(enc, "UTF-8", 65001);
    ENC_MACHING_CP(enc, "Windows-1250", 1250);
    ENC_MACHING_CP(enc, "Windows-1251", 1251);
    ENC_MACHING_CP(enc, "Windows-1252", 1252);
    ENC_MACHING_CP(enc, "Windows-1253", 1253);
    ENC_MACHING_CP(enc, "Windows-1254", 1254);
    ENC_MACHING_CP(enc, "Windows-1255", 1255);
    ENC_MACHING_CP(enc, "Windows-1256", 1256);
    ENC_MACHING_CP(enc, "Windows-1257", 1257);
    ENC_MACHING_CP(enc, "Windows-1258", 1258);
    ENC_MACHING_CP(enc, "Windows-31J", 932);
    ENC_MACHING_CP(enc, "Windows-874", 874);
    ENC_MACHING_CP(enc, "eucJP-ms", 20932);
    return CP_ACP;
}

static void
failed_load_conv51932(mrb_state *mrb)
{
    mrb_raise(mrb, E_WIN32OLE_RUNTIME_ERROR, "fail to load convert function for CP51932");
}

#ifndef pIMultiLanguage
static void
load_conv_function51932(mrb_state *mrb)
{
    HRESULT hr = E_NOINTERFACE;
    void *p;
    if (!pIMultiLanguage) {
#if defined(HAVE_TYPE_IMULTILANGUAGE2)
	hr = CoCreateInstance(&CLSID_CMultiLanguage, NULL, CLSCTX_INPROC_SERVER,
		              &IID_IMultiLanguage2, &p);
#elif defined(HAVE_TYPE_IMULTILANGUAGE)
	hr = CoCreateInstance(&CLSID_CMultiLanguage, NULL, CLSCTX_INPROC_SERVER,
		              &IID_IMultiLanguage, &p);
#endif
	if (FAILED(hr)) {
	    failed_load_conv51932(mrb);
	}
	pIMultiLanguage = p;
    }
}
#else
#define load_conv_function51932(mrb) failed_load_conv51932(mrb)
#endif

#define conv_51932(mrb, cp) ((cp) == 51932 && (load_conv_function51932(mrb), 1))

static void
set_ole_codepage(mrb_state *mrb, UINT cp)
{
    if (code_page_installed(cp)) {
        cWIN32OLE_cp = cp;
    } else {
        switch(cp) {
        case CP_ACP:
        case CP_OEMCP:
        case CP_MACCP:
        case CP_THREAD_ACP:
        case CP_SYMBOL:
        case CP_UTF7:
        case CP_UTF8:
            cWIN32OLE_cp = cp;
            break;
        case 51932:
            cWIN32OLE_cp = cp;
            load_conv_function51932(mrb);
            break;
        default:
            mrb_raise(mrb, E_WIN32OLE_RUNTIME_ERROR, "codepage should be WIN32OLE::CP_ACP, WIN32OLE::CP_OEMCP, WIN32OLE::CP_MACCP, WIN32OLE::CP_THREAD_ACP, WIN32OLE::CP_SYMBOL, WIN32OLE::CP_UTF7, WIN32OLE::CP_UTF8, or installed codepage.");
            break;
        }
    }
    cWIN32OLE_enc = ole_cp2encoding(mrb, cWIN32OLE_cp);
}


static UINT
ole_init_cp(mrb_state *mrb)
{
    UINT cp;
    mrb_encoding *encdef;
    encdef = mrb_default_internal_encoding(mrb);
    if (!encdef) {
	encdef = mrb_default_external_encoding(mrb);
    }
    cp = ole_encoding2cp(encdef);
    set_ole_codepage(mrb, cp);
    return cp;
}

struct myCPINFOEX {
  UINT MaxCharSize;
  BYTE DefaultChar[2];
  BYTE LeadByte[12];
  WCHAR UnicodeDefaultChar;
  UINT CodePage;
  char CodePageName[MAX_PATH];
};

static mrb_encoding *
ole_cp2encoding(mrb_state *mrb, UINT cp)
{
    static BOOL (*pGetCPInfoEx)(UINT, DWORD, struct myCPINFOEX *) = NULL;
    struct myCPINFOEX* buf;
    mrb_value enc_name;
    const char *enc_cstr;
    int idx;

    if (!code_page_installed(cp)) {
	switch(cp) {
	  case CP_ACP:
	    cp = GetACP();
	    break;
	  case CP_OEMCP:
	    cp = GetOEMCP();
	    break;
	  case CP_MACCP:
	  case CP_THREAD_ACP:
	    if (!pGetCPInfoEx) {
		pGetCPInfoEx = (BOOL (*)(UINT, DWORD, struct myCPINFOEX *))
		    GetProcAddress(GetModuleHandleA("kernel32"), "GetCPInfoEx");
		if (!pGetCPInfoEx) {
		    pGetCPInfoEx = (void*)-1;
		}
	    }
	    buf = ALLOCA_N(struct myCPINFOEX, 1);
	    ZeroMemory(buf, sizeof(struct myCPINFOEX));
	    if (pGetCPInfoEx == (void*)-1 || !pGetCPInfoEx(cp, 0, buf)) {
		mrb_raise(mrb, E_WIN32OLE_RUNTIME_ERROR, "cannot map codepage to encoding.");
		break;	/* never reach here */
	    }
	    cp = buf->CodePage;
	    break;
	  case CP_SYMBOL:
	  case CP_UTF7:
	  case CP_UTF8:
	    break;
	  case 51932:
	    load_conv_function51932(mrb);
	    break;
	  default:
            mrb_raise(mrb, E_WIN32OLE_RUNTIME_ERROR, "codepage should be WIN32OLE::CP_ACP, WIN32OLE::CP_OEMCP, WIN32OLE::CP_MACCP, WIN32OLE::CP_THREAD_ACP, WIN32OLE::CP_SYMBOL, WIN32OLE::CP_UTF7, WIN32OLE::CP_UTF8, or installed codepage.");
            break;
        }
    }

    enc_name = mrb_format(mrb, "CP%S", mrb_fixnum_value(cp));
    idx = mrb_enc_find_index(mrb, enc_cstr = mrb_string_value_cstr(mrb, &enc_name));
    if (idx < 0)
	idx = mrb_define_dummy_encoding(mrb, enc_cstr);
    return mrb_enc_from_index(mrb, idx);
}

static char *
ole_wc2mb_alloc(mrb_state *mrb, LPWSTR pw, char *(alloc)(mrb_state *mrb, UINT size, void *arg), void *arg)
{
    LPSTR pm;
    UINT size = 0;
    if (conv_51932(mrb, cWIN32OLE_cp)) {
#ifndef pIMultiLanguage
	DWORD dw = 0;
	HRESULT hr = pIMultiLanguage->lpVtbl->ConvertStringFromUnicode(pIMultiLanguage,
		&dw, cWIN32OLE_cp, pw, NULL, NULL, &size);
	if (FAILED(hr)) {
            ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "fail to convert Unicode to CP%d", cWIN32OLE_cp);
	}
	pm = alloc(mrb, size, arg);
	hr = pIMultiLanguage->lpVtbl->ConvertStringFromUnicode(pIMultiLanguage,
		&dw, cWIN32OLE_cp, pw, NULL, pm, &size);
	if (FAILED(hr)) {
            ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "fail to convert Unicode to CP%d", cWIN32OLE_cp);
	}
	pm[size] = '\0';
#endif
        return pm;
    }
    size = WideCharToMultiByte(cWIN32OLE_cp, 0, pw, -1, NULL, 0, NULL, NULL);
    if (size) {
        pm = alloc(mrb, size, arg);
        WideCharToMultiByte(cWIN32OLE_cp, 0, pw, -1, pm, size, NULL, NULL);
        pm[size] = '\0';
    }
    else {
        pm = alloc(mrb, 0, arg);
        *pm = '\0';
    }
    return pm;
}

static char *
ole_alloc_str(mrb_state *mrb, UINT size, void *arg)
{
    return ALLOC_N(char, size + 1);
}

char *
ole_wc2mb(mrb_state *mrb, LPWSTR pw)
{
    return ole_wc2mb_alloc(mrb, pw, ole_alloc_str, NULL);
}

static void
ole_freeexceptinfo(EXCEPINFO *pExInfo)
{
    SysFreeString(pExInfo->bstrDescription);
    SysFreeString(pExInfo->bstrSource);
    SysFreeString(pExInfo->bstrHelpFile);
}

static mrb_value
ole_excepinfo2msg(mrb_state *mrb, EXCEPINFO *pExInfo)
{
    char error_code[40];
    char *pSource = NULL;
    char *pDescription = NULL;
    mrb_value error_msg;
    if(pExInfo->pfnDeferredFillIn != NULL) {
        (*pExInfo->pfnDeferredFillIn)(pExInfo);
    }
    if (pExInfo->bstrSource != NULL) {
        pSource = ole_wc2mb(mrb, pExInfo->bstrSource);
    }
    if (pExInfo->bstrDescription != NULL) {
        pDescription = ole_wc2mb(mrb, pExInfo->bstrDescription);
    }
    if(pExInfo->wCode == 0) {
        sprintf(error_code, "\n    OLE error code:%lX in ", (unsigned long)pExInfo->scode);
    }
    else{
        sprintf(error_code, "\n    OLE error code:%u in ", pExInfo->wCode);
    }
    error_msg = mrb_str_new_cstr(mrb, error_code);
    if(pSource != NULL) {
        mrb_str_cat_cstr(mrb, error_msg, pSource);
    }
    else {
        mrb_str_cat_lit(mrb, error_msg, "<Unknown>");
    }
    mrb_str_cat_lit(mrb, error_msg, "\n      ");
    if(pDescription != NULL) {
        mrb_str_cat_cstr(mrb, error_msg, pDescription);
    }
    else {
        mrb_str_cat_lit(mrb, error_msg, "<No Description>");
    }
    if(pSource) free(pSource);
    if(pDescription) free(pDescription);
    ole_freeexceptinfo(pExInfo);
    return error_msg;
}

void
ole_uninitialize(void)
{
    if (!g_ole_initialized) return;
    OleUninitialize();
    g_ole_initialized_set(FALSE);
}

void
ole_initialize(mrb_state *mrb)
{
    HRESULT hr;

    if(!g_uninitialize_hooked) {
/* FIXME:
	rb_add_event_hook(ole_uninitialize_hook, RUBY_EVENT_THREAD_END, mrb_nil_value());
*/
/* */
	g_uninitialize_hooked = TRUE;
    }

    if(g_ole_initialized == FALSE) {
        hr = OleInitialize(NULL);
        if(FAILED(hr)) {
            ole_raise(mrb, hr, E_RUNTIME_ERROR, "fail: OLE initialize");
        }
        g_ole_initialized_set(TRUE);

        hr = CoRegisterMessageFilter(&imessage_filter, &previous_filter);
        if(FAILED(hr)) {
            previous_filter = NULL;
            ole_raise(mrb, hr, E_RUNTIME_ERROR, "fail: install OLE MessageFilter");
        }
    }
}

static void
ole_free(mrb_state *mrb, void *ptr)
{
    struct oledata *pole = (struct oledata *)ptr;
    OLE_FREE(pole->pDispatch);
    mrb_free(mrb, pole);
}

BSTR
ole_vstr2wc(mrb_state *mrb, mrb_value vstr)
{
    mrb_encoding *enc = mrb_enc_get(mrb, vstr);
    int cp;
    UINT size = 0;
    LPWSTR pw;
    mrb_value data = mrb_hash_fetch(mrb, enc2cp_table, mrb_cptr_value(mrb, enc), mrb_undef_value());

    if (!mrb_undef_p(data)) {
        cp = mrb_fixnum(data);
    } else {
        cp = ole_encoding2cp(enc);
        if (code_page_installed(cp) ||
            cp == CP_ACP ||
            cp == CP_OEMCP ||
            cp == CP_MACCP ||
            cp == CP_THREAD_ACP ||
            cp == CP_SYMBOL ||
            cp == CP_UTF7 ||
            cp == CP_UTF8 ||
            cp == 51932) {
            mrb_hash_set(mrb, enc2cp_table, mrb_cptr_value(mrb, enc), mrb_fixnum_value(cp));
        } else {
            mrb_raisef(mrb, E_WIN32OLE_RUNTIME_ERROR, "not installed Windows codepage(%S) according to `%S'", mrb_fixnum_value(cp), mrb_str_new_cstr(mrb, mrb_enc_name(enc)));
        }
    }
    if (conv_51932(mrb, cp)) {
#ifndef pIMultiLanguage
	DWORD dw = 0;
	UINT len = RSTRING_LENINT(vstr);
	HRESULT hr = pIMultiLanguage->lpVtbl->ConvertStringToUnicode(pIMultiLanguage,
		&dw, cp, RSTRING_PTR(vstr), &len, NULL, &size);
	if (FAILED(hr)) {
            ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "fail to convert CP%d to Unicode", cp);
	}
	pw = SysAllocStringLen(NULL, size);
	len = RSTRING_LEN(vstr);
	hr = pIMultiLanguage->lpVtbl->ConvertStringToUnicode(pIMultiLanguage,
		&dw, cp, RSTRING_PTR(vstr), &len, pw, &size);
	if (FAILED(hr)) {
            ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "fail to convert CP%d to Unicode", cp);
	}
#endif
	return pw;
    }
    size = MultiByteToWideChar(cp, 0, RSTRING_PTR(vstr), RSTRING_LEN(vstr), NULL, 0);
    pw = SysAllocStringLen(NULL, size);
    MultiByteToWideChar(cp, 0, RSTRING_PTR(vstr), RSTRING_LEN(vstr), pw, size);
    return pw;
}

static BSTR
ole_mb2wc(mrb_state *mrb, char *pm, int len)
{
    UINT size = 0;
    LPWSTR pw;

    if (conv_51932(mrb, cWIN32OLE_cp)) {
#ifndef pIMultiLanguage
	DWORD dw = 0;
	UINT n = len;
	HRESULT hr = pIMultiLanguage->lpVtbl->ConvertStringToUnicode(pIMultiLanguage,
		&dw, cWIN32OLE_cp, pm, &n, NULL, &size);
	if (FAILED(hr)) {
            ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "fail to convert CP%d to Unicode", cWIN32OLE_cp);
	}
	pw = SysAllocStringLen(NULL, size);
	hr = pIMultiLanguage->lpVtbl->ConvertStringToUnicode(pIMultiLanguage,
		&dw, cWIN32OLE_cp, pm, &n, pw, &size);
	if (FAILED(hr)) {
            ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "fail to convert CP%d to Unicode", cWIN32OLE_cp);
	}
#endif
	return pw;
    }
    size = MultiByteToWideChar(cWIN32OLE_cp, 0, pm, len, NULL, 0);
    pw = SysAllocStringLen(NULL, size - 1);
    MultiByteToWideChar(cWIN32OLE_cp, 0, pm, len, pw, size);
    return pw;
}

static char *
ole_alloc_vstr(mrb_state *mrb, UINT size, void *arg)
{
    mrb_value str = mrb_enc_str_new(mrb, NULL, size, cWIN32OLE_enc);
    *(mrb_value *)arg = str;
    return RSTRING_PTR(str);
}

mrb_value
ole_wc2vstr(mrb_state *mrb, LPWSTR pw, BOOL isfree)
{
    mrb_value vstr;
    ole_wc2mb_alloc(mrb, pw, ole_alloc_vstr, &vstr);
    mrb_str_resize(mrb, vstr, (long)strlen(RSTRING_PTR(vstr)));
    if(isfree)
        SysFreeString(pw);
    return vstr;
}

static mrb_value
ole_ary_m_entry(mrb_state *mrb, mrb_value val, LONG *pid)
{
    mrb_value obj = mrb_nil_value();
    int i = 0;
    obj = val;
    while(mrb_array_p(obj)) {
        obj = mrb_ary_entry(obj, pid[i]);
        i++;
    }
    return obj;
}

static mrb_value
is_all_index_under(LONG *pid, long *pub, long dim)
{
  long i = 0;
  for (i = 0; i < dim; i++) {
    if (pid[i] > pub[i]) {
      return mrb_false_value();
    }
  }
  return mrb_true_value();
}

void
ole_val2variant_ex(mrb_state *mrb, mrb_value val, VARIANT *var, VARTYPE vt)
{
    if (mrb_nil_p(val)) {
        if (vt == VT_VARIANT) {
            ole_val2variant2(mrb, val, var);
        } else {
            V_VT(var) = (vt & ~VT_BYREF);
            if (V_VT(var) == VT_DISPATCH) {
                V_DISPATCH(var) = NULL;
            } else if (V_VT(var) == VT_UNKNOWN) {
                V_UNKNOWN(var) = NULL;
            }
        }
        return;
    }
#if (_MSC_VER >= 1300) || defined(__CYGWIN__) || defined(__MINGW32__)
    switch(vt & ~VT_BYREF) {
    case VT_I8:
        V_VT(var) = VT_I8;
        V_I8(var) = NUM2I8 (val);
        break;
    case VT_UI8:
        V_VT(var) = VT_UI8;
        V_UI8(var) = NUM2UI8(val);
        break;
    default:
        ole_val2variant2(mrb, val, var);
        break;
    }
#else  /* (_MSC_VER >= 1300) || defined(__CYGWIN__) || defined(__MINGW32__) */
    ole_val2variant2(mrb, val, var);
#endif
}

VOID *
val2variant_ptr(mrb_state *mrb, mrb_value val, VARIANT *var, VARTYPE vt)
{
    VOID *p = NULL;
    HRESULT hr = S_OK;
    ole_val2variant_ex(mrb, val, var, vt);
    if ((vt & ~VT_BYREF) == VT_VARIANT) {
        p = var;
    } else {
        if ( (vt & ~VT_BYREF) != V_VT(var)) {
            hr = VariantChangeTypeEx(var, var,
                    cWIN32OLE_lcid, 0, (VARTYPE)(vt & ~VT_BYREF));
            if (FAILED(hr)) {
                ole_raise(mrb, hr, E_RUNTIME_ERROR, "failed to change type");
            }
        }
        p = get_ptr_of_variant(var);
    }
    if (p == NULL) {
        mrb_raise(mrb, E_RUNTIME_ERROR, "failed to get pointer of variant");
    }
    return p;
}

static void *
get_ptr_of_variant(VARIANT *pvar)
{
    switch(V_VT(pvar)) {
    case VT_UI1:
        return &V_UI1(pvar);
        break;
    case VT_I2:
        return &V_I2(pvar);
        break;
    case VT_UI2:
        return &V_UI2(pvar);
        break;
    case VT_I4:
        return &V_I4(pvar);
        break;
    case VT_UI4:
        return &V_UI4(pvar);
        break;
    case VT_R4:
        return &V_R4(pvar);
        break;
    case VT_R8:
        return &V_R8(pvar);
        break;
#if (_MSC_VER >= 1300) || defined(__CYGWIN__) || defined(__MINGW32__)
    case VT_I8:
        return &V_I8(pvar);
        break;
    case VT_UI8:
        return &V_UI8(pvar);
        break;
#endif
    case VT_INT:
        return &V_INT(pvar);
        break;
    case VT_UINT:
        return &V_UINT(pvar);
        break;
    case VT_CY:
        return &V_CY(pvar);
        break;
    case VT_DATE:
        return &V_DATE(pvar);
        break;
    case VT_BSTR:
        return V_BSTR(pvar);
        break;
    case VT_DISPATCH:
        return V_DISPATCH(pvar);
        break;
    case VT_ERROR:
        return &V_ERROR(pvar);
        break;
    case VT_BOOL:
        return &V_BOOL(pvar);
        break;
    case VT_UNKNOWN:
        return V_UNKNOWN(pvar);
        break;
    case VT_ARRAY:
        return &V_ARRAY(pvar);
        break;
    default:
        return NULL;
        break;
    }
}

static void
ole_set_safe_array(mrb_state *mrb, long n, SAFEARRAY *psa, LONG *pid, long *pub, mrb_value val, long dim,  VARTYPE vt)
{
    mrb_value val1;
    HRESULT hr = S_OK;
    VARIANT var;
    VOID *p = NULL;
    long i = n;
    while(i >= 0) {
        val1 = ole_ary_m_entry(mrb, val, pid);
        VariantInit(&var);
        p = val2variant_ptr(mrb, val1, &var, vt);
        if (mrb_bool(is_all_index_under(pid, pub, dim))) {
            if ((V_VT(&var) == VT_DISPATCH && V_DISPATCH(&var) == NULL) ||
                (V_VT(&var) == VT_UNKNOWN && V_UNKNOWN(&var) == NULL)) {
                mrb_raise(mrb, E_WIN32OLE_RUNTIME_ERROR, "element of array does not have IDispatch or IUnknown Interface");
            }
            hr = SafeArrayPutElement(psa, pid, p);
        }
        if (FAILED(hr)) {
            ole_raise(mrb, hr, E_RUNTIME_ERROR, "failed to SafeArrayPutElement");
        }
        pid[i] += 1;
        if (pid[i] > pub[i]) {
            pid[i] = 0;
            i -= 1;
        } else {
            i = dim - 1;
        }
    }
}

static long
dimension(mrb_value val) {
    long dim = 0;
    long dim1 = 0;
    long len = 0;
    long i = 0;
    if (mrb_array_p(val)) {
        len = RARRAY_LEN(val);
        for (i = 0; i < len; i++) {
            dim1 = dimension(mrb_ary_entry(val, i));
            if (dim < dim1) {
                dim = dim1;
            }
        }
        dim += 1;
    }
    return dim;
}

static long
ary_len_of_dim(mrb_value ary, long dim) {
    long ary_len = 0;
    long ary_len1 = 0;
    long len = 0;
    long i = 0;
    mrb_value val;
    if (dim == 0) {
        if (mrb_array_p(ary)) {
            ary_len = RARRAY_LEN(ary);
        }
    } else {
        if (mrb_array_p(ary)) {
            len = RARRAY_LEN(ary);
            for (i = 0; i < len; i++) {
                val = mrb_ary_entry(ary, i);
                ary_len1 = ary_len_of_dim(val, dim-1);
                if (ary_len < ary_len1) {
                    ary_len = ary_len1;
                }
            }
        }
    }
    return ary_len;
}

HRESULT
ole_val_ary2variant_ary(mrb_state *mrb, mrb_value val, VARIANT *var, VARTYPE vt)
{
    long dim = 0;
    int  i = 0;
    HRESULT hr = S_OK;

    SAFEARRAYBOUND *psab = NULL;
    SAFEARRAY *psa = NULL;
    long      *pub;
    LONG      *pid;

    mrb_check_type(mrb, val, MRB_TT_ARRAY);

    dim = dimension(val);

    psab = ALLOC_N(SAFEARRAYBOUND, dim);
    pub  = ALLOC_N(long, dim);
    pid  = ALLOC_N(LONG, dim);

    if(!psab || !pub || !pid) {
        if(pub) free(pub);
        if(psab) free(psab);
        if(pid) free(pid);
        mrb_raise(mrb, E_RUNTIME_ERROR, "memory allocation error");
    }

    for (i = 0; i < dim; i++) {
        psab[i].cElements = ary_len_of_dim(val, i);
        psab[i].lLbound = 0;
        pub[i] = psab[i].cElements - 1;
        pid[i] = 0;
    }
    /* Create and fill VARIANT array */
    if ((vt & ~VT_BYREF) == VT_ARRAY) {
        vt = (vt | VT_VARIANT);
    }
    psa = SafeArrayCreate((VARTYPE)(vt & VT_TYPEMASK), dim, psab);
    if (psa == NULL)
        hr = E_OUTOFMEMORY;
    else
        hr = SafeArrayLock(psa);
    if (SUCCEEDED(hr)) {
        ole_set_safe_array(mrb, dim-1, psa, pid, pub, val, dim, (VARTYPE)(vt & VT_TYPEMASK));
        hr = SafeArrayUnlock(psa);
    }

    if(pub) free(pub);
    if(psab) free(psab);
    if(pid) free(pid);

    if (SUCCEEDED(hr)) {
        V_VT(var) = vt;
        V_ARRAY(var) = psa;
    }
    else {
        if (psa != NULL)
            SafeArrayDestroy(psa);
    }
    return hr;
}

void
ole_val2variant(mrb_state *mrb, mrb_value val, VARIANT *var)
{
    struct oledata *pole;
    if(mrb_obj_is_kind_of(mrb, val, C_WIN32OLE)) {
        Data_Get_Struct(mrb, val, &ole_datatype, pole);
        OLE_ADDREF(pole->pDispatch);
        V_VT(var) = VT_DISPATCH;
        V_DISPATCH(var) = pole->pDispatch;
        return;
    }
    if (mrb_obj_is_kind_of(mrb, val, C_WIN32OLE_VARIANT)) {
        ole_variant2variant(mrb, val, var);
        return;
    }
    if (mrb_obj_is_kind_of(mrb, val, C_WIN32OLE_RECORD)) {
        ole_rec2variant(mrb, val, var);
        return;
    }
    if (mrb_obj_is_kind_of(mrb, val, mrb_class_get(mrb, "Time"))) {
        V_VT(var) = VT_DATE;
        V_DATE(var) = rbtime2vtdate(mrb, val);
        return;
    }
    switch (mrb_type(val)) {
    case MRB_TT_ARRAY:
        ole_val_ary2variant_ary(mrb, val, var, VT_VARIANT|VT_ARRAY);
        break;
    case MRB_TT_STRING:
        V_VT(var) = VT_BSTR;
        V_BSTR(var) = ole_vstr2wc(mrb, val);
        break;
    case MRB_TT_FIXNUM:
        V_VT(var) = VT_I4;
        V_I4(var) = NUM2INT(val);
        break;
/*
    FIXME:
    case MRB_TT_BIGNUM:
        V_VT(var) = VT_R8;
        V_R8(var) = rb_big2dbl(val);
        break;
*/
    case MRB_TT_FLOAT:
        V_VT(var) = VT_R8;
        V_R8(var) = NUM2DBL(val);
        break;
    case MRB_TT_TRUE:
        V_VT(var) = VT_BOOL;
        V_BOOL(var) = VARIANT_TRUE;
        break;
    case MRB_TT_FALSE:
        V_VT(var) = VT_BOOL;
        V_BOOL(var) = VARIANT_FALSE;
        break;
    default:
        if (mrb_nil_p(val)) {
            if (g_nil_to == VT_ERROR) {
                V_VT(var) = VT_ERROR;
                V_ERROR(var) = DISP_E_PARAMNOTFOUND;
            }else {
                V_VT(var) = VT_EMPTY;
            }
            break;
        } else {
            V_VT(var) = VT_DISPATCH;
            V_DISPATCH(var) = val2dispatch(mrb, val);
        }
        break;
    }
}

void
ole_val2variant2(mrb_state *mrb, mrb_value val, VARIANT *var)
{
    g_nil_to = VT_EMPTY;
    ole_val2variant(mrb, val, var);
    g_nil_to = VT_ERROR;
}

mrb_value
make_inspect(mrb_state *mrb, const char *class_name, mrb_value detail)
{
    mrb_value str;
    str = mrb_str_new_lit(mrb, "#<");
    mrb_str_cat_cstr(mrb, str, class_name);
    mrb_str_cat_lit(mrb, str, ":");
    mrb_str_concat(mrb, str, detail);
    mrb_str_cat_lit(mrb, str, ">");
    return str;
}

mrb_value
default_inspect(mrb_state *mrb, mrb_value self, const char *class_name)
{
    mrb_value detail = mrb_funcall(mrb, self, "to_s", 0);
    return make_inspect(mrb, class_name, detail);
}

static mrb_value
ole_set_member(mrb_state *mrb, mrb_value self, IDispatch *dispatch)
{
    struct oledata *pole;
    Data_Get_Struct(mrb, self, &ole_datatype, pole);
    if (pole->pDispatch) {
        OLE_RELEASE(pole->pDispatch);
        pole->pDispatch = NULL;
    }
    pole->pDispatch = dispatch;
    return self;
}

static struct oledata *
oledata_alloc(mrb_state *mrb)
{
    struct oledata *pole = ALLOC(struct oledata);
    ole_initialize(mrb);
    pole->pDispatch = NULL;
	return pole;
}

static mrb_value
fole_s_allocate(mrb_state *mrb, struct RClass *klass)
{
    mrb_value obj;
    struct oledata *pole = oledata_alloc(mrb);
    obj = mrb_obj_value(Data_Wrap_Struct(mrb, klass,&ole_datatype,pole));
    return obj;
}

static mrb_value
create_win32ole_object(mrb_state *mrb, struct RClass *klass, IDispatch *pDispatch, int argc, mrb_value *argv)
{
    mrb_value obj = fole_s_allocate(mrb, klass);
    ole_set_member(mrb, obj, pDispatch);
    return obj;
}

static mrb_value
ary_new_dim(mrb_state *mrb, mrb_value myary, LONG *pid, LONG *plb, LONG dim) {
    long i;
    mrb_value obj = mrb_nil_value();
    mrb_value pobj = mrb_nil_value();
    long *ids = ALLOC_N(long, dim);
    if (!ids) {
        mrb_raise(mrb, E_RUNTIME_ERROR, "memory allocation error");
    }
    for(i = 0; i < dim; i++) {
        ids[i] = pid[i] - plb[i];
    }
    obj = myary;
    pobj = myary;
    for(i = 0; i < dim-1; i++) {
        obj = mrb_ary_entry(pobj, ids[i]);
        if (mrb_nil_p(obj)) {
            mrb_ary_set(mrb, pobj, ids[i], mrb_ary_new(mrb));
        }
        obj = mrb_ary_entry(pobj, ids[i]);
        pobj = obj;
    }
    if (ids) free(ids);
    return obj;
}

static void
ary_store_dim(mrb_state *mrb, mrb_value myary, LONG *pid, LONG *plb, LONG dim, mrb_value val) {
    long id = pid[dim - 1] - plb[dim - 1];
    mrb_value obj = ary_new_dim(mrb, myary, pid, plb, dim);
    mrb_ary_set(mrb, obj, id, val);
}

mrb_value
ole_variant2val(mrb_state *mrb, VARIANT *pvar)
{
    mrb_value obj = mrb_nil_value();
    VARTYPE vt = V_VT(pvar);
    HRESULT hr;
    while ( vt == (VT_BYREF | VT_VARIANT) ) {
        pvar = V_VARIANTREF(pvar);
        vt = V_VT(pvar);
    }

    if(V_ISARRAY(pvar)) {
        VARTYPE vt_base = vt & VT_TYPEMASK;
        SAFEARRAY *psa = V_ISBYREF(pvar) ? *V_ARRAYREF(pvar) : V_ARRAY(pvar);
        UINT i = 0;
        LONG *pid, *plb, *pub;
        VARIANT variant;
        mrb_value val;
        UINT dim = 0;
        if (!psa) {
            return obj;
        }
        dim = SafeArrayGetDim(psa);
        pid = ALLOC_N(LONG, dim);
        plb = ALLOC_N(LONG, dim);
        pub = ALLOC_N(LONG, dim);

        if(!pid || !plb || !pub) {
            if(pid) free(pid);
            if(plb) free(plb);
            if(pub) free(pub);
            mrb_raise(mrb, E_RUNTIME_ERROR, "memory allocation error");
        }

        for(i = 0; i < dim; ++i) {
            SafeArrayGetLBound(psa, i+1, &plb[i]);
            SafeArrayGetLBound(psa, i+1, &pid[i]);
            SafeArrayGetUBound(psa, i+1, &pub[i]);
        }
        hr = SafeArrayLock(psa);
        if (SUCCEEDED(hr)) {
            obj = mrb_ary_new(mrb);
            i = 0;
            VariantInit(&variant);
            V_VT(&variant) = vt_base | VT_BYREF;
            if (vt_base == VT_RECORD) {
                hr = SafeArrayGetRecordInfo(psa, &V_RECORDINFO(&variant));
                if (SUCCEEDED(hr)) {
                    V_VT(&variant) = VT_RECORD;
                }
            }
            while (i < dim) {
                ary_new_dim(mrb, obj, pid, plb, dim);
                if (vt_base == VT_RECORD)
                    hr = SafeArrayPtrOfIndex(psa, pid, &V_RECORD(&variant));
                else
                    hr = SafeArrayPtrOfIndex(psa, pid, &V_BYREF(&variant));
                if (SUCCEEDED(hr)) {
                    val = ole_variant2val(mrb, &variant);
                    ary_store_dim(mrb, obj, pid, plb, dim, val);
                }
                for (i = 0; i < dim; ++i) {
                    if (++pid[i] <= pub[i])
                        break;
                    pid[i] = plb[i];
                }
            }
            SafeArrayUnlock(psa);
        }
        if(pid) free(pid);
        if(plb) free(plb);
        if(pub) free(pub);
        return obj;
    }
    switch(V_VT(pvar) & ~VT_BYREF){
    case VT_EMPTY:
        break;
    case VT_NULL:
        break;
    case VT_I1:
        if(V_ISBYREF(pvar))
            obj = INT2NUM((long)*V_I1REF(pvar));
        else
            obj = INT2NUM((long)V_I1(pvar));
        break;

    case VT_UI1:
        if(V_ISBYREF(pvar))
            obj = INT2NUM((long)*V_UI1REF(pvar));
        else
            obj = INT2NUM((long)V_UI1(pvar));
        break;

    case VT_I2:
        if(V_ISBYREF(pvar))
            obj = INT2NUM((long)*V_I2REF(pvar));
        else
            obj = INT2NUM((long)V_I2(pvar));
        break;

    case VT_UI2:
        if(V_ISBYREF(pvar))
            obj = INT2NUM((long)*V_UI2REF(pvar));
        else
            obj = INT2NUM((long)V_UI2(pvar));
        break;

    case VT_I4:
        if(V_ISBYREF(pvar))
            obj = INT2NUM((long)*V_I4REF(pvar));
        else
            obj = INT2NUM((long)V_I4(pvar));
        break;

    case VT_UI4:
        if(V_ISBYREF(pvar))
            obj = INT2NUM((long)*V_UI4REF(pvar));
        else
            obj = INT2NUM((long)V_UI4(pvar));
        break;

    case VT_INT:
        if(V_ISBYREF(pvar))
            obj = INT2NUM((long)*V_INTREF(pvar));
        else
            obj = INT2NUM((long)V_INT(pvar));
        break;

    case VT_UINT:
        if(V_ISBYREF(pvar))
            obj = INT2NUM((long)*V_UINTREF(pvar));
        else
            obj = INT2NUM((long)V_UINT(pvar));
        break;

#if (_MSC_VER >= 1300) || defined(__CYGWIN__) || defined(__MINGW32__)
    case VT_I8:
        if(V_ISBYREF(pvar))
#if (_MSC_VER >= 1300) || defined(__CYGWIN__) || defined(__MINGW32__)
#ifdef V_I8REF
            obj = I8_2_NUM(*V_I8REF(pvar));
#endif
#else
            obj = mrb_nil_value();
#endif
        else
            obj = I8_2_NUM(V_I8(pvar));
        break;
    case VT_UI8:
        if(V_ISBYREF(pvar))
#if (_MSC_VER >= 1300) || defined(__CYGWIN__) || defined(__MINGW32__)
#ifdef V_UI8REF
            obj = UI8_2_NUM(*V_UI8REF(pvar));
#endif
#else
            obj = mrb_nil_value();
#endif
        else
            obj = UI8_2_NUM(V_UI8(pvar));
        break;
#endif  /* (_MSC_VER >= 1300) || defined(__CYGWIN__) || defined(__MINGW32__) */

    case VT_R4:
        if(V_ISBYREF(pvar))
            obj = mrb_float_value(mrb, *V_R4REF(pvar));
        else
            obj = mrb_float_value(mrb, V_R4(pvar));
        break;

    case VT_R8:
        if(V_ISBYREF(pvar))
            obj = mrb_float_value(mrb, *V_R8REF(pvar));
        else
            obj = mrb_float_value(mrb, V_R8(pvar));
        break;

    case VT_BSTR:
    {
        if(V_ISBYREF(pvar))
            obj = ole_wc2vstr(mrb, *V_BSTRREF(pvar), FALSE);
        else
            obj = ole_wc2vstr(mrb, V_BSTR(pvar), FALSE);
        break;
    }

    case VT_ERROR:
        if(V_ISBYREF(pvar))
            obj = INT2NUM(*V_ERRORREF(pvar));
        else
            obj = INT2NUM(V_ERROR(pvar));
        break;

    case VT_BOOL:
        if (V_ISBYREF(pvar))
            obj = (*V_BOOLREF(pvar) ? mrb_true_value() : mrb_false_value());
        else
            obj = (V_BOOL(pvar) ? mrb_true_value() : mrb_false_value());
        break;

    case VT_DISPATCH:
    {
        IDispatch *pDispatch;

        if (V_ISBYREF(pvar))
            pDispatch = *V_DISPATCHREF(pvar);
        else
            pDispatch = V_DISPATCH(pvar);

        if (pDispatch != NULL ) {
            OLE_ADDREF(pDispatch);
            obj = create_win32ole_object(mrb, C_WIN32OLE, pDispatch, 0, 0);
        }
        break;
    }

    case VT_UNKNOWN:
    {
        /* get IDispatch interface from IUnknown interface */
        IUnknown *punk;
        IDispatch *pDispatch;
        void *p;
        HRESULT hr;

        if (V_ISBYREF(pvar))
            punk = *V_UNKNOWNREF(pvar);
        else
            punk = V_UNKNOWN(pvar);

        if(punk != NULL) {
           hr = punk->lpVtbl->QueryInterface(punk, &IID_IDispatch, &p);
           if(SUCCEEDED(hr)) {
               pDispatch = p;
               obj = create_win32ole_object(mrb, C_WIN32OLE, pDispatch, 0, 0);
           }
        }
        break;
    }

    case VT_DATE:
    {
        DATE date;
        if(V_ISBYREF(pvar))
            date = *V_DATEREF(pvar);
        else
            date = V_DATE(pvar);

        obj =  vtdate2rbtime(mrb, date);
        break;
    }

    case VT_RECORD:
    {
        IRecordInfo *pri = V_RECORDINFO(pvar);
        void *prec = V_RECORD(pvar);
        obj = create_win32ole_record(mrb, pri, prec);
        break;
    }

    case VT_CY:
    default:
        {
        HRESULT hr;
        VARIANT variant;
        VariantInit(&variant);
        hr = VariantChangeTypeEx(&variant, pvar,
                                  cWIN32OLE_lcid, 0, VT_BSTR);
        if (SUCCEEDED(hr) && V_VT(&variant) == VT_BSTR) {
            obj = ole_wc2vstr(mrb, V_BSTR(&variant), FALSE);
        }
        VariantClear(&variant);
        break;
        }
    }
    return obj;
}

LONG
reg_open_key(HKEY hkey, const char *name, HKEY *phkey)
{
    return RegOpenKeyExA(hkey, name, 0, KEY_READ, phkey);
}

LONG
reg_open_vkey(mrb_state *mrb, HKEY hkey, mrb_value key, HKEY *phkey)
{
    return reg_open_key(hkey, mrb_string_value_ptr(mrb, key), phkey);
}

mrb_value
reg_enum_key(mrb_state *mrb, HKEY hkey, DWORD i)
{
    char buf[BUFSIZ + 1];
    DWORD size_buf = sizeof(buf);
    FILETIME ft;
    LONG err = RegEnumKeyExA(hkey, i, buf, &size_buf,
                            NULL, NULL, NULL, &ft);
    if(err == ERROR_SUCCESS) {
        buf[BUFSIZ] = '\0';
        return mrb_str_new_cstr(mrb, buf);
    }
    return mrb_nil_value();
}

mrb_value
reg_get_val(mrb_state *mrb, HKEY hkey, const char *subkey)
{
    char *pbuf;
    DWORD dwtype = 0;
    DWORD size = 0;
    mrb_value val = mrb_nil_value();
    LONG err = RegQueryValueExA(hkey, subkey, NULL, &dwtype, NULL, &size);

    if (err == ERROR_SUCCESS) {
        pbuf = ALLOC_N(char, size + 1);
        err = RegQueryValueExA(hkey, subkey, NULL, &dwtype, (BYTE *)pbuf, &size);
        if (err == ERROR_SUCCESS) {
            pbuf[size] = '\0';
            if (dwtype == REG_EXPAND_SZ) {
		char* pbuf2 = (char *)pbuf;
		DWORD len = ExpandEnvironmentStringsA(pbuf2, NULL, 0);
		pbuf = ALLOC_N(char, len + 1);
		ExpandEnvironmentStringsA(pbuf2, pbuf, len + 1);
		free(pbuf2);
            }
            val = mrb_str_new_cstr(mrb, (char *)pbuf);
        }
        free(pbuf);
    }
    return val;
}

mrb_value
reg_get_val2(mrb_state *mrb, HKEY hkey, const char *subkey)
{
    HKEY hsubkey;
    LONG err;
    mrb_value val = mrb_nil_value();
    err = RegOpenKeyExA(hkey, subkey, 0, KEY_READ, &hsubkey);
    if (err == ERROR_SUCCESS) {
        val = reg_get_val(mrb, hsubkey, NULL);
        RegCloseKey(hsubkey);
    }
    if (mrb_nil_p(val)) {
        val = reg_get_val(mrb, hkey, subkey);
    }
    return val;
}

static void
ole_const_load(mrb_state *mrb, ITypeLib *pTypeLib, mrb_value klass, mrb_value self)
{
    unsigned int count;
    unsigned int index;
    int iVar;
    ITypeInfo *pTypeInfo;
    TYPEATTR  *pTypeAttr;
    VARDESC   *pVarDesc;
    HRESULT hr;
    unsigned int len;
    BSTR bstr;
    char *pName = NULL;
    mrb_value val;
    mrb_value constant;
    constant = mrb_hash_new(mrb);
    count = pTypeLib->lpVtbl->GetTypeInfoCount(pTypeLib);
    for (index = 0; index < count; index++) {
        hr = pTypeLib->lpVtbl->GetTypeInfo(pTypeLib, index, &pTypeInfo);
        if (FAILED(hr))
            continue;
        hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
        if(FAILED(hr)) {
            OLE_RELEASE(pTypeInfo);
            continue;
        }
        for(iVar = 0; iVar < pTypeAttr->cVars; iVar++) {
            hr = pTypeInfo->lpVtbl->GetVarDesc(pTypeInfo, iVar, &pVarDesc);
            if(FAILED(hr))
                continue;
            if(pVarDesc->varkind == VAR_CONST &&
               !(pVarDesc->wVarFlags & (VARFLAG_FHIDDEN |
                                        VARFLAG_FRESTRICTED |
                                        VARFLAG_FNONBROWSABLE))) {
                hr = pTypeInfo->lpVtbl->GetNames(pTypeInfo, pVarDesc->memid, &bstr,
                                                 1, &len);
                if(FAILED(hr) || len == 0 || !bstr)
                    continue;
                pName = ole_wc2mb(mrb, bstr);
                val = ole_variant2val(mrb, V_UNION1(pVarDesc, lpvarValue));
                *pName = toupper((int)*pName);
                if (isupper((int)(unsigned char)*pName)) {
                    mrb_define_const(mrb, mrb_class_ptr(klass), pName, val);
                }
                else {
                    mrb_hash_set(mrb, constant, mrb_str_new_cstr(mrb, pName), val);
                }
                SysFreeString(bstr);
                if(pName) {
                    free(pName);
                    pName = NULL;
                }
            }
            pTypeInfo->lpVtbl->ReleaseVarDesc(pTypeInfo, pVarDesc);
        }
        pTypeInfo->lpVtbl->ReleaseTypeAttr(pTypeInfo, pTypeAttr);
        OLE_RELEASE(pTypeInfo);
    }
    mrb_define_const(mrb, mrb_class_ptr(klass), "CONSTANTS", constant);
}

static HRESULT
clsid_from_remote(mrb_state *mrb, mrb_value host, mrb_value com, CLSID *pclsid)
{
    HKEY hlm;
    HKEY hpid;
    mrb_value subkey;
    LONG err;
    char clsid[100];
    BSTR pbuf;
    DWORD len;
    DWORD dwtype;
    HRESULT hr = S_OK;
    err = RegConnectRegistryA(mrb_string_value_ptr(mrb, host), HKEY_LOCAL_MACHINE, &hlm);
    if (err != ERROR_SUCCESS)
        return HRESULT_FROM_WIN32(err);
    subkey = mrb_str_new_lit(mrb, "SOFTWARE\\Classes\\");
    mrb_str_concat(mrb, subkey, com);
    mrb_str_cat_lit(mrb, subkey, "\\CLSID");
    err = RegOpenKeyExA(hlm, mrb_string_value_ptr(mrb, subkey), 0, KEY_READ, &hpid);
    if (err != ERROR_SUCCESS)
        hr = HRESULT_FROM_WIN32(err);
    else {
        len = sizeof(clsid);
        err = RegQueryValueExA(hpid, "", NULL, &dwtype, (BYTE *)clsid, &len);
        if (err == ERROR_SUCCESS && dwtype == REG_SZ) {
            pbuf  = ole_mb2wc(mrb, clsid, -1);
            hr = CLSIDFromString(pbuf, pclsid);
            SysFreeString(pbuf);
        }
        else {
            hr = HRESULT_FROM_WIN32(err);
        }
        RegCloseKey(hpid);
    }
    RegCloseKey(hlm);
    return hr;
}

static mrb_value
ole_create_dcom(mrb_state *mrb, mrb_value self, mrb_value ole, mrb_value host, mrb_value others)
{
    HRESULT hr;
    CLSID   clsid;
    BSTR pbuf;

    COSERVERINFO serverinfo;
    MULTI_QI multi_qi;
    DWORD clsctx = CLSCTX_REMOTE_SERVER;

    if (!gole32)
        gole32 = LoadLibraryA("OLE32");
    if (!gole32)
        mrb_raise(mrb, E_RUNTIME_ERROR, "failed to load OLE32");
    if (!gCoCreateInstanceEx)
        gCoCreateInstanceEx = (FNCOCREATEINSTANCEEX*)
            GetProcAddress(gole32, "CoCreateInstanceEx");
    if (!gCoCreateInstanceEx)
        mrb_raise(mrb, E_RUNTIME_ERROR, "CoCreateInstanceEx is not supported in this environment");

    pbuf  = ole_vstr2wc(mrb, ole);
    hr = CLSIDFromProgID(pbuf, &clsid);
    if (FAILED(hr))
        hr = clsid_from_remote(mrb, host, ole, &clsid);
    if (FAILED(hr))
        hr = CLSIDFromString(pbuf, &clsid);
    SysFreeString(pbuf);
    if (FAILED(hr))
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR,
                  "unknown OLE server: `%s'",
                  mrb_string_value_ptr(mrb, ole));
    memset(&serverinfo, 0, sizeof(COSERVERINFO));
    serverinfo.pwszName = ole_vstr2wc(mrb, host);
    memset(&multi_qi, 0, sizeof(MULTI_QI));
    multi_qi.pIID = &IID_IDispatch;
    hr = gCoCreateInstanceEx(&clsid, NULL, clsctx, &serverinfo, 1, &multi_qi);
    SysFreeString(serverinfo.pwszName);
    if (FAILED(hr))
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR,
                  "failed to create DCOM server `%s' in `%s'",
                  mrb_string_value_ptr(mrb, ole),
                  mrb_string_value_ptr(mrb, host));

    ole_set_member(mrb, self, (IDispatch*)multi_qi.pItf);
    return self;
}

static mrb_value
ole_bind_obj(mrb_state *mrb, mrb_value moniker, int argc, mrb_value *argv, mrb_value self)
{
    IBindCtx *pBindCtx;
    IMoniker *pMoniker;
    IDispatch *pDispatch;
    void *p;
    HRESULT hr;
    BSTR pbuf;
    ULONG eaten = 0;

    ole_initialize(mrb);

    hr = CreateBindCtx(0, &pBindCtx);
    if(FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR,
                  "failed to create bind context");
    }

    pbuf  = ole_vstr2wc(mrb, moniker);
    hr = MkParseDisplayName(pBindCtx, pbuf, &eaten, &pMoniker);
    SysFreeString(pbuf);
    if(FAILED(hr)) {
        OLE_RELEASE(pBindCtx);
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR,
                  "failed to parse display name of moniker `%s'",
                  mrb_string_value_ptr(mrb, moniker));
    }
    hr = pMoniker->lpVtbl->BindToObject(pMoniker, pBindCtx, NULL,
                                        &IID_IDispatch, &p);
    pDispatch = p;
    OLE_RELEASE(pMoniker);
    OLE_RELEASE(pBindCtx);

    if(FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR,
                  "failed to bind moniker `%s'",
                  mrb_string_value_ptr(mrb, moniker));
    }
    return create_win32ole_object(mrb, mrb_class_ptr(self), pDispatch, argc, argv);
}

/*
 *  call-seq:
 *     WIN32OLE.connect( ole ) --> aWIN32OLE
 *
 *  Returns running OLE Automation object or WIN32OLE object from moniker.
 *  1st argument should be OLE program id or class id or moniker.
 *
 *     WIN32OLE.connect('Excel.Application') # => WIN32OLE object which represents running Excel.
 */
static mrb_value
fole_s_connect(mrb_state *mrb, mrb_value self)
{
    int argc;
    mrb_value *argv;
    mrb_value svr_name;
    HRESULT hr;
    CLSID   clsid;
    BSTR pBuf;
    IDispatch *pDispatch;
    void *p;
    IUnknown *pUnknown;

    mrb_get_args(mrb, "S", &svr_name);
    mrb_get_args(mrb, "*", &argv, &argc);

    /* initialize to use OLE */
    ole_initialize(mrb);
/*
    if (rb_safe_level() > 0 && OBJ_TAINTED(svr_name)) {
        mrb_raise(mrb, rb_eSecurityError, "insecure connection - `%s'",
		mrb_string_value_ptr(mrb, svr_name));
    }
*/

    /* get CLSID from OLE server name */
    pBuf = ole_vstr2wc(mrb, svr_name);
    hr = CLSIDFromProgID(pBuf, &clsid);
    if(FAILED(hr)) {
        hr = CLSIDFromString(pBuf, &clsid);
    }
    SysFreeString(pBuf);
    if(FAILED(hr)) {
        return ole_bind_obj(mrb, svr_name, argc, argv, self);
    }

    hr = GetActiveObject(&clsid, 0, &pUnknown);
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR,
                  "OLE server `%s' not running", mrb_string_value_ptr(mrb, svr_name));
    }
    hr = pUnknown->lpVtbl->QueryInterface(pUnknown, &IID_IDispatch, &p);
    pDispatch = p;
    if(FAILED(hr)) {
        OLE_RELEASE(pUnknown);
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR,
                  "failed to create WIN32OLE server `%s'",
                  mrb_string_value_ptr(mrb, svr_name));
    }

    OLE_RELEASE(pUnknown);

    return create_win32ole_object(mrb, mrb_class_ptr(self), pDispatch, argc, argv);
}

/*
 *  call-seq:
 *     WIN32OLE.const_load( ole, mod = WIN32OLE)
 *
 *  Defines the constants of OLE Automation server as mod's constants.
 *  The first argument is WIN32OLE object or type library name.
 *  If 2nd argument is omitted, the default is WIN32OLE.
 *  The first letter of Ruby's constant variable name is upper case,
 *  so constant variable name of WIN32OLE object is capitalized.
 *  For example, the 'xlTop' constant of Excel is changed to 'XlTop'
 *  in WIN32OLE.
 *  If the first letter of constant variable is not [A-Z], then
 *  the constant is defined as CONSTANTS hash element.
 *
 *     module EXCEL_CONST
 *     end
 *     excel = WIN32OLE.new('Excel.Application')
 *     WIN32OLE.const_load(excel, EXCEL_CONST)
 *     puts EXCEL_CONST::XlTop # => -4160
 *     puts EXCEL_CONST::CONSTANTS['_xlDialogChartSourceData'] # => 541
 *
 *     WIN32OLE.const_load(excel)
 *     puts WIN32OLE::XlTop # => -4160
 *
 *     module MSO
 *     end
 *     WIN32OLE.const_load('Microsoft Office 9.0 Object Library', MSO)
 *     puts MSO::MsoLineSingle # => 1
 */
static mrb_value
fole_s_const_load(mrb_state *mrb, mrb_value self)
{
    mrb_value ole;
    mrb_value klass = mrb_nil_value();
    struct oledata *pole;
    ITypeInfo *pTypeInfo;
    ITypeLib *pTypeLib;
    unsigned int index;
    HRESULT hr;
    BSTR pBuf;
    mrb_value file;
    LCID    lcid = cWIN32OLE_lcid;

    mrb_get_args(mrb, "o|o", &ole, &klass);
    if (mrb_type(klass) != MRB_TT_CLASS &&
        mrb_type(klass) != MRB_TT_MODULE &&
        !mrb_nil_p(klass)) {
        mrb_raise(mrb, E_TYPE_ERROR, "2nd parameter must be Class or Module");
    }
    if (mrb_obj_is_kind_of(mrb, ole, C_WIN32OLE)) {
        OLEData_Get_Struct(mrb, ole, pole);
        hr = pole->pDispatch->lpVtbl->GetTypeInfo(pole->pDispatch,
                                                  0, lcid, &pTypeInfo);
        if(FAILED(hr)) {
            ole_raise(mrb, hr, E_RUNTIME_ERROR, "failed to GetTypeInfo");
        }
        hr = pTypeInfo->lpVtbl->GetContainingTypeLib(pTypeInfo, &pTypeLib, &index);
        if(FAILED(hr)) {
            OLE_RELEASE(pTypeInfo);
            ole_raise(mrb, hr, E_RUNTIME_ERROR, "failed to GetContainingTypeLib");
        }
        OLE_RELEASE(pTypeInfo);
        if(!mrb_nil_p(klass)) {
            ole_const_load(mrb, pTypeLib, klass, self);
        }
        else {
            ole_const_load(mrb, pTypeLib, mrb_obj_value(C_WIN32OLE), self);
        }
        OLE_RELEASE(pTypeLib);
    }
    else if(mrb_string_p(ole)) {
        file = typelib_file(mrb, ole);
        if (mrb_nil_p(file)) {
            file = ole;
        }
        pBuf = ole_vstr2wc(mrb, file);
        hr = LoadTypeLibEx(pBuf, REGKIND_NONE, &pTypeLib);
        SysFreeString(pBuf);
        if (FAILED(hr))
          ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to LoadTypeLibEx");
        if(!mrb_nil_p(klass)) {
            ole_const_load(mrb, pTypeLib, klass, self);
        }
        else {
            ole_const_load(mrb, pTypeLib, mrb_obj_value(C_WIN32OLE), self);
        }
        OLE_RELEASE(pTypeLib);
    }
    else {
        mrb_raise(mrb, E_TYPE_ERROR, "1st parameter must be WIN32OLE instance");
    }
    return mrb_nil_value();
}

static ULONG
reference_count(struct oledata * pole)
{
    ULONG n = 0;
    if(pole->pDispatch) {
        OLE_ADDREF(pole->pDispatch);
        n = OLE_RELEASE(pole->pDispatch);
    }
    return n;
}

/*
 *  call-seq:
 *     WIN32OLE.ole_reference_count(aWIN32OLE) --> number
 *
 *  Returns reference counter of Dispatch interface of WIN32OLE object.
 *  You should not use this method because this method
 *  exists only for debugging WIN32OLE.
 */
static mrb_value
fole_s_reference_count(mrb_state *mrb, mrb_value self)
{
    mrb_value obj;
    struct oledata * pole;
    mrb_get_args(mrb, "o", &obj);
    OLEData_Get_Struct(mrb, obj, pole);
    return INT2NUM(reference_count(pole));
}

/*
 *  call-seq:
 *     WIN32OLE.ole_free(aWIN32OLE) --> number
 *
 *  Invokes Release method of Dispatch interface of WIN32OLE object.
 *  You should not use this method because this method
 *  exists only for debugging WIN32OLE.
 *  The return value is reference counter of OLE object.
 */
static mrb_value
fole_s_free(mrb_state *mrb, mrb_value self)
{
    mrb_value obj;
    ULONG n = 0;
    struct oledata * pole;
    mrb_get_args(mrb, "o", &obj);
    OLEData_Get_Struct(mrb, obj, pole);
    if(pole->pDispatch) {
        if (reference_count(pole) > 0) {
            n = OLE_RELEASE(pole->pDispatch);
        }
    }
    return INT2NUM(n);
}

static HWND
ole_show_help(mrb_state *mrb, mrb_value helpfile, mrb_value helpcontext)
{
    FNHTMLHELP *pfnHtmlHelp;
    HWND hwnd = 0;

    if(!ghhctrl)
        ghhctrl = LoadLibraryA("HHCTRL.OCX");
    if (!ghhctrl)
        return hwnd;
    pfnHtmlHelp = (FNHTMLHELP*)GetProcAddress(ghhctrl, "HtmlHelpA");
    if (!pfnHtmlHelp)
        return hwnd;
    hwnd = pfnHtmlHelp(GetDesktopWindow(), mrb_string_value_ptr(mrb, helpfile),
                    0x0f, NUM2INT(helpcontext));
    if (hwnd == 0)
        hwnd = pfnHtmlHelp(GetDesktopWindow(), mrb_string_value_ptr(mrb, helpfile),
                 0,  NUM2INT(helpcontext));
    return hwnd;
}

/*
 *  call-seq:
 *     WIN32OLE.ole_show_help(obj [,helpcontext])
 *
 *  Displays helpfile. The 1st argument specifies WIN32OLE_TYPE
 *  object or WIN32OLE_METHOD object or helpfile.
 *
 *     excel = WIN32OLE.new('Excel.Application')
 *     typeobj = excel.ole_type
 *     WIN32OLE.ole_show_help(typeobj)
 */
static mrb_value
fole_s_show_help(mrb_state *mrb, mrb_value self)
{
    mrb_value target;
    mrb_value helpcontext = mrb_nil_value();
    mrb_value helpfile;
    mrb_value name;
    HWND  hwnd;
    mrb_get_args(mrb, "o|o", &target, &helpcontext);
    if (mrb_obj_is_kind_of(mrb, target, C_WIN32OLE_TYPE) ||
        mrb_obj_is_kind_of(mrb, target, C_WIN32OLE_METHOD)) {
        helpfile = mrb_funcall(mrb, target, "helpfile", 0);
        if(strlen(mrb_string_value_ptr(mrb, helpfile)) == 0) {
            name = mrb_iv_get(mrb, target, mrb_intern_lit(mrb, "name"));
            mrb_raisef(mrb, E_RUNTIME_ERROR, "no helpfile of `%S'",
                      name);
        }
        helpcontext = mrb_funcall(mrb, target, "helpcontext", 0);
    } else {
        helpfile = target;
    }
    if (!mrb_string_p(helpfile)) {
        mrb_raise(mrb, E_TYPE_ERROR, "1st parameter must be (String|WIN32OLE_TYPE|WIN32OLE_METHOD)");
    }
    hwnd = ole_show_help(mrb, helpfile, helpcontext);
    if(hwnd == 0) {
        mrb_raisef(mrb, E_RUNTIME_ERROR, "failed to open help file `%S'",
                   helpfile);
    }
    return mrb_nil_value();
}

/*
 *  call-seq:
 *     WIN32OLE.codepage
 *
 *  Returns current codepage.
 *     WIN32OLE.codepage # => WIN32OLE::CP_ACP
 */
static mrb_value
fole_s_get_code_page(mrb_state *mrb, mrb_value self)
{
    return INT2FIX(cWIN32OLE_cp);
}

static BOOL CALLBACK
installed_code_page_proc(LPSTR str) {
    if (strtoul(str, NULL, 10) == g_cp_to_check) {
        g_cp_installed = TRUE;
        return FALSE;
    }
    return TRUE;
}

static BOOL
code_page_installed(UINT cp)
{
    g_cp_installed = FALSE;
    g_cp_to_check = cp;
    EnumSystemCodePagesA(installed_code_page_proc, CP_INSTALLED);
    return g_cp_installed;
}

/*
 *  call-seq:
 *     WIN32OLE.codepage = CP
 *
 *  Sets current codepage.
 *  The WIN32OLE.codepage is initialized according to
 *  Encoding.default_internal.
 *  If Encoding.default_internal is nil then WIN32OLE.codepage
 *  is initialized according to Encoding.default_external.
 *
 *     WIN32OLE.codepage = WIN32OLE::CP_UTF8
 *     WIN32OLE.codepage = 65001
 */
static mrb_value
fole_s_set_code_page(mrb_state *mrb, mrb_value self)
{
    UINT cp;
    mrb_get_args(mrb, "i", &cp);
    set_ole_codepage(mrb, cp);
    /*
     * Should this method return old codepage?
     */
    return mrb_nil_value();
}

/*
 *  call-seq:
 *     WIN32OLE.locale -> locale id.
 *
 *  Returns current locale id (lcid). The default locale is
 *  WIN32OLE::LOCALE_SYSTEM_DEFAULT.
 *
 *     lcid = WIN32OLE.locale
 */
static mrb_value
fole_s_get_locale(mrb_state *mrb, mrb_value self)
{
    return INT2FIX(cWIN32OLE_lcid);
}

static BOOL
CALLBACK installed_lcid_proc(LPSTR str)
{
    if (strcmp(str, g_lcid_to_check) == 0) {
        g_lcid_installed = TRUE;
        return FALSE;
    }
    return TRUE;
}

static BOOL
lcid_installed(LCID lcid)
{
    g_lcid_installed = FALSE;
    snprintf(g_lcid_to_check, sizeof(g_lcid_to_check), "%08lx", (unsigned long)lcid);
    EnumSystemLocalesA(installed_lcid_proc, LCID_INSTALLED);
    return g_lcid_installed;
}

/*
 *  call-seq:
 *     WIN32OLE.locale = lcid
 *
 *  Sets current locale id (lcid).
 *
 *     WIN32OLE.locale = 1033 # set locale English(U.S)
 *     obj = WIN32OLE_VARIANT.new("$100,000", WIN32OLE::VARIANT::VT_CY)
 *
 */
static mrb_value
fole_s_set_locale(mrb_state *mrb, mrb_value self)
{
    LCID lcid;
    mrb_get_args(mrb, "i", &lcid);
    if (lcid_installed(lcid)) {
        cWIN32OLE_lcid = lcid;
    } else {
        switch (lcid) {
        case LOCALE_SYSTEM_DEFAULT:
        case LOCALE_USER_DEFAULT:
            cWIN32OLE_lcid = lcid;
            break;
        default:
            mrb_raisef(mrb, E_WIN32OLE_RUNTIME_ERROR, "not installed locale: %S", INT2FIX(lcid));
        }
    }
    return mrb_nil_value();
}

/*
 *  call-seq:
 *     WIN32OLE.create_guid
 *
 *  Creates GUID.
 *     WIN32OLE.create_guid # => {1CB530F1-F6B1-404D-BCE6-1959BF91F4A8}
 */
static mrb_value
fole_s_create_guid(mrb_state *mrb, mrb_value self)
{
    GUID guid;
    HRESULT hr;
    OLECHAR bstr[80];
    int len = 0;
    hr = CoCreateGuid(&guid);
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to create GUID");
    }
    len = StringFromGUID2(&guid, bstr, sizeof(bstr)/sizeof(OLECHAR));
    if (len == 0) {
        mrb_raise(mrb, E_RUNTIME_ERROR, "failed to create GUID(buffer over)");
    }
    return ole_wc2vstr(mrb, bstr, FALSE);
}

/*
 * WIN32OLE.ole_initialize and WIN32OLE.ole_uninitialize
 * are used in win32ole.rb to fix the issue bug #2618 (ruby-core:27634).
 * You must not use these method.
 */

/* :nodoc: */
static mrb_value
fole_s_ole_initialize(mrb_state *mrb, mrb_value self)
{
    ole_initialize(mrb);
    return mrb_nil_value();
}

/* :nodoc: */
static mrb_value
fole_s_ole_uninitialize(mrb_state *mrb, mrb_value self)
{
    ole_uninitialize();
    return mrb_nil_value();
}

/*
 * Document-class: WIN32OLE
 *
 *   <code>WIN32OLE</code> objects represent OLE Automation object in Ruby.
 *
 *   By using WIN32OLE, you can access OLE server like VBScript.
 *
 *   Here is sample script.
 *
 *     require 'win32ole'
 *
 *     excel = WIN32OLE.new('Excel.Application')
 *     excel.visible = true
 *     workbook = excel.Workbooks.Add();
 *     worksheet = workbook.Worksheets(1);
 *     worksheet.Range("A1:D1").value = ["North","South","East","West"];
 *     worksheet.Range("A2:B2").value = [5.2, 10];
 *     worksheet.Range("C2").value = 8;
 *     worksheet.Range("D2").value = 20;
 *
 *     range = worksheet.Range("A1:D2");
 *     range.select
 *     chart = workbook.Charts.Add;
 *
 *     workbook.saved = true;
 *
 *     excel.ActiveWorkbook.Close(0);
 *     excel.Quit();
 *
 *   Unfortunately, Win32OLE doesn't support the argument passed by
 *   reference directly.
 *   Instead, Win32OLE provides WIN32OLE::ARGV or WIN32OLE_VARIANT object.
 *   If you want to get the result value of argument passed by reference,
 *   you can use WIN32OLE::ARGV or WIN32OLE_VARIANT.
 *
 *     oleobj.method(arg1, arg2, refargv3)
 *     puts WIN32OLE::ARGV[2]   # the value of refargv3 after called oleobj.method
 *
 *   or
 *
 *     refargv3 = WIN32OLE_VARIANT.new(XXX,
 *                 WIN32OLE::VARIANT::VT_BYREF|WIN32OLE::VARIANT::VT_XXX)
 *     oleobj.method(arg1, arg2, refargv3)
 *     p refargv3.value # the value of refargv3 after called oleobj.method.
 *
 */

/*
 *  call-seq:
 *     WIN32OLE.new(server, [host]) -> WIN32OLE object
 *
 *  Returns a new WIN32OLE object(OLE Automation object).
 *  The first argument server specifies OLE Automation server.
 *  The first argument should be CLSID or PROGID.
 *  If second argument host specified, then returns OLE Automation
 *  object on host.
 *
 *      WIN32OLE.new('Excel.Application') # => Excel OLE Automation WIN32OLE object.
 *      WIN32OLE.new('{00024500-0000-0000-C000-000000000046}') # => Excel OLE Automation WIN32OLE object.
 */
static mrb_value
fole_initialize(mrb_state *mrb, mrb_value self)
{
    int argc = 0;
    mrb_value *argv;
    mrb_value svr_name;
    mrb_value host = mrb_nil_value();
    mrb_value others;
    HRESULT hr;
    CLSID   clsid;
    BSTR pBuf;
    IDispatch *pDispatch;
    void *p;
	struct oledata *pole = (struct oledata *)DATA_PTR(self);
    if (pole) {
        mrb_free(mrb, pole);
    }
    mrb_data_init(self, NULL, &ole_datatype);

    /* FIXME: why call rb_call_super()? 
    rb_call_super(0, 0);
    */
    mrb_get_args(mrb, "S|S*", &svr_name, &host, &argv, &argc);
    others = mrb_ary_new_from_values(mrb, argc, argv);

    pole = oledata_alloc(mrb);
    mrb_data_init(self, pole, &ole_datatype);

/*
 FIXME:
    if (rb_safe_level() > 0 && OBJ_TAINTED(svr_name)) {
        mrb_raise(mrb, rb_eSecurityError, "insecure object creation - `%s'",
                 mrb_string_value_ptr(mrb, svr_name));
    }
*/
    if (!mrb_nil_p(host)) {
/*
 FIXME:
        if (rb_safe_level() > 0 && OBJ_TAINTED(host)) {
            mrb_raise(mrb, rb_eSecurityError, "insecure object creation - `%s'",
                     mrb_string_value_ptr(mrb, host));
        }
*/
        return ole_create_dcom(mrb, self, svr_name, host, others);
    }

    /* get CLSID from OLE server name */
    pBuf  = ole_vstr2wc(mrb, svr_name);
    hr = CLSIDFromProgID(pBuf, &clsid);
    if(FAILED(hr)) {
        hr = CLSIDFromString(pBuf, &clsid);
    }
    SysFreeString(pBuf);
    if(FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR,
                  "unknown OLE server: `%s'",
                  mrb_string_value_ptr(mrb, svr_name));
    }

    /* get IDispatch interface */
    hr = CoCreateInstance(&clsid, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
                          &IID_IDispatch, &p);
    pDispatch = p;
    if(FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR,
                  "failed to create WIN32OLE object from `%s'",
                  mrb_string_value_ptr(mrb, svr_name));
    }

    ole_set_member(mrb, self, pDispatch);
    return self;
}

static int
hash2named_arg(mrb_state *mrb, mrb_value key, mrb_value val, struct oleparam* pOp)
{
    unsigned int index, i;
    index = pOp->dp.cNamedArgs;
    /*---------------------------------------------
      the data-type of key must be String or Symbol
    -----------------------------------------------*/
    if(!mrb_string_p(key) && !mrb_symbol_p(key)) {
        /* clear name of dispatch parameters */
        for(i = 1; i < index + 1; i++) {
            SysFreeString(pOp->pNamedArgs[i]);
        }
        /* clear dispatch parameters */
        for(i = 0; i < index; i++ ) {
            VariantClear(&(pOp->dp.rgvarg[i]));
        }
        /* raise an exception */
        mrb_raise(mrb, E_TYPE_ERROR, "wrong argument type (expected String or Symbol)");
    }
    if (mrb_symbol_p(key)) {
	key = mrb_sym2str(mrb, mrb_symbol(key));
    }

    /* pNamedArgs[0] is <method name>, so "index + 1" */
    pOp->pNamedArgs[index + 1] = ole_vstr2wc(mrb, key);

    VariantInit(&(pOp->dp.rgvarg[index]));
    ole_val2variant(mrb, val, &(pOp->dp.rgvarg[index]));

    pOp->dp.cNamedArgs += 1;
    return 0;
}

static mrb_value
set_argv(mrb_state *mrb, VARIANTARG* realargs, unsigned int beg, unsigned int end)
{
    mrb_value argv = mrb_const_get(mrb, mrb_obj_value(C_WIN32OLE), mrb_intern_lit(mrb, "ARGV"));

    mrb_check_type(mrb, argv, MRB_TT_ARRAY);
    mrb_ary_clear(mrb, argv);
    while (end-- > beg) {
        mrb_ary_push(mrb, argv, ole_variant2val(mrb, &realargs[end]));
        if (V_VT(&realargs[end]) != VT_RECORD) {
            VariantClear(&realargs[end]);
        }
    }
    return argv;
}

static mrb_value
ole_invoke(mrb_state *mrb, int argc, mrb_value *argv, mrb_value self, USHORT wFlags, BOOL is_bracket)
{
    LCID    lcid = cWIN32OLE_lcid;
    struct oledata *pole;
    HRESULT hr;
    mrb_value cmd;
    mrb_value paramS;
    mrb_value param;
    mrb_value obj;
    mrb_value v;

    BSTR wcmdname;

    DISPID DispID;
    DISPID* pDispID;
    EXCEPINFO excepinfo;
    VARIANT result;
    VARIANTARG* realargs = NULL;
    unsigned int argErr = 0;
    unsigned int i;
    unsigned int cNamedArgs;
    int n;
    struct oleparam op;
    memset(&excepinfo, 0, sizeof(EXCEPINFO));

    VariantInit(&result);

    op.dp.rgvarg = NULL;
    op.dp.rgdispidNamedArgs = NULL;
    op.dp.cNamedArgs = 0;
    op.dp.cArgs = 0;

    if (argc < 1)
        mrb_raise(mrb, E_ARGUMENT_ERROR, "wrong number of arguments (0 for 1..)");
    cmd = argv[0];
    paramS = mrb_ary_new_from_values(mrb, argc - 1, argv + 1);
    if(!mrb_string_p(cmd) && !mrb_symbol_p(cmd) && !is_bracket) {
	mrb_raise(mrb, E_TYPE_ERROR, "method is wrong type (expected String or Symbol)");
    }
    if (mrb_symbol_p(cmd)) {
	cmd = mrb_sym2str(mrb, mrb_symbol(cmd));
    }
    OLEData_Get_Struct(mrb, self, pole);
    if(!pole->pDispatch) {
        mrb_raise(mrb, E_RUNTIME_ERROR, "failed to get dispatch interface");
    }
    if (is_bracket) {
        DispID = DISPID_VALUE;
        argc += 1;
        mrb_ary_unshift(mrb, paramS, cmd);
    } else {
        wcmdname = ole_vstr2wc(mrb, cmd);
        hr = pole->pDispatch->lpVtbl->GetIDsOfNames( pole->pDispatch, &IID_NULL,
                &wcmdname, 1, lcid, &DispID);
        SysFreeString(wcmdname);
        if(FAILED(hr)) {
            ole_raise(mrb, hr, E_NOMETHOD_ERROR,
                    "unknown property or method: `%s'",
                    mrb_string_value_ptr(mrb, cmd));
        }
    }

    /* pick up last argument of method */
    param = mrb_ary_entry(paramS, argc-2);

    op.dp.cNamedArgs = 0;

    /* if last arg is hash object */
    if(mrb_hash_p(param)) {
        /*------------------------------------------
          hash object ==> named dispatch parameters
        --------------------------------------------*/
        size_t i;
        mrb_value keys = mrb_hash_keys(mrb, param);
        cNamedArgs = RARRAY_LEN(keys);
        op.dp.cArgs = cNamedArgs + argc - 2;
        op.pNamedArgs = ALLOCA_N(OLECHAR*, cNamedArgs + 1);
        op.dp.rgvarg = ALLOCA_N(VARIANTARG, op.dp.cArgs);

        for (i = 0; i < cNamedArgs; i++) {
            mrb_value key = RARRAY_PTR(keys)[i];
            hash2named_arg(mrb, key, mrb_hash_get(mrb, param, key), &op);
        }

        pDispID = ALLOCA_N(DISPID, cNamedArgs + 1);
        op.pNamedArgs[0] = ole_vstr2wc(mrb, cmd);
        hr = pole->pDispatch->lpVtbl->GetIDsOfNames(pole->pDispatch,
                                                    &IID_NULL,
                                                    op.pNamedArgs,
                                                    op.dp.cNamedArgs + 1,
                                                    lcid, pDispID);
        for(i = 0; i < op.dp.cNamedArgs + 1; i++) {
            SysFreeString(op.pNamedArgs[i]);
            op.pNamedArgs[i] = NULL;
        }
        if(FAILED(hr)) {
            /* clear dispatch parameters */
            for(i = 0; i < op.dp.cArgs; i++ ) {
                VariantClear(&op.dp.rgvarg[i]);
            }
            ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR,
                      "failed to get named argument info: `%s'",
                      mrb_string_value_ptr(mrb, cmd));
        }
        op.dp.rgdispidNamedArgs = &(pDispID[1]);
    }
    else {
        cNamedArgs = 0;
        op.dp.cArgs = argc - 1;
        op.pNamedArgs = ALLOCA_N(OLECHAR*, cNamedArgs + 1);
        if (op.dp.cArgs > 0) {
            op.dp.rgvarg  = ALLOCA_N(VARIANTARG, op.dp.cArgs);
        }
    }
    /*--------------------------------------
      non hash args ==> dispatch parameters
     ----------------------------------------*/
    if(op.dp.cArgs > cNamedArgs) {
        realargs = ALLOCA_N(VARIANTARG, op.dp.cArgs-cNamedArgs+1);
        for(i = cNamedArgs; i < op.dp.cArgs; i++) {
            n = op.dp.cArgs - i + cNamedArgs - 1;
            VariantInit(&realargs[n]);
            VariantInit(&op.dp.rgvarg[n]);
            param = mrb_ary_entry(paramS, i-cNamedArgs);
            if (mrb_obj_is_kind_of(mrb, param, C_WIN32OLE_VARIANT)) {
                ole_variant2variant(mrb, param, &op.dp.rgvarg[n]);
            } else {
                ole_val2variant(mrb, param, &realargs[n]);
                V_VT(&op.dp.rgvarg[n]) = VT_VARIANT | VT_BYREF;
                V_VARIANTREF(&op.dp.rgvarg[n]) = &realargs[n];
            }
        }
    }
    /* apparent you need to call propput, you need this */
    if (wFlags & DISPATCH_PROPERTYPUT) {
        if (op.dp.cArgs == 0)
            ole_raise(mrb, ResultFromScode(E_INVALIDARG), E_WIN32OLE_RUNTIME_ERROR, "argument error");

        op.dp.cNamedArgs = 1;
        op.dp.rgdispidNamedArgs = ALLOCA_N( DISPID, 1 );
        op.dp.rgdispidNamedArgs[0] = DISPID_PROPERTYPUT;
    }
    hr = pole->pDispatch->lpVtbl->Invoke(pole->pDispatch, DispID,
                                         &IID_NULL, lcid, wFlags, &op.dp,
                                         &result, &excepinfo, &argErr);

    if (FAILED(hr)) {
        /* retry to call args by value */
        if(op.dp.cArgs >= cNamedArgs) {
            for(i = cNamedArgs; i < op.dp.cArgs; i++) {
                n = op.dp.cArgs - i + cNamedArgs - 1;
                param = mrb_ary_entry(paramS, i-cNamedArgs);
                ole_val2variant(mrb, param, &op.dp.rgvarg[n]);
            }
            if (hr == DISP_E_EXCEPTION) {
                ole_freeexceptinfo(&excepinfo);
            }
            memset(&excepinfo, 0, sizeof(EXCEPINFO));
            VariantInit(&result);
            hr = pole->pDispatch->lpVtbl->Invoke(pole->pDispatch, DispID,
                                                 &IID_NULL, lcid, wFlags,
                                                 &op.dp, &result,
                                                 &excepinfo, &argErr);

            /* mega kludge. if a method in WORD is called and we ask
             * for a result when one is not returned then
             * hResult == DISP_E_EXCEPTION. this only happens on
             * functions whose DISPID > 0x8000 */
            if ((hr == DISP_E_EXCEPTION || hr == DISP_E_MEMBERNOTFOUND) && DispID > 0x8000) {
                if (hr == DISP_E_EXCEPTION) {
                    ole_freeexceptinfo(&excepinfo);
                }
                memset(&excepinfo, 0, sizeof(EXCEPINFO));
                hr = pole->pDispatch->lpVtbl->Invoke(pole->pDispatch, DispID,
                        &IID_NULL, lcid, wFlags,
                        &op.dp, NULL,
                        &excepinfo, &argErr);

            }
            for(i = cNamedArgs; i < op.dp.cArgs; i++) {
                n = op.dp.cArgs - i + cNamedArgs - 1;
                if (V_VT(&op.dp.rgvarg[n]) != VT_RECORD) {
                    VariantClear(&op.dp.rgvarg[n]);
                }
            }
        }

        if (FAILED(hr)) {
            /* retry after converting nil to VT_EMPTY */
            if (op.dp.cArgs > cNamedArgs) {
                for(i = cNamedArgs; i < op.dp.cArgs; i++) {
                    n = op.dp.cArgs - i + cNamedArgs - 1;
                    param = mrb_ary_entry(paramS, i-cNamedArgs);
                    ole_val2variant2(mrb, param, &op.dp.rgvarg[n]);
                }
                if (hr == DISP_E_EXCEPTION) {
                    ole_freeexceptinfo(&excepinfo);
                }
                memset(&excepinfo, 0, sizeof(EXCEPINFO));
                VariantInit(&result);
                hr = pole->pDispatch->lpVtbl->Invoke(pole->pDispatch, DispID,
                        &IID_NULL, lcid, wFlags,
                        &op.dp, &result,
                        &excepinfo, &argErr);
                for(i = cNamedArgs; i < op.dp.cArgs; i++) {
                    n = op.dp.cArgs - i + cNamedArgs - 1;
                    if (V_VT(&op.dp.rgvarg[n]) != VT_RECORD) {
                        VariantClear(&op.dp.rgvarg[n]);
                    }
                }
            }
        }

    }
    /* clear dispatch parameter */
    if(op.dp.cArgs > cNamedArgs) {
        for(i = cNamedArgs; i < op.dp.cArgs; i++) {
            n = op.dp.cArgs - i + cNamedArgs - 1;
            param = mrb_ary_entry(paramS, i-cNamedArgs);
            if (mrb_obj_is_kind_of(mrb, param, C_WIN32OLE_VARIANT)) {
                ole_val2variant(mrb, param, &realargs[n]);
            } else if ( mrb_obj_is_kind_of(mrb, param, C_WIN32OLE_RECORD) &&
                        V_VT(&realargs[n]) == VT_RECORD ) {
                olerecord_set_ivar(mrb, param, V_RECORDINFO(&realargs[n]), V_RECORD(&realargs[n]));
            }
        }
        set_argv(mrb, realargs, cNamedArgs, op.dp.cArgs);
    }
    else {
        for(i = 0; i < op.dp.cArgs; i++) {
            VariantClear(&op.dp.rgvarg[i]);
        }
    }

    if (FAILED(hr)) {
        v = ole_excepinfo2msg(mrb, &excepinfo);
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "(in OLE method `%s': )%s",
                  mrb_string_value_ptr(mrb, cmd),
                  mrb_string_value_ptr(mrb, v));
    }
    obj = ole_variant2val(mrb, &result);
    VariantClear(&result);
    return obj;
}

/*
 *  call-seq:
 *     WIN32OLE#invoke(method, [arg1,...])  => return value of method.
 *
 *  Runs OLE method.
 *  The first argument specifies the method name of OLE Automation object.
 *  The others specify argument of the <i>method</i>.
 *  If you can not execute <i>method</i> directly, then use this method instead.
 *
 *    excel = WIN32OLE.new('Excel.Application')
 *    excel.invoke('Quit')  # => same as excel.Quit
 *
 */
static mrb_value
fole_invoke(mrb_state *mrb, mrb_value self)
{
    int argc;
    mrb_value *argv;
    mrb_get_args(mrb, "*", &argv, &argc);
    return ole_invoke(mrb, argc, argv, self, DISPATCH_METHOD|DISPATCH_PROPERTYGET, FALSE);
}

static mrb_value
ole_invoke2(mrb_state *mrb, mrb_value self, mrb_value dispid, mrb_value args, mrb_value types, USHORT dispkind)
{
    HRESULT hr;
    struct oledata *pole;
    unsigned int argErr = 0;
    EXCEPINFO excepinfo;
    VARIANT result;
    DISPPARAMS dispParams;
    VARIANTARG* realargs = NULL;
    int i, j;
    mrb_value obj = mrb_nil_value();
    mrb_value tp, param;
    mrb_value v;
    VARTYPE vt;

    mrb_check_type(mrb, args, MRB_TT_ARRAY);
    mrb_check_type(mrb, types, MRB_TT_ARRAY);

    memset(&excepinfo, 0, sizeof(EXCEPINFO));
    memset(&dispParams, 0, sizeof(DISPPARAMS));
    VariantInit(&result);
    OLEData_Get_Struct(mrb, self, pole);

    dispParams.cArgs = RARRAY_LEN(args);
    dispParams.rgvarg = ALLOCA_N(VARIANTARG, dispParams.cArgs);
    realargs = ALLOCA_N(VARIANTARG, dispParams.cArgs);
    for (i = 0, j = dispParams.cArgs - 1; i < (int)dispParams.cArgs; i++, j--)
    {
        VariantInit(&realargs[i]);
        VariantInit(&dispParams.rgvarg[i]);
        tp = mrb_ary_entry(types, j);
        vt = (VARTYPE)FIX2INT(tp);
        V_VT(&dispParams.rgvarg[i]) = vt;
        param = mrb_ary_entry(args, j);
        if (mrb_nil_p(param))
        {

            V_VT(&dispParams.rgvarg[i]) = V_VT(&realargs[i]) = VT_ERROR;
            V_ERROR(&dispParams.rgvarg[i]) = V_ERROR(&realargs[i]) = DISP_E_PARAMNOTFOUND;
        }
        else
        {
            if (vt & VT_ARRAY)
            {
                int ent;
                LPBYTE pb;
                short* ps;
                LPLONG pl;
                VARIANT* pv;
                CY *py;
                VARTYPE v;
                SAFEARRAYBOUND rgsabound[1];
                mrb_check_type(mrb, param, MRB_TT_ARRAY);
                rgsabound[0].lLbound = 0;
                rgsabound[0].cElements = RARRAY_LEN(param);
                v = vt & ~(VT_ARRAY | VT_BYREF);
                V_ARRAY(&realargs[i]) = SafeArrayCreate(v, 1, rgsabound);
                V_VT(&realargs[i]) = VT_ARRAY | v;
                SafeArrayLock(V_ARRAY(&realargs[i]));
                pb = V_ARRAY(&realargs[i])->pvData;
                ps = V_ARRAY(&realargs[i])->pvData;
                pl = V_ARRAY(&realargs[i])->pvData;
                py = V_ARRAY(&realargs[i])->pvData;
                pv = V_ARRAY(&realargs[i])->pvData;
                for (ent = 0; ent < (int)rgsabound[0].cElements; ent++)
                {
                    VARIANT velem;
                    mrb_value elem = mrb_ary_entry(param, ent);
                    ole_val2variant(mrb, elem, &velem);
                    if (v != VT_VARIANT)
                    {
                        VariantChangeTypeEx(&velem, &velem,
                            cWIN32OLE_lcid, 0, v);
                    }
                    switch (v)
                    {
                    /* 128 bits */
                    case VT_VARIANT:
                        *pv++ = velem;
                        break;
                    /* 64 bits */
                    case VT_R8:
                    case VT_CY:
                    case VT_DATE:
                        *py++ = V_CY(&velem);
                        break;
                    /* 16 bits */
                    case VT_BOOL:
                    case VT_I2:
                    case VT_UI2:
                        *ps++ = V_I2(&velem);
                        break;
                    /* 8 bites */
                    case VT_UI1:
                    case VT_I1:
                        *pb++ = V_UI1(&velem);
                        break;
                    /* 32 bits */
                    default:
                        *pl++ = V_I4(&velem);
                        break;
                    }
                }
                SafeArrayUnlock(V_ARRAY(&realargs[i]));
            }
            else
            {
                ole_val2variant(mrb, param, &realargs[i]);
                if ((vt & (~VT_BYREF)) != VT_VARIANT)
                {
                    hr = VariantChangeTypeEx(&realargs[i], &realargs[i],
                                             cWIN32OLE_lcid, 0,
                                             (VARTYPE)(vt & (~VT_BYREF)));
                    if (hr != S_OK)
                    {
                        mrb_raise(mrb, E_TYPE_ERROR, "not valid value");
                    }
                }
            }
            if ((vt & VT_BYREF) || vt == VT_VARIANT)
            {
                if (vt == VT_VARIANT)
                    V_VT(&dispParams.rgvarg[i]) = VT_VARIANT | VT_BYREF;
                switch (vt & (~VT_BYREF))
                {
                /* 128 bits */
                case VT_VARIANT:
                    V_VARIANTREF(&dispParams.rgvarg[i]) = &realargs[i];
                    break;
                /* 64 bits */
                case VT_R8:
                case VT_CY:
                case VT_DATE:
                    V_CYREF(&dispParams.rgvarg[i]) = &V_CY(&realargs[i]);
                    break;
                /* 16 bits */
                case VT_BOOL:
                case VT_I2:
                case VT_UI2:
                    V_I2REF(&dispParams.rgvarg[i]) = &V_I2(&realargs[i]);
                    break;
                /* 8 bites */
                case VT_UI1:
                case VT_I1:
                    V_UI1REF(&dispParams.rgvarg[i]) = &V_UI1(&realargs[i]);
                    break;
                /* 32 bits */
                default:
                    V_I4REF(&dispParams.rgvarg[i]) = &V_I4(&realargs[i]);
                    break;
                }
            }
            else
            {
                /* copy 64 bits of data */
                V_CY(&dispParams.rgvarg[i]) = V_CY(&realargs[i]);
            }
        }
    }

    if (dispkind & DISPATCH_PROPERTYPUT) {
        dispParams.cNamedArgs = 1;
        dispParams.rgdispidNamedArgs = ALLOCA_N( DISPID, 1 );
        dispParams.rgdispidNamedArgs[0] = DISPID_PROPERTYPUT;
    }

    hr = pole->pDispatch->lpVtbl->Invoke(pole->pDispatch, NUM2INT(dispid),
                                         &IID_NULL, cWIN32OLE_lcid,
                                         dispkind,
                                         &dispParams, &result,
                                         &excepinfo, &argErr);

    if (FAILED(hr)) {
        v = ole_excepinfo2msg(mrb, &excepinfo);
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "(in OLE method `<dispatch id:%d>': )%s",
                  NUM2INT(dispid),
                  mrb_string_value_ptr(mrb, v));
    }

    /* clear dispatch parameter */
    if(dispParams.cArgs > 0) {
        set_argv(mrb, realargs, 0, dispParams.cArgs);
    }

    obj = ole_variant2val(mrb, &result);
    VariantClear(&result);
    return obj;
}

/*
 *   call-seq:
 *      WIN32OLE#_invoke(dispid, args, types)
 *
 *   Runs the early binding method.
 *   The 1st argument specifies dispatch ID,
 *   the 2nd argument specifies the array of arguments,
 *   the 3rd argument specifies the array of the type of arguments.
 *
 *      excel = WIN32OLE.new('Excel.Application')
 *      excel._invoke(302, [], []) #  same effect as excel.Quit
 */
static mrb_value
fole_invoke2(mrb_state *mrb, mrb_value self)
{
    mrb_value dispid, args, types;
    mrb_get_args(mrb, "oAA", &dispid, &args, &types);
    return ole_invoke2(mrb, self, dispid, args, types, DISPATCH_METHOD);
}

/*
 *  call-seq:
 *     WIN32OLE#_getproperty(dispid, args, types)
 *
 *  Runs the early binding method to get property.
 *  The 1st argument specifies dispatch ID,
 *  the 2nd argument specifies the array of arguments,
 *  the 3rd argument specifies the array of the type of arguments.
 *
 *     excel = WIN32OLE.new('Excel.Application')
 *     puts excel._getproperty(558, [], []) # same effect as puts excel.visible
 */
static mrb_value
fole_getproperty2(mrb_state *mrb, mrb_value self)
{
    mrb_value dispid, args, types;
    mrb_get_args(mrb, "iAA", &dispid, &args, &types);
    return ole_invoke2(mrb, self, dispid, args, types, DISPATCH_PROPERTYGET);
}

/*
 *   call-seq:
 *      WIN32OLE#_setproperty(dispid, args, types)
 *
 *   Runs the early binding method to set property.
 *   The 1st argument specifies dispatch ID,
 *   the 2nd argument specifies the array of arguments,
 *   the 3rd argument specifies the array of the type of arguments.
 *
 *      excel = WIN32OLE.new('Excel.Application')
 *      excel._setproperty(558, [true], [WIN32OLE::VARIANT::VT_BOOL]) # same effect as excel.visible = true
 */
static mrb_value
fole_setproperty2(mrb_state *mrb, mrb_value self)
{
    mrb_value dispid, args, types;
    mrb_get_args(mrb, "iAA", &dispid, &args, &types);
    return ole_invoke2(mrb, self, dispid, args, types, DISPATCH_PROPERTYPUT);
}

/*
 *  call-seq:
 *     WIN32OLE[a1, a2, ...]=val
 *
 *  Sets the value to WIN32OLE object specified by a1, a2, ...
 *
 *     dict = WIN32OLE.new('Scripting.Dictionary')
 *     dict.add('ruby', 'RUBY')
 *     dict['ruby'] = 'Ruby'
 *     puts dict['ruby'] # => 'Ruby'
 *
 *  Remark: You can not use this method to set the property value.
 *
 *     excel = WIN32OLE.new('Excel.Application')
 *     # excel['Visible'] = true # This is error !!!
 *     excel.Visible = true # You should to use this style to set the property.
 *
 */
static mrb_value
fole_setproperty_with_bracket(mrb_state *mrb, mrb_value self)
{
    int argc;
    mrb_value *argv;
    mrb_get_args(mrb, "*", &argv, &argc);
    return ole_invoke(mrb, argc, argv, self, DISPATCH_PROPERTYPUT, TRUE);
}

/*
 *  call-seq:
 *     WIN32OLE.setproperty('property', [arg1, arg2,...] val)
 *
 *  Sets property of OLE object.
 *  When you want to set property with argument, you can use this method.
 *
 *     excel = WIN32OLE.new('Excel.Application')
 *     excel.Visible = true
 *     book = excel.workbooks.add
 *     sheet = book.worksheets(1)
 *     sheet.setproperty('Cells', 1, 2, 10) # => The B1 cell value is 10.
 */
static mrb_value
fole_setproperty(mrb_state *mrb, mrb_value self)
{
    int argc;
    mrb_value *argv;
    mrb_get_args(mrb, "*", &argv, &argc);
    return ole_invoke(mrb, argc, argv, self, DISPATCH_PROPERTYPUT, FALSE);
}

/*
 *  call-seq:
 *     WIN32OLE[a1,a2,...]
 *
 *  Returns the value of Collection specified by a1, a2,....
 *
 *     dict = WIN32OLE.new('Scripting.Dictionary')
 *     dict.add('ruby', 'Ruby')
 *     puts dict['ruby'] # => 'Ruby' (same as `puts dict.item('ruby')')
 *
 *  Remark: You can not use this method to get the property.
 *     excel = WIN32OLE.new('Excel.Application')
 *     # puts excel['Visible']  This is error !!!
 *     puts excel.Visible # You should to use this style to get the property.
 *
 */
static mrb_value
fole_getproperty_with_bracket(mrb_state *mrb, mrb_value self)
{
    int argc;
    mrb_value *argv;
    mrb_get_args(mrb, "*", &argv, &argc);
    return ole_invoke(mrb, argc, argv, self, DISPATCH_PROPERTYGET, TRUE);
}

static mrb_value
ole_propertyput(mrb_state *mrb, mrb_value self, mrb_value property, mrb_value value)
{
    struct oledata *pole;
    unsigned argErr;
    unsigned int index;
    HRESULT hr;
    EXCEPINFO excepinfo;
    DISPID dispID = DISPID_VALUE;
    DISPID dispIDParam = DISPID_PROPERTYPUT;
    USHORT wFlags = DISPATCH_PROPERTYPUT|DISPATCH_PROPERTYPUTREF;
    DISPPARAMS dispParams;
    VARIANTARG propertyValue[2];
    OLECHAR* pBuf[1];
    mrb_value v;
    LCID    lcid = cWIN32OLE_lcid;
    dispParams.rgdispidNamedArgs = &dispIDParam;
    dispParams.rgvarg = propertyValue;
    dispParams.cNamedArgs = 1;
    dispParams.cArgs = 1;

    VariantInit(&propertyValue[0]);
    VariantInit(&propertyValue[1]);
    memset(&excepinfo, 0, sizeof(excepinfo));

    OLEData_Get_Struct(mrb, self, pole);

    /* get ID from property name */
    pBuf[0]  = ole_vstr2wc(mrb, property);
    hr = pole->pDispatch->lpVtbl->GetIDsOfNames(pole->pDispatch, &IID_NULL,
                                                pBuf, 1, lcid, &dispID);
    SysFreeString(pBuf[0]);
    pBuf[0] = NULL;

    if(FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR,
                  "unknown property or method: `%s'",
                  mrb_string_value_ptr(mrb, property));
    }
    /* set property value */
    ole_val2variant(mrb, value, &propertyValue[0]);
    hr = pole->pDispatch->lpVtbl->Invoke(pole->pDispatch, dispID, &IID_NULL,
                                         lcid, wFlags, &dispParams,
                                         NULL, &excepinfo, &argErr);

    for(index = 0; index < dispParams.cArgs; ++index) {
        VariantClear(&propertyValue[index]);
    }
    if (FAILED(hr)) {
        v = ole_excepinfo2msg(mrb, &excepinfo);
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "(in setting property `%s': )%s",
                  mrb_string_value_ptr(mrb, property),
                  mrb_string_value_ptr(mrb, v));
    }
    return mrb_nil_value();
}

/*
 *  call-seq:
 *     WIN32OLE#ole_free
 *
 *  invokes Release method of Dispatch interface of WIN32OLE object.
 *  Usually, you do not need to call this method because Release method
 *  called automatically when WIN32OLE object garbaged.
 *
 */
static mrb_value
fole_free(mrb_state *mrb, mrb_value self)
{
    struct oledata *pole;
    OLEData_Get_Struct(mrb, self, pole);
    OLE_FREE(pole->pDispatch);
    pole->pDispatch = NULL;
    return mrb_nil_value();
}

static mrb_value
ole_each_sub(mrb_state *mrb, mrb_value self)
{
    IEnumVARIANT *pEnum;
    mrb_value pEnumV, block;
    VARIANT variant;
    mrb_value obj = mrb_nil_value();
    mrb_get_args(mrb, "oo", &pEnumV, &block);
    pEnum = (IEnumVARIANT *)mrb_cptr(pEnumV);
    VariantInit(&variant);
    while(pEnum->lpVtbl->Next(pEnum, 1, &variant, NULL) == S_OK) {
        obj = ole_variant2val(mrb, &variant);
        VariantClear(&variant);
        VariantInit(&variant);
        mrb_yield(mrb, block, obj);
    }
    return mrb_nil_value();
}

static mrb_value
ole_ienum_free(mrb_state *mrb, IEnumVARIANT *pEnum)
{
    OLE_RELEASE(pEnum);
    return mrb_nil_value();
}

/*
 *  call-seq:
 *     WIN32OLE#each {|i|...}
 *
 *  Iterates over each item of OLE collection which has IEnumVARIANT interface.
 *
 *     excel = WIN32OLE.new('Excel.Application')
 *     book = excel.workbooks.add
 *     sheets = book.worksheets(1)
 *     cells = sheets.cells("A1:A5")
 *     cells.each do |cell|
 *       cell.value = 10
 *     end
 */
static mrb_value
fole_each(mrb_state *mrb, mrb_value self)
{
    LCID    lcid = cWIN32OLE_lcid;

    struct oledata *pole;

    unsigned int argErr;
    EXCEPINFO excepinfo;
    DISPPARAMS dispParams;
    VARIANT result;
    HRESULT hr;
    IEnumVARIANT *pEnum = NULL;
    void *p;
    mrb_value block;

    mrb_get_args(mrb, "&", &block);

/* FIXME:
    RETURN_ENUMERATOR(self, 0, 0);
*/
    if (mrb_nil_p(block))
        return mrb_funcall(mrb, self, "to_enum", 0);
/* */

    VariantInit(&result);
    dispParams.rgvarg = NULL;
    dispParams.rgdispidNamedArgs = NULL;
    dispParams.cNamedArgs = 0;
    dispParams.cArgs = 0;
    memset(&excepinfo, 0, sizeof(excepinfo));

    OLEData_Get_Struct(mrb, self, pole);
    hr = pole->pDispatch->lpVtbl->Invoke(pole->pDispatch, DISPID_NEWENUM,
                                         &IID_NULL, lcid,
                                         DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                                         &dispParams, &result,
                                         &excepinfo, &argErr);

    if (FAILED(hr)) {
        VariantClear(&result);
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to get IEnum Interface");
    }

    if (V_VT(&result) == VT_UNKNOWN) {
        hr = V_UNKNOWN(&result)->lpVtbl->QueryInterface(V_UNKNOWN(&result),
                                                        &IID_IEnumVARIANT,
                                                        &p);
        pEnum = p;
    } else if (V_VT(&result) == VT_DISPATCH) {
        hr = V_DISPATCH(&result)->lpVtbl->QueryInterface(V_DISPATCH(&result),
                                                         &IID_IEnumVARIANT,
                                                         &p);
        pEnum = p;
    }
    if (FAILED(hr) || !pEnum) {
        VariantClear(&result);
        ole_raise(mrb, hr, E_RUNTIME_ERROR, "failed to get IEnum Interface");
    }

    VariantClear(&result);
    /* FIXME: Is there better way to do things like rb_ensure() in mruby?
    rb_ensure(ole_each_sub, (mrb_value)pEnum, ole_ienum_free, (mrb_value)pEnum);
    */
    mrb_define_method(mrb, mrb_obj_class(mrb, self), "_tmp_method_ole_each_", ole_each_sub, MRB_ARGS_REQ(2));
    mrb_funcall(mrb, self, "_tmp_method_ole_each_", 2, mrb_cptr_value(mrb, pEnum), block);
    mrb_undef_method(mrb, mrb_obj_class(mrb, self), "_tmp_method_ole_each_");
    ole_ienum_free(mrb, pEnum);
    /* FIXME: */
    return mrb_nil_value();
}

/*
 *  call-seq:
 *     WIN32OLE#method_missing(id [,arg1, arg2, ...])
 *
 *  Calls WIN32OLE#invoke method.
 */
static mrb_value
fole_missing(mrb_state *mrb, mrb_value self)
{
    int argc;
    mrb_value *argv;
    mrb_sym id;
    const char* mname;
    size_t n;
    mrb_get_args(mrb, "*", &argv, &argc);
    if (argc < 1)
        mrb_raisef(mrb, E_ARGUMENT_ERROR, "wrong number of arguments (%S for 1..)", mrb_fixnum_value(argc));
    id = mrb_symbol(argv[0]);
    mname = mrb_sym2name(mrb, id);
    if(!mname) {
        mrb_raise(mrb, E_RUNTIME_ERROR, "fail: unknown method or property");
    }
    n = strlen(mname);
#if SIZEOF_SIZE_T > SIZEOF_LONG
    if (n >= LONG_MAX) {
	mrb_raise(mrb, E_RUNTIME_ERROR, "too long method or property name");
    }
#endif
    if(mname[n-1] == '=') {
        if (argc != 2)
            mrb_raisef(mrb, E_ARGUMENT_ERROR, "wrong number of arguments (%S for 2)", mrb_fixnum_value(argc));
        argv[0] = mrb_enc_str_new(mrb, mname, (long)(n-1), cWIN32OLE_enc);

        return ole_propertyput(mrb, self, argv[0], argv[1]);
    }
    else {
        argv[0] = mrb_enc_str_new(mrb, mname, (long)n, cWIN32OLE_enc);
        return ole_invoke(mrb, argc, argv, self, DISPATCH_METHOD|DISPATCH_PROPERTYGET, FALSE);
    }
}

static HRESULT
typeinfo_from_ole(mrb_state *mrb, struct oledata *pole, ITypeInfo **ppti)
{
    ITypeInfo *pTypeInfo;
    ITypeLib *pTypeLib;
    BSTR bstr;
    mrb_value type;
    UINT i;
    UINT count;
    LCID    lcid = cWIN32OLE_lcid;
    HRESULT hr = pole->pDispatch->lpVtbl->GetTypeInfo(pole->pDispatch,
                                                      0, lcid, &pTypeInfo);
    if(FAILED(hr)) {
        ole_raise(mrb, hr, E_RUNTIME_ERROR, "failed to GetTypeInfo");
    }
    hr = pTypeInfo->lpVtbl->GetDocumentation(pTypeInfo,
                                             -1,
                                             &bstr,
                                             NULL, NULL, NULL);
    type = WC2VSTR(mrb, bstr);
    hr = pTypeInfo->lpVtbl->GetContainingTypeLib(pTypeInfo, &pTypeLib, &i);
    OLE_RELEASE(pTypeInfo);
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_RUNTIME_ERROR, "failed to GetContainingTypeLib");
    }
    count = pTypeLib->lpVtbl->GetTypeInfoCount(pTypeLib);
    for (i = 0; i < count; i++) {
        hr = pTypeLib->lpVtbl->GetDocumentation(pTypeLib, i,
                                                &bstr, NULL, NULL, NULL);
        if (SUCCEEDED(hr) && mrb_str_cmp(mrb, WC2VSTR(mrb, bstr), type) == 0) {
            hr = pTypeLib->lpVtbl->GetTypeInfo(pTypeLib, i, &pTypeInfo);
            if (SUCCEEDED(hr)) {
                *ppti = pTypeInfo;
                break;
            }
        }
    }
    OLE_RELEASE(pTypeLib);
    return hr;
}

static mrb_value
ole_methods(mrb_state *mrb, mrb_value self, int mask)
{
    ITypeInfo *pTypeInfo;
    HRESULT hr;
    mrb_value methods;
    struct oledata *pole;

    OLEData_Get_Struct(mrb, self, pole);
    methods = mrb_ary_new(mrb);

    hr = typeinfo_from_ole(mrb, pole, &pTypeInfo);
    if(FAILED(hr))
        return methods;
    mrb_ary_concat(mrb, methods, ole_methods_from_typeinfo(mrb, pTypeInfo, mask));
    OLE_RELEASE(pTypeInfo);
    return methods;
}

/*
 *  call-seq:
 *     WIN32OLE#ole_methods
 *
 *  Returns the array of WIN32OLE_METHOD object.
 *  The element is OLE method of WIN32OLE object.
 *
 *     excel = WIN32OLE.new('Excel.Application')
 *     methods = excel.ole_methods
 *
 */
static mrb_value
fole_methods(mrb_state *mrb, mrb_value self)
{
    return ole_methods(mrb,  self, INVOKE_FUNC | INVOKE_PROPERTYGET | INVOKE_PROPERTYPUT | INVOKE_PROPERTYPUTREF);
}

/*
 *  call-seq:
 *     WIN32OLE#ole_get_methods
 *
 *  Returns the array of WIN32OLE_METHOD object .
 *  The element of the array is property (gettable) of WIN32OLE object.
 *
 *     excel = WIN32OLE.new('Excel.Application')
 *     properties = excel.ole_get_methods
 */
static mrb_value
fole_get_methods(mrb_state *mrb, mrb_value self)
{
    return ole_methods(mrb,  self, INVOKE_PROPERTYGET);
}

/*
 *  call-seq:
 *     WIN32OLE#ole_put_methods
 *
 *  Returns the array of WIN32OLE_METHOD object .
 *  The element of the array is property (settable) of WIN32OLE object.
 *
 *     excel = WIN32OLE.new('Excel.Application')
 *     properties = excel.ole_put_methods
 */
static mrb_value
fole_put_methods(mrb_state *mrb, mrb_value self)
{
    return ole_methods(mrb,  self, INVOKE_PROPERTYPUT|INVOKE_PROPERTYPUTREF);
}

/*
 *  call-seq:
 *     WIN32OLE#ole_func_methods
 *
 *  Returns the array of WIN32OLE_METHOD object .
 *  The element of the array is property (settable) of WIN32OLE object.
 *
 *     excel = WIN32OLE.new('Excel.Application')
 *     properties = excel.ole_func_methods
 *
 */
static mrb_value
fole_func_methods(mrb_state *mrb, mrb_value self)
{
    return ole_methods( mrb, self, INVOKE_FUNC);
}

/*
 *   call-seq:
 *      WIN32OLE#ole_type
 *
 *   Returns WIN32OLE_TYPE object.
 *
 *      excel = WIN32OLE.new('Excel.Application')
 *      tobj = excel.ole_type
 */
static mrb_value
fole_type(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo;
    HRESULT hr;
    struct oledata *pole;
    LCID  lcid = cWIN32OLE_lcid;
    mrb_value type = mrb_nil_value();

    OLEData_Get_Struct(mrb, self, pole);

    hr = pole->pDispatch->lpVtbl->GetTypeInfo( pole->pDispatch, 0, lcid, &pTypeInfo );
    if(FAILED(hr)) {
        ole_raise(mrb, hr, E_RUNTIME_ERROR, "failed to GetTypeInfo");
    }
    type = ole_type_from_itypeinfo(mrb, pTypeInfo);
    OLE_RELEASE(pTypeInfo);
    if (mrb_nil_p(type)) {
        mrb_raise(mrb, E_RUNTIME_ERROR, "failed to create WIN32OLE_TYPE obj from ITypeInfo");
    }
    return type;
}

/*
 *  call-seq:
 *     WIN32OLE#ole_typelib -> The WIN32OLE_TYPELIB object
 *
 *  Returns the WIN32OLE_TYPELIB object. The object represents the
 *  type library which contains the WIN32OLE object.
 *
 *     excel = WIN32OLE.new('Excel.Application')
 *     tlib = excel.ole_typelib
 *     puts tlib.name  # -> 'Microsoft Excel 9.0 Object Library'
 */
static mrb_value
fole_typelib(mrb_state *mrb, mrb_value self)
{
    struct oledata *pole;
    HRESULT hr;
    ITypeInfo *pTypeInfo;
    LCID  lcid = cWIN32OLE_lcid;
    mrb_value vtlib = mrb_nil_value();

    OLEData_Get_Struct(mrb, self, pole);
    hr = pole->pDispatch->lpVtbl->GetTypeInfo(pole->pDispatch,
                                              0, lcid, &pTypeInfo);
    if(FAILED(hr)) {
        ole_raise(mrb, hr, E_RUNTIME_ERROR, "failed to GetTypeInfo");
    }
    vtlib = ole_typelib_from_itypeinfo(mrb, pTypeInfo);
    OLE_RELEASE(pTypeInfo);
    if (mrb_nil_p(vtlib)) {
        mrb_raise(mrb, E_RUNTIME_ERROR, "failed to get type library info.");
    }
    return vtlib;
}

/*
 *  call-seq:
 *     WIN32OLE#ole_query_interface(iid) -> WIN32OLE object
 *
 *  Returns WIN32OLE object for a specific dispatch or dual
 *  interface specified by iid.
 *
 *      ie = WIN32OLE.new('InternetExplorer.Application')
 *      ie_web_app = ie.ole_query_interface('{0002DF05-0000-0000-C000-000000000046}') # => WIN32OLE object for dispinterface IWebBrowserApp
 */
static mrb_value
fole_query_interface(mrb_state *mrb, mrb_value self)
{
    mrb_value str_iid;
    HRESULT hr;
    BSTR pBuf;
    IID iid;
    struct oledata *pole;
    IDispatch *pDispatch;
    void *p;

    mrb_get_args(mrb, "S", &str_iid);

    pBuf  = ole_vstr2wc(mrb, str_iid);
    hr = CLSIDFromString(pBuf, &iid);
    SysFreeString(pBuf);
    if(FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR,
                  "invalid iid: `%s'",
                  mrb_string_value_ptr(mrb, str_iid));
    }

    OLEData_Get_Struct(mrb, self, pole);
    if(!pole->pDispatch) {
        mrb_raise(mrb, E_RUNTIME_ERROR, "failed to get dispatch interface");
    }

    hr = pole->pDispatch->lpVtbl->QueryInterface(pole->pDispatch, &iid,
                                                 &p);
    if(FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR,
                  "failed to get interface `%s'",
                  mrb_string_value_ptr(mrb, str_iid));
    }

    pDispatch = p;
    return create_win32ole_object(mrb, C_WIN32OLE, pDispatch, 0, 0);
}

/*
 *  call-seq:
 *     WIN32OLE#ole_respond_to?(method) -> true or false
 *
 *  Returns true when OLE object has OLE method, otherwise returns false.
 *
 *      ie = WIN32OLE.new('InternetExplorer.Application')
 *      ie.ole_respond_to?("gohome") => true
 */
static mrb_value
fole_respond_to(mrb_state *mrb, mrb_value self)
{
    mrb_value method;
    struct oledata *pole;
    BSTR wcmdname;
    DISPID DispID;
    HRESULT hr;
    mrb_get_args(mrb, "o", &method);
    if(!mrb_string_p(method) && !mrb_symbol_p(method)) {
        mrb_raise(mrb, E_TYPE_ERROR, "wrong argument type (expected String or Symbol)");
    }
    if (mrb_symbol_p(method)) {
        method = mrb_sym2str(mrb, mrb_symbol(method));
    }
    OLEData_Get_Struct(mrb, self, pole);
    wcmdname = ole_vstr2wc(mrb, method);
    hr = pole->pDispatch->lpVtbl->GetIDsOfNames( pole->pDispatch, &IID_NULL,
	    &wcmdname, 1, cWIN32OLE_lcid, &DispID);
    SysFreeString(wcmdname);
    return SUCCEEDED(hr) ? mrb_true_value() : mrb_false_value();
}

HRESULT
ole_docinfo_from_type(mrb_state *mrb, ITypeInfo *pTypeInfo, BSTR *name, BSTR *helpstr, DWORD *helpcontext, BSTR *helpfile)
{
    HRESULT hr;
    ITypeLib *pTypeLib;
    UINT i;

    hr = pTypeInfo->lpVtbl->GetContainingTypeLib(pTypeInfo, &pTypeLib, &i);
    if (FAILED(hr)) {
        return hr;
    }

    hr = pTypeLib->lpVtbl->GetDocumentation(pTypeLib, i,
                                            name, helpstr,
                                            helpcontext, helpfile);
    if (FAILED(hr)) {
        OLE_RELEASE(pTypeLib);
        return hr;
    }
    OLE_RELEASE(pTypeLib);
    return hr;
}

static mrb_value
ole_usertype2val(mrb_state *mrb, ITypeInfo *pTypeInfo, TYPEDESC *pTypeDesc, mrb_value typedetails)
{
    HRESULT hr;
    BSTR bstr;
    ITypeInfo *pRefTypeInfo;
    mrb_value type = mrb_nil_value();

    hr = pTypeInfo->lpVtbl->GetRefTypeInfo(pTypeInfo,
                                           V_UNION1(pTypeDesc, hreftype),
                                           &pRefTypeInfo);
    if(FAILED(hr))
        return mrb_nil_value();
    hr = ole_docinfo_from_type(mrb, pRefTypeInfo, &bstr, NULL, NULL, NULL);
    if(FAILED(hr)) {
        OLE_RELEASE(pRefTypeInfo);
        return mrb_nil_value();
    }
    OLE_RELEASE(pRefTypeInfo);
    type = WC2VSTR(mrb, bstr);
    if(!mrb_nil_p(typedetails))
        mrb_ary_push(mrb, typedetails, type);
    return type;
}

static mrb_value
ole_ptrtype2val(mrb_state *mrb, ITypeInfo *pTypeInfo, TYPEDESC *pTypeDesc, mrb_value typedetails)
{
    TYPEDESC *p = pTypeDesc;
    mrb_value type = mrb_str_new_lit(mrb, "");

    if (p->vt == VT_PTR || p->vt == VT_SAFEARRAY) {
        p = V_UNION1(p, lptdesc);
        type = ole_typedesc2val(mrb, pTypeInfo, p, typedetails);
    }
    return type;
}

mrb_value
ole_typedesc2val(mrb_state *mrb, ITypeInfo *pTypeInfo, TYPEDESC *pTypeDesc, mrb_value typedetails)
{
    mrb_value str;
    mrb_value typestr = mrb_nil_value();
    switch(pTypeDesc->vt) {
    case VT_I2:
        typestr = mrb_str_new_lit(mrb, "I2");
        break;
    case VT_I4:
        typestr = mrb_str_new_lit(mrb, "I4");
        break;
    case VT_R4:
        typestr = mrb_str_new_lit(mrb, "R4");
        break;
    case VT_R8:
        typestr = mrb_str_new_lit(mrb, "R8");
        break;
    case VT_CY:
        typestr = mrb_str_new_lit(mrb, "CY");
        break;
    case VT_DATE:
        typestr = mrb_str_new_lit(mrb, "DATE");
        break;
    case VT_BSTR:
        typestr = mrb_str_new_lit(mrb, "BSTR");
        break;
    case VT_BOOL:
        typestr = mrb_str_new_lit(mrb, "BOOL");
        break;
    case VT_VARIANT:
        typestr = mrb_str_new_lit(mrb, "VARIANT");
        break;
    case VT_DECIMAL:
        typestr = mrb_str_new_lit(mrb, "DECIMAL");
        break;
    case VT_I1:
        typestr = mrb_str_new_lit(mrb, "I1");
        break;
    case VT_UI1:
        typestr = mrb_str_new_lit(mrb, "UI1");
        break;
    case VT_UI2:
        typestr = mrb_str_new_lit(mrb, "UI2");
        break;
    case VT_UI4:
        typestr = mrb_str_new_lit(mrb, "UI4");
        break;
#if (_MSC_VER >= 1300) || defined(__CYGWIN__) || defined(__MINGW32__)
    case VT_I8:
        typestr = mrb_str_new_lit(mrb, "I8");
        break;
    case VT_UI8:
        typestr = mrb_str_new_lit(mrb, "UI8");
        break;
#endif
    case VT_INT:
        typestr = mrb_str_new_lit(mrb, "INT");
        break;
    case VT_UINT:
        typestr = mrb_str_new_lit(mrb, "UINT");
        break;
    case VT_VOID:
        typestr = mrb_str_new_lit(mrb, "VOID");
        break;
    case VT_HRESULT:
        typestr = mrb_str_new_lit(mrb, "HRESULT");
        break;
    case VT_PTR:
        typestr = mrb_str_new_lit(mrb, "PTR");
        if(!mrb_nil_p(typedetails))
            mrb_ary_push(mrb, typedetails, typestr);
        return ole_ptrtype2val(mrb, pTypeInfo, pTypeDesc, typedetails);
    case VT_SAFEARRAY:
        typestr = mrb_str_new_lit(mrb, "SAFEARRAY");
        if(!mrb_nil_p(typedetails))
            mrb_ary_push(mrb, typedetails, typestr);
        return ole_ptrtype2val(mrb, pTypeInfo, pTypeDesc, typedetails);
    case VT_CARRAY:
        typestr = mrb_str_new_lit(mrb, "CARRAY");
        break;
    case VT_USERDEFINED:
        typestr = mrb_str_new_lit(mrb, "USERDEFINED");
        if (!mrb_nil_p(typedetails))
            mrb_ary_push(mrb, typedetails, typestr);
        str = ole_usertype2val(mrb, pTypeInfo, pTypeDesc, typedetails);
        if (!mrb_nil_p(str)) {
            return str;
        }
        return typestr;
    case VT_UNKNOWN:
        typestr = mrb_str_new_lit(mrb, "UNKNOWN");
        break;
    case VT_DISPATCH:
        typestr = mrb_str_new_lit(mrb, "DISPATCH");
        break;
    case VT_ERROR:
        typestr = mrb_str_new_lit(mrb, "ERROR");
        break;
    case VT_LPWSTR:
        typestr = mrb_str_new_lit(mrb, "LPWSTR");
        break;
    case VT_LPSTR:
        typestr = mrb_str_new_lit(mrb, "LPSTR");
        break;
    case VT_RECORD:
        typestr = mrb_str_new_lit(mrb, "RECORD");
        break;
    default:
        typestr = mrb_str_new_lit(mrb, "Unknown Type ");
        mrb_str_concat(mrb, typestr, mrb_fixnum_to_str(mrb, INT2FIX(pTypeDesc->vt), 10));
        break;
    }
    if (!mrb_nil_p(typedetails))
        mrb_ary_push(mrb, typedetails, typestr);
    return typestr;
}

/*
 *   call-seq:
 *      WIN32OLE#ole_method_help(method)
 *
 *   Returns WIN32OLE_METHOD object corresponding with method
 *   specified by 1st argument.
 *
 *      excel = WIN32OLE.new('Excel.Application')
 *      method = excel.ole_method_help('Quit')
 *
 */
static mrb_value
fole_method_help(mrb_state *mrb, mrb_value self)
{
    mrb_value cmdname;
    ITypeInfo *pTypeInfo;
    HRESULT hr;
    struct oledata *pole;
    mrb_value obj;

    mrb_get_args(mrb, "S", &cmdname);

    OLEData_Get_Struct(mrb, self, pole);
    hr = typeinfo_from_ole(mrb, pole, &pTypeInfo);
    if(FAILED(hr))
        ole_raise(mrb, hr, E_RUNTIME_ERROR, "failed to get ITypeInfo");

    obj = create_win32ole_method(mrb, pTypeInfo, cmdname);

    OLE_RELEASE(pTypeInfo);
    if (mrb_nil_p(obj))
        mrb_raisef(mrb, E_WIN32OLE_RUNTIME_ERROR, "not found %S",
                   cmdname);
    return obj;
}

/*
 *  call-seq:
 *     WIN32OLE#ole_activex_initialize() -> mrb_nil_value()
 *
 *  Initialize WIN32OLE object(ActiveX Control) by calling
 *  IPersistMemory::InitNew.
 *
 *  Before calling OLE method, some kind of the ActiveX controls
 *  created with MFC should be initialized by calling
 *  IPersistXXX::InitNew.
 *
 *  If and only if you received the exception "HRESULT error code:
 *  0x8000ffff catastrophic failure", try this method before
 *  invoking any ole_method.
 *
 *     obj = WIN32OLE.new("ProgID_or_GUID_of_ActiveX_Control")
 *     obj.ole_activex_initialize
 *     obj.method(...)
 *
 */
static mrb_value
fole_activex_initialize(mrb_state *mrb, mrb_value self)
{
    struct oledata *pole;
    IPersistMemory *pPersistMemory;
    void *p;

    HRESULT hr = S_OK;

    OLEData_Get_Struct(mrb, self, pole);

    hr = pole->pDispatch->lpVtbl->QueryInterface(pole->pDispatch, &IID_IPersistMemory, &p);
    pPersistMemory = p;
    if (SUCCEEDED(hr)) {
        hr = pPersistMemory->lpVtbl->InitNew(pPersistMemory);
        OLE_RELEASE(pPersistMemory);
        if (SUCCEEDED(hr)) {
            return mrb_nil_value();
        }
    }

    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "fail to initialize ActiveX control");
    }

    return mrb_nil_value();
}

HRESULT
typelib_from_val(mrb_state *mrb, mrb_value obj, ITypeLib **pTypeLib)
{
    LCID lcid = cWIN32OLE_lcid;
    HRESULT hr;
    struct oledata *pole;
    unsigned int index;
    ITypeInfo *pTypeInfo;
    OLEData_Get_Struct(mrb, obj, pole);
    hr = pole->pDispatch->lpVtbl->GetTypeInfo(pole->pDispatch,
                                              0, lcid, &pTypeInfo);
    if (FAILED(hr)) {
        return hr;
    }
    hr = pTypeInfo->lpVtbl->GetContainingTypeLib(pTypeInfo, pTypeLib, &index);
    OLE_RELEASE(pTypeInfo);
    return hr;
}

static void
init_enc2cp(mrb_state *mrb)
{
    mrb_gv_set(mrb, mrb_intern_lit(mrb, "win32ole_enc2cp_table"), mrb_hash_new(mrb));
}

static void
free_enc2cp(mrb_state *mrb)
{
    mrb_hash_clear(mrb, enc2cp_table);
}

void
mrb_mruby_win32ole_gem_init(mrb_state* mrb)
{
    struct RClass *cWIN32OLE;
    cWIN32OLE_lcid = LOCALE_SYSTEM_DEFAULT;
    g_ole_initialized_init();

    com_vtbl.QueryInterface = QueryInterface;
    com_vtbl.AddRef = AddRef;
    com_vtbl.Release = Release;
    com_vtbl.GetTypeInfoCount = GetTypeInfoCount;
    com_vtbl.GetTypeInfo = GetTypeInfo;
    com_vtbl.GetIDsOfNames = GetIDsOfNames;
    com_vtbl.Invoke = Invoke;

    message_filter.QueryInterface = mf_QueryInterface;
    message_filter.AddRef = mf_AddRef;
    message_filter.Release = mf_Release;
    message_filter.HandleInComingCall = mf_HandleInComingCall;
    message_filter.RetryRejectedCall = mf_RetryRejectedCall;
    message_filter.MessagePending = mf_MessagePending;

    mrb_gv_set(mrb, mrb_intern_lit(mrb, "win32ole_com_hash"), mrb_hash_new(mrb));

    cWIN32OLE = mrb_define_class(mrb, "WIN32OLE", mrb->object_class);

    MRB_SET_INSTANCE_TT(cWIN32OLE, MRB_TT_DATA);

    mrb_define_method(mrb, cWIN32OLE, "initialize", fole_initialize, MRB_ARGS_ANY());

    mrb_define_class_method(mrb, cWIN32OLE, "connect", fole_s_connect, MRB_ARGS_ANY());
    mrb_define_class_method(mrb, cWIN32OLE, "const_load", fole_s_const_load, MRB_ARGS_ANY());

    mrb_define_class_method(mrb, cWIN32OLE, "ole_free", fole_s_free, MRB_ARGS_REQ(1));
    mrb_define_class_method(mrb, cWIN32OLE, "ole_reference_count", fole_s_reference_count, MRB_ARGS_REQ(1));
    mrb_define_class_method(mrb, cWIN32OLE, "ole_show_help", fole_s_show_help, MRB_ARGS_ANY());
    mrb_define_class_method(mrb, cWIN32OLE, "codepage", fole_s_get_code_page, MRB_ARGS_NONE());
    mrb_define_class_method(mrb, cWIN32OLE, "codepage=", fole_s_set_code_page, MRB_ARGS_REQ(1));
    mrb_define_class_method(mrb, cWIN32OLE, "locale", fole_s_get_locale, MRB_ARGS_NONE());
    mrb_define_class_method(mrb, cWIN32OLE, "locale=", fole_s_set_locale, MRB_ARGS_REQ(1));
    mrb_define_class_method(mrb, cWIN32OLE, "create_guid", fole_s_create_guid, MRB_ARGS_NONE());
    mrb_define_class_method(mrb, cWIN32OLE, "ole_initialize", fole_s_ole_initialize, MRB_ARGS_NONE());
    mrb_define_class_method(mrb, cWIN32OLE, "ole_uninitialize", fole_s_ole_uninitialize, MRB_ARGS_NONE());

    mrb_define_method(mrb, cWIN32OLE, "invoke", fole_invoke, MRB_ARGS_ANY());
    mrb_define_method(mrb, cWIN32OLE, "[]", fole_getproperty_with_bracket, MRB_ARGS_ANY());
    mrb_define_method(mrb, cWIN32OLE, "_invoke", fole_invoke2, MRB_ARGS_REQ(3));
    mrb_define_method(mrb, cWIN32OLE, "_getproperty", fole_getproperty2, MRB_ARGS_REQ(3));
    mrb_define_method(mrb, cWIN32OLE, "_setproperty", fole_setproperty2, MRB_ARGS_REQ(3));

    /* support propput method that takes an argument */
    mrb_define_method(mrb, cWIN32OLE, "[]=", fole_setproperty_with_bracket, MRB_ARGS_ANY());

    mrb_define_method(mrb, cWIN32OLE, "ole_free", fole_free, MRB_ARGS_NONE());

    mrb_define_method(mrb, cWIN32OLE, "each", fole_each, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE, "method_missing", fole_missing, MRB_ARGS_ANY());

    /* support setproperty method much like Perl ;-) */
    mrb_define_method(mrb, cWIN32OLE, "setproperty", fole_setproperty, MRB_ARGS_ANY());

    mrb_define_method(mrb, cWIN32OLE, "ole_methods", fole_methods, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE, "ole_get_methods", fole_get_methods, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE, "ole_put_methods", fole_put_methods, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE, "ole_func_methods", fole_func_methods, MRB_ARGS_NONE());

    mrb_define_method(mrb, cWIN32OLE, "ole_method", fole_method_help, MRB_ARGS_REQ(1));
    mrb_define_alias(mrb, cWIN32OLE, "ole_method_help", "ole_method");
    mrb_define_method(mrb, cWIN32OLE, "ole_activex_initialize", fole_activex_initialize, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE, "ole_type", fole_type, MRB_ARGS_NONE());
    mrb_define_alias(mrb, cWIN32OLE, "ole_obj_help", "ole_type");
    mrb_define_method(mrb, cWIN32OLE, "ole_typelib", fole_typelib, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE, "ole_query_interface", fole_query_interface, MRB_ARGS_REQ(1));
    mrb_define_method(mrb, cWIN32OLE, "ole_respond_to?", fole_respond_to, MRB_ARGS_REQ(1));

    /* Constants definition */

    /*
     * Version string of WIN32OLE.
     */
    mrb_define_const(mrb, cWIN32OLE, "VERSION", mrb_str_new_lit(mrb, WIN32OLE_VERSION));

    /*
     * After invoking OLE methods with reference arguments, you can access
     * the value of arguments by using ARGV.
     *
     * If the method of OLE(COM) server written by C#.NET is following:
     *
     *   void calcsum(int a, int b, out int c) {
     *       c = a + b;
     *   }
     *
     * then, the Ruby OLE(COM) client script to retrieve the value of
     * argument c after invoking calcsum method is following:
     *
     *   a = 10
     *   b = 20
     *   c = 0
     *   comserver.calcsum(a, b, c)
     *   p c # => 0
     *   p WIN32OLE::ARGV # => [10, 20, 30]
     *
     * You can use WIN32OLE_VARIANT object to retrieve the value of reference
     * arguments instead of refering WIN32OLE::ARGV.
     *
     */
    mrb_define_const(mrb, cWIN32OLE, "ARGV", mrb_ary_new(mrb));

    /*
     * 0: ANSI code page. See WIN32OLE.codepage and WIN32OLE.codepage=.
     */
    mrb_define_const(mrb, cWIN32OLE, "CP_ACP", INT2FIX(CP_ACP));

    /*
     * 1: OEM code page. See WIN32OLE.codepage and WIN32OLE.codepage=.
     */
    mrb_define_const(mrb, cWIN32OLE, "CP_OEMCP", INT2FIX(CP_OEMCP));

    /*
     * 2
     */
    mrb_define_const(mrb, cWIN32OLE, "CP_MACCP", INT2FIX(CP_MACCP));

    /*
     * 3: current thread ANSI code page. See WIN32OLE.codepage and
     * WIN32OLE.codepage=.
     */
    mrb_define_const(mrb, cWIN32OLE, "CP_THREAD_ACP", INT2FIX(CP_THREAD_ACP));

    /*
     * 42: symbol code page. See WIN32OLE.codepage and WIN32OLE.codepage=.
     */
    mrb_define_const(mrb, cWIN32OLE, "CP_SYMBOL", INT2FIX(CP_SYMBOL));

    /*
     * 65000: UTF-7 code page. See WIN32OLE.codepage and WIN32OLE.codepage=.
     */
    mrb_define_const(mrb, cWIN32OLE, "CP_UTF7", INT2FIX(CP_UTF7));

    /*
     * 65001: UTF-8 code page. See WIN32OLE.codepage and WIN32OLE.codepage=.
     */
    mrb_define_const(mrb, cWIN32OLE, "CP_UTF8", INT2FIX(CP_UTF8));

    /*
     * 0x0800: default locale for the operating system. See WIN32OLE.locale
     * and WIN32OLE.locale=.
     */
    mrb_define_const(mrb, cWIN32OLE, "LOCALE_SYSTEM_DEFAULT", INT2FIX(LOCALE_SYSTEM_DEFAULT));

    /*
     * 0x0400: default locale for the user or process. See WIN32OLE.locale
     * and WIN32OLE.locale=.
     */
    mrb_define_const(mrb, cWIN32OLE, "LOCALE_USER_DEFAULT", INT2FIX(LOCALE_USER_DEFAULT));

    Init_win32ole_variant_m(mrb);
    Init_win32ole_typelib(mrb);
    Init_win32ole_type(mrb);
    Init_win32ole_variable(mrb);
    Init_win32ole_method(mrb);
    Init_win32ole_param(mrb);
    Init_win32ole_event(mrb);
    Init_win32ole_variant(mrb);
    Init_win32ole_record(mrb);
    Init_win32ole_error(mrb);

    init_enc2cp(mrb);
    ole_init_cp(mrb);
}

void
mrb_mruby_win32ole_gem_final(mrb_state* mrb)
{
	free_enc2cp(mrb);
	ole_uninitialize();
}

