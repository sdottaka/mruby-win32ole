#include "win32ole.h"

static void olemethod_free(mrb_state *mrb, void *ptr);
static mrb_value ole_method_sub(mrb_state *mrb, mrb_value self, ITypeInfo *pOwnerTypeInfo, ITypeInfo *pTypeInfo, mrb_value name);
static mrb_value olemethod_from_typeinfo(mrb_state *mrb, mrb_value self, ITypeInfo *pTypeInfo, mrb_value name);
static mrb_value ole_methods_sub(mrb_state *mrb, ITypeInfo *pOwnerTypeInfo, ITypeInfo *pTypeInfo, mrb_value methods, int mask);
static mrb_value olemethod_set_member(mrb_state *mrb, mrb_value self, ITypeInfo *pTypeInfo, ITypeInfo *pOwnerTypeInfo, int index, mrb_value name);
static mrb_value folemethod_initialize(mrb_state *mrb, mrb_value self);
static mrb_value folemethod_name(mrb_state *mrb, mrb_value self);
static mrb_value ole_method_return_type(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index);
static mrb_value folemethod_return_type(mrb_state *mrb, mrb_value self);
static mrb_value ole_method_return_vtype(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index);
static mrb_value folemethod_return_vtype(mrb_state *mrb, mrb_value self);
static mrb_value ole_method_return_type_detail(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index);
static mrb_value folemethod_return_type_detail(mrb_state *mrb, mrb_value self);
static mrb_value ole_method_invkind(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index);
static mrb_value ole_method_invoke_kind(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index);
static mrb_value folemethod_invkind(mrb_state *mrb, mrb_value self);
static mrb_value folemethod_invoke_kind(mrb_state *mrb, mrb_value self);
static mrb_value ole_method_visible(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index);
static mrb_value folemethod_visible(mrb_state *mrb, mrb_value self);
static mrb_value ole_method_event(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index, mrb_value method_name);
static mrb_value folemethod_event(mrb_state *mrb, mrb_value self);
static mrb_value folemethod_event_interface(mrb_state *mrb, mrb_value self);
static HRESULT ole_method_docinfo_from_type(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index, BSTR *name, BSTR *helpstr, DWORD *helpcontext, BSTR *helpfile);
static mrb_value ole_method_helpstring(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index);
static mrb_value folemethod_helpstring(mrb_state *mrb, mrb_value self);
static mrb_value ole_method_helpfile(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index);
static mrb_value folemethod_helpfile(mrb_state *mrb, mrb_value self);
static mrb_value ole_method_helpcontext(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index);
static mrb_value folemethod_helpcontext(mrb_state *mrb, mrb_value self);
static mrb_value ole_method_dispid(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index);
static mrb_value folemethod_dispid(mrb_state *mrb, mrb_value self);
static mrb_value ole_method_offset_vtbl(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index);
static mrb_value folemethod_offset_vtbl(mrb_state *mrb, mrb_value self);
static mrb_value ole_method_size_params(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index);
static mrb_value folemethod_size_params(mrb_state *mrb, mrb_value self);
static mrb_value ole_method_size_opt_params(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index);
static mrb_value folemethod_size_opt_params(mrb_state *mrb, mrb_value self);
static mrb_value ole_method_params(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index);
static mrb_value folemethod_params(mrb_state *mrb, mrb_value self);
static mrb_value folemethod_inspect(mrb_state *mrb, mrb_value self);

static const mrb_data_type olemethod_datatype = {
    "win32ole_method",
    olemethod_free
};

static void
olemethod_free(mrb_state *mrb, void *ptr)
{
    struct olemethoddata *polemethod = ptr;
    OLE_FREE(polemethod->pTypeInfo);
    OLE_FREE(polemethod->pOwnerTypeInfo);
    free(polemethod);
}

struct olemethoddata *
olemethod_data_get_struct(mrb_state *mrb, mrb_value obj)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, obj, &olemethod_datatype, pmethod);
    return pmethod;
}

static mrb_value
ole_method_sub(mrb_state *mrb, mrb_value self, ITypeInfo *pOwnerTypeInfo, ITypeInfo *pTypeInfo, mrb_value name)
{
    HRESULT hr;
    TYPEATTR *pTypeAttr;
    BSTR bstr;
    FUNCDESC *pFuncDesc;
    WORD i;
    mrb_value fname;
    mrb_value method = mrb_nil_value();
    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to GetTypeAttr");
    }
    for(i = 0; i < pTypeAttr->cFuncs && mrb_nil_p(method); i++) {
        hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, i, &pFuncDesc);
        if (FAILED(hr))
             continue;

        hr = pTypeInfo->lpVtbl->GetDocumentation(pTypeInfo, pFuncDesc->memid,
                                                 &bstr, NULL, NULL, NULL);
        if (FAILED(hr)) {
            pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
            continue;
        }
        fname = WC2VSTR(mrb, bstr);
        if (strcasecmp(mrb_string_value_ptr(mrb, name), mrb_string_value_ptr(mrb, fname)) == 0) {
            olemethod_set_member(mrb, self, pTypeInfo, pOwnerTypeInfo, i, fname);
            method = self;
        }
        pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
        pFuncDesc=NULL;
    }
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
    return method;
}

mrb_value
ole_methods_from_typeinfo(mrb_state *mrb, ITypeInfo *pTypeInfo, int mask)
{
    HRESULT hr;
    TYPEATTR *pTypeAttr;
    WORD i;
    HREFTYPE href;
    ITypeInfo *pRefTypeInfo;
    mrb_value methods = mrb_ary_new(mrb);
    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to GetTypeAttr");
    }

    ole_methods_sub(mrb, 0, pTypeInfo, methods, mask);
    for(i=0; i < pTypeAttr->cImplTypes; i++){
       hr = pTypeInfo->lpVtbl->GetRefTypeOfImplType(pTypeInfo, i, &href);
       if(FAILED(hr))
           continue;
       hr = pTypeInfo->lpVtbl->GetRefTypeInfo(pTypeInfo, href, &pRefTypeInfo);
       if (FAILED(hr))
           continue;
       ole_methods_sub(mrb, pTypeInfo, pRefTypeInfo, methods, mask);
       OLE_RELEASE(pRefTypeInfo);
    }
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
    return methods;
}

static mrb_value
olemethod_from_typeinfo(mrb_state *mrb, mrb_value self, ITypeInfo *pTypeInfo, mrb_value name)
{
    HRESULT hr;
    TYPEATTR *pTypeAttr;
    WORD i;
    HREFTYPE href;
    ITypeInfo *pRefTypeInfo;
    mrb_value method = mrb_nil_value();
    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to GetTypeAttr");
    }
    method = ole_method_sub(mrb, self, 0, pTypeInfo, name);
    if (!mrb_nil_p(method)) {
       return method;
    }
    for(i=0; i < pTypeAttr->cImplTypes && mrb_nil_p(method); i++){
       hr = pTypeInfo->lpVtbl->GetRefTypeOfImplType(pTypeInfo, i, &href);
       if(FAILED(hr))
           continue;
       hr = pTypeInfo->lpVtbl->GetRefTypeInfo(pTypeInfo, href, &pRefTypeInfo);
       if (FAILED(hr))
           continue;
       method = ole_method_sub(mrb, self, pTypeInfo, pRefTypeInfo, name);
       OLE_RELEASE(pRefTypeInfo);
    }
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
    return method;
}

static mrb_value
ole_methods_sub(mrb_state *mrb, ITypeInfo *pOwnerTypeInfo, ITypeInfo *pTypeInfo, mrb_value methods, int mask)
{
    HRESULT hr;
    TYPEATTR *pTypeAttr;
    BSTR bstr;
    FUNCDESC *pFuncDesc;
    mrb_value method;
    WORD i;
    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to GetTypeAttr");
    }
    for(i = 0; i < pTypeAttr->cFuncs; i++) {
        hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, i, &pFuncDesc);
        if (FAILED(hr))
             continue;

        hr = pTypeInfo->lpVtbl->GetDocumentation(pTypeInfo, pFuncDesc->memid,
                                                 &bstr, NULL, NULL, NULL);
        if (FAILED(hr)) {
            pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
            continue;
        }
        if(pFuncDesc->invkind & mask) {
            method = folemethod_s_allocate(mrb, C_WIN32OLE_METHOD);
            olemethod_set_member(mrb, method, pTypeInfo, pOwnerTypeInfo,
                                 i, WC2VSTR(mrb, bstr));
            mrb_ary_push(mrb, methods, method);
        }
        pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
        pFuncDesc=NULL;
    }
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);

    return methods;
}

mrb_value
create_win32ole_method(mrb_state *mrb, ITypeInfo *pTypeInfo, mrb_value name)
{

    mrb_value method = folemethod_s_allocate(mrb, C_WIN32OLE_METHOD);
    mrb_value obj = olemethod_from_typeinfo(mrb, method, pTypeInfo, name);
    return obj;
}

/*
 * Document-class: WIN32OLE_METHOD
 *
 *   <code>WIN32OLE_METHOD</code> objects represent OLE method information.
 */

static mrb_value
olemethod_set_member(mrb_state *mrb, mrb_value self, ITypeInfo *pTypeInfo, ITypeInfo *pOwnerTypeInfo, int index, mrb_value name)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    pmethod->pTypeInfo = pTypeInfo;
    OLE_ADDREF(pTypeInfo);
    pmethod->pOwnerTypeInfo = pOwnerTypeInfo;
    OLE_ADDREF(pOwnerTypeInfo);
    pmethod->index = index;
    mrb_iv_set(mrb, self, mrb_intern_lit(mrb, "name"), name);
    return self;
}

static struct olemethoddata *
olemethoddata_alloc(mrb_state *mrb)
{
    struct olemethoddata *pmethod = ALLOC(struct olemethoddata);
    pmethod->pTypeInfo = NULL;
    pmethod->pOwnerTypeInfo = NULL;
    pmethod->index = 0;
    return pmethod;
}

mrb_value
folemethod_s_allocate(mrb_state *mrb, struct RClass *klass)
{
    mrb_value obj;
    struct olemethoddata *pmethod = olemethoddata_alloc(mrb);
    obj = mrb_obj_value(Data_Wrap_Struct(mrb, klass,
                                &olemethod_datatype, pmethod));
    return obj;
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD.new(ole_type,  method) -> WIN32OLE_METHOD object
 *
 *  Returns a new WIN32OLE_METHOD object which represents the information
 *  about OLE method.
 *  The first argument <i>ole_type</i> specifies WIN32OLE_TYPE object.
 *  The second argument <i>method</i> specifies OLE method name defined OLE class
 *  which represents WIN32OLE_TYPE object.
 *
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbook')
 *     method = WIN32OLE_METHOD.new(tobj, 'SaveAs')
 */
static mrb_value
folemethod_initialize(mrb_state *mrb, mrb_value self)
{
    mrb_value oletype;
    mrb_value method;
    mrb_value obj = mrb_nil_value();
    ITypeInfo *pTypeInfo;
    struct olemethoddata *pmethod = (struct olemethoddata *)DATA_PTR(self);
    if (pmethod) {
        mrb_free(mrb, pmethod);
    }
    mrb_data_init(self, NULL, &olemethod_datatype);

    mrb_get_args(mrb, "oS", &oletype, &method);

    pmethod = olemethoddata_alloc(mrb);
    mrb_data_init(self, pmethod, &olemethod_datatype); /* FIXME: timing */

    if (mrb_obj_is_kind_of(mrb, oletype, C_WIN32OLE_TYPE)) {
        pTypeInfo = itypeinfo(mrb, oletype);
        obj = olemethod_from_typeinfo(mrb, self, pTypeInfo, method);
        if (mrb_nil_p(obj)) {
            mrb_raisef(mrb, E_WIN32OLE_RUNTIME_ERROR, "not found %S",
                     method);
        }
    }
    else {
        mrb_raise(mrb, E_TYPE_ERROR, "1st argument should be WIN32OLE_TYPE object");
    }
    return obj;
}

/*
 *  call-seq
 *     WIN32OLE_METHOD#name
 *
 *  Returns the name of the method.
 *
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbook')
 *     method = WIN32OLE_METHOD.new(tobj, 'SaveAs')
 *     puts method.name # => SaveAs
 *
 */
static mrb_value
folemethod_name(mrb_state *mrb, mrb_value self)
{
    return mrb_iv_get(mrb, self, mrb_intern_lit(mrb, "name"));
}

static mrb_value
ole_method_return_type(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index)
{
    FUNCDESC *pFuncDesc;
    HRESULT hr;
    mrb_value type;

    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if (FAILED(hr))
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to GetFuncDesc");

    type = ole_typedesc2val(mrb, pTypeInfo, &(pFuncDesc->elemdescFunc.tdesc), mrb_nil_value());
    pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
    return type;
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#return_type
 *
 *  Returns string of return value type of method.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbooks')
 *     method = WIN32OLE_METHOD.new(tobj, 'Add')
 *     puts method.return_type # => Workbook
 *
 */
static mrb_value
folemethod_return_type(mrb_state *mrb, mrb_value self)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    return ole_method_return_type(mrb, pmethod->pTypeInfo, pmethod->index);
}

static mrb_value
ole_method_return_vtype(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index)
{
    FUNCDESC *pFuncDesc;
    HRESULT hr;
    mrb_value vvt;

    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if (FAILED(hr))
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to GetFuncDesc");

    vvt = INT2FIX(pFuncDesc->elemdescFunc.tdesc.vt);
    pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
    return vvt;
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#return_vtype
 *
 *  Returns number of return value type of method.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbooks')
 *     method = WIN32OLE_METHOD.new(tobj, 'Add')
 *     puts method.return_vtype # => 26
 *
 */
static mrb_value
folemethod_return_vtype(mrb_state *mrb, mrb_value self)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    return ole_method_return_vtype(mrb, pmethod->pTypeInfo, pmethod->index);
}

static mrb_value
ole_method_return_type_detail(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index)
{
    FUNCDESC *pFuncDesc;
    HRESULT hr;
    mrb_value type = mrb_ary_new(mrb);

    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if (FAILED(hr))
        return type;

    ole_typedesc2val(mrb, pTypeInfo, &(pFuncDesc->elemdescFunc.tdesc), type);
    pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
    return type;
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#return_type_detail
 *
 *  Returns detail information of return value type of method.
 *  The information is array.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbooks')
 *     method = WIN32OLE_METHOD.new(tobj, 'Add')
 *     p method.return_type_detail # => ["PTR", "USERDEFINED", "Workbook"]
 */
static mrb_value
folemethod_return_type_detail(mrb_state *mrb, mrb_value self)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    return ole_method_return_type_detail(mrb, pmethod->pTypeInfo, pmethod->index);
}

static mrb_value
ole_method_invkind(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index)
{
    FUNCDESC *pFuncDesc;
    HRESULT hr;
    mrb_value invkind;
    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if(FAILED(hr))
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to GetFuncDesc");
    invkind = INT2FIX(pFuncDesc->invkind);
    pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
    return invkind;
}

static mrb_value
ole_method_invoke_kind(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index)
{
    mrb_value type = mrb_str_new_lit(mrb, "UNKNOWN");
    mrb_value invkind = ole_method_invkind(mrb, pTypeInfo, method_index);
    if((FIX2INT(invkind) & INVOKE_PROPERTYGET) &&
       (FIX2INT(invkind) & INVOKE_PROPERTYPUT) ) {
        type = mrb_str_new_lit(mrb, "PROPERTY");
    } else if(FIX2INT(invkind) & INVOKE_PROPERTYGET) {
        type =  mrb_str_new_lit(mrb, "PROPERTYGET");
    } else if(FIX2INT(invkind) & INVOKE_PROPERTYPUT) {
        type = mrb_str_new_lit(mrb, "PROPERTYPUT");
    } else if(FIX2INT(invkind) & INVOKE_PROPERTYPUTREF) {
        type = mrb_str_new_lit(mrb, "PROPERTYPUTREF");
    } else if(FIX2INT(invkind) & INVOKE_FUNC) {
        type = mrb_str_new_lit(mrb, "FUNC");
    }
    return type;
}

/*
 *   call-seq:
 *      WIN32OLE_MTHOD#invkind
 *
 *   Returns the method invoke kind.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbooks')
 *     method = WIN32OLE_METHOD.new(tobj, 'Add')
 *     puts method.invkind # => 1
 *
 */
static mrb_value
folemethod_invkind(mrb_state *mrb, mrb_value self)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    return ole_method_invkind(mrb, pmethod->pTypeInfo, pmethod->index);
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#invoke_kind
 *
 *  Returns the method kind string. The string is "UNKNOWN" or "PROPERTY"
 *  or "PROPERTY" or "PROPERTYGET" or "PROPERTYPUT" or "PROPERTYPPUTREF"
 *  or "FUNC".
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbooks')
 *     method = WIN32OLE_METHOD.new(tobj, 'Add')
 *     puts method.invoke_kind # => "FUNC"
 */
static mrb_value
folemethod_invoke_kind(mrb_state *mrb, mrb_value self)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    return ole_method_invoke_kind(mrb, pmethod->pTypeInfo, pmethod->index);
}

static mrb_value
ole_method_visible(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index)
{
    FUNCDESC *pFuncDesc;
    HRESULT hr;
    mrb_value visible;
    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if(FAILED(hr))
        return mrb_false_value();
    if (pFuncDesc->wFuncFlags & (FUNCFLAG_FRESTRICTED |
                                 FUNCFLAG_FHIDDEN |
                                 FUNCFLAG_FNONBROWSABLE)) {
        visible = mrb_false_value();
    } else {
        visible = mrb_true_value();
    }
    pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
    return visible;
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#visible?
 *
 *  Returns true if the method is public.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbooks')
 *     method = WIN32OLE_METHOD.new(tobj, 'Add')
 *     puts method.visible? # => true
 */
static mrb_value
folemethod_visible(mrb_state *mrb, mrb_value self)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    return ole_method_visible(mrb, pmethod->pTypeInfo, pmethod->index);
}

static mrb_value
ole_method_event(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index, mrb_value method_name)
{
    TYPEATTR *pTypeAttr;
    HRESULT hr;
    WORD i;
    int flags;
    HREFTYPE href;
    ITypeInfo *pRefTypeInfo;
    FUNCDESC *pFuncDesc;
    BSTR bstr;
    mrb_value name;
    mrb_value event = mrb_false_value();

    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if (FAILED(hr))
        return event;
    if(pTypeAttr->typekind != TKIND_COCLASS) {
        pTypeInfo->lpVtbl->ReleaseTypeAttr(pTypeInfo, pTypeAttr);
        return event;
    }
    for (i = 0; i < pTypeAttr->cImplTypes; i++) {
        hr = pTypeInfo->lpVtbl->GetImplTypeFlags(pTypeInfo, i, &flags);
        if (FAILED(hr))
            continue;

        if (flags & IMPLTYPEFLAG_FSOURCE) {
            hr = pTypeInfo->lpVtbl->GetRefTypeOfImplType(pTypeInfo,
                                                         i, &href);
            if (FAILED(hr))
                continue;
            hr = pTypeInfo->lpVtbl->GetRefTypeInfo(pTypeInfo,
                                                   href, &pRefTypeInfo);
            if (FAILED(hr))
                continue;
            hr = pRefTypeInfo->lpVtbl->GetFuncDesc(pRefTypeInfo, method_index,
                                                   &pFuncDesc);
            if (FAILED(hr)) {
                OLE_RELEASE(pRefTypeInfo);
                continue;
            }

            hr = pRefTypeInfo->lpVtbl->GetDocumentation(pRefTypeInfo,
                                                        pFuncDesc->memid,
                                                        &bstr, NULL, NULL, NULL);
            if (FAILED(hr)) {
                pRefTypeInfo->lpVtbl->ReleaseFuncDesc(pRefTypeInfo, pFuncDesc);
                OLE_RELEASE(pRefTypeInfo);
                continue;
            }

            name = WC2VSTR(mrb, bstr);
            pRefTypeInfo->lpVtbl->ReleaseFuncDesc(pRefTypeInfo, pFuncDesc);
            OLE_RELEASE(pRefTypeInfo);
            if (mrb_str_cmp(mrb, method_name, name) == 0) {
                event = mrb_true_value();
                break;
            }
        }
    }
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
    return event;
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#event?
 *
 *  Returns true if the method is event.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbook')
 *     method = WIN32OLE_METHOD.new(tobj, 'SheetActivate')
 *     puts method.event? # => true
 *
 */
static mrb_value
folemethod_event(mrb_state *mrb, mrb_value self)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    if (!pmethod->pOwnerTypeInfo)
        return mrb_false_value();
    return ole_method_event(mrb, pmethod->pOwnerTypeInfo,
                            pmethod->index,
                            mrb_iv_get(mrb, self, mrb_intern_lit(mrb, "name")));
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#event_interface
 *
 *  Returns event interface name if the method is event.
 *    tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbook')
 *    method = WIN32OLE_METHOD.new(tobj, 'SheetActivate')
 *    puts method.event_interface # =>  WorkbookEvents
 */
static mrb_value
folemethod_event_interface(mrb_state *mrb, mrb_value self)
{
    BSTR name;
    struct olemethoddata *pmethod;
    HRESULT hr;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    if(mrb_bool(folemethod_event(mrb, self))) {
        hr = ole_docinfo_from_type(mrb, pmethod->pTypeInfo, &name, NULL, NULL, NULL);
        if(SUCCEEDED(hr))
            return WC2VSTR(mrb, name);
    }
    return mrb_nil_value();
}

static HRESULT
ole_method_docinfo_from_type(
    mrb_state *mrb, 
    ITypeInfo *pTypeInfo,
    UINT method_index,
    BSTR *name,
    BSTR *helpstr,
    DWORD *helpcontext,
    BSTR *helpfile
    )
{
    FUNCDESC *pFuncDesc;
    HRESULT hr;
    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if (FAILED(hr))
        return hr;
    hr = pTypeInfo->lpVtbl->GetDocumentation(pTypeInfo, pFuncDesc->memid,
                                             name, helpstr,
                                             helpcontext, helpfile);
    pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
    return hr;
}

static mrb_value
ole_method_helpstring(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index)
{
    HRESULT hr;
    BSTR bhelpstring;
    hr = ole_method_docinfo_from_type(mrb, pTypeInfo, method_index, NULL, &bhelpstring,
                                      NULL, NULL);
    if (FAILED(hr))
        return mrb_nil_value();
    return WC2VSTR(mrb, bhelpstring);
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#helpstring
 *
 *  Returns help string of OLE method. If the help string is not found,
 *  then the method returns nil.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Internet Controls', 'IWebBrowser')
 *     method = WIN32OLE_METHOD.new(tobj, 'Navigate')
 *     puts method.helpstring # => Navigates to a URL or file.
 *
 */
static mrb_value
folemethod_helpstring(mrb_state *mrb, mrb_value self)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    return ole_method_helpstring(mrb, pmethod->pTypeInfo, pmethod->index);
}

static mrb_value
ole_method_helpfile(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index)
{
    HRESULT hr;
    BSTR bhelpfile;
    hr = ole_method_docinfo_from_type(mrb, pTypeInfo, method_index, NULL, NULL,
                                      NULL, &bhelpfile);
    if (FAILED(hr))
        return mrb_nil_value();
    return WC2VSTR(mrb, bhelpfile);
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#helpfile
 *
 *  Returns help file. If help file is not found, then
 *  the method returns nil.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbooks')
 *     method = WIN32OLE_METHOD.new(tobj, 'Add')
 *     puts method.helpfile # => C:\...\VBAXL9.CHM
 */
static mrb_value
folemethod_helpfile(mrb_state *mrb, mrb_value self)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);

    return ole_method_helpfile(mrb, pmethod->pTypeInfo, pmethod->index);
}

static mrb_value
ole_method_helpcontext(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index)
{
    HRESULT hr;
    DWORD helpcontext = 0;
    hr = ole_method_docinfo_from_type(mrb, pTypeInfo, method_index, NULL, NULL,
                                      &helpcontext, NULL);
    if (FAILED(hr))
        return mrb_nil_value();
    return INT2FIX(helpcontext);
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#helpcontext
 *
 *  Returns help context.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbooks')
 *     method = WIN32OLE_METHOD.new(tobj, 'Add')
 *     puts method.helpcontext # => 65717
 */
static mrb_value
folemethod_helpcontext(mrb_state *mrb, mrb_value self)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    return ole_method_helpcontext(mrb, pmethod->pTypeInfo, pmethod->index);
}

static mrb_value
ole_method_dispid(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index)
{
    FUNCDESC *pFuncDesc;
    HRESULT hr;
    mrb_value dispid = mrb_nil_value();
    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if (FAILED(hr))
        return dispid;
    dispid = INT2NUM(pFuncDesc->memid);
    pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
    return dispid;
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#dispid
 *
 *  Returns dispatch ID.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbooks')
 *     method = WIN32OLE_METHOD.new(tobj, 'Add')
 *     puts method.dispid # => 181
 */
static mrb_value
folemethod_dispid(mrb_state *mrb, mrb_value self)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    return ole_method_dispid(mrb, pmethod->pTypeInfo, pmethod->index);
}

static mrb_value
ole_method_offset_vtbl(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index)
{
    FUNCDESC *pFuncDesc;
    HRESULT hr;
    mrb_value offset_vtbl = mrb_nil_value();
    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if (FAILED(hr))
        return offset_vtbl;
    offset_vtbl = INT2FIX(pFuncDesc->oVft);
    pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
    return offset_vtbl;
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#offset_vtbl
 *
 *  Returns the offset ov VTBL.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbooks')
 *     method = WIN32OLE_METHOD.new(tobj, 'Add')
 *     puts method.offset_vtbl # => 40
 */
static mrb_value
folemethod_offset_vtbl(mrb_state *mrb, mrb_value self)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    return ole_method_offset_vtbl(mrb, pmethod->pTypeInfo, pmethod->index);
}

static mrb_value
ole_method_size_params(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index)
{
    FUNCDESC *pFuncDesc;
    HRESULT hr;
    mrb_value size_params = mrb_nil_value();
    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if (FAILED(hr))
        return size_params;
    size_params = INT2FIX(pFuncDesc->cParams);
    pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
    return size_params;
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#size_params
 *
 *  Returns the size of arguments of the method.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbook')
 *     method = WIN32OLE_METHOD.new(tobj, 'SaveAs')
 *     puts method.size_params # => 11
 *
 */
static mrb_value
folemethod_size_params(mrb_state *mrb, mrb_value self)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    return ole_method_size_params(mrb, pmethod->pTypeInfo, pmethod->index);
}

static mrb_value
ole_method_size_opt_params(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index)
{
    FUNCDESC *pFuncDesc;
    HRESULT hr;
    mrb_value size_opt_params = mrb_nil_value();
    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if (FAILED(hr))
        return size_opt_params;
    size_opt_params = INT2FIX(pFuncDesc->cParamsOpt);
    pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
    return size_opt_params;
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#size_opt_params
 *
 *  Returns the size of optional parameters.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbook')
 *     method = WIN32OLE_METHOD.new(tobj, 'SaveAs')
 *     puts method.size_opt_params # => 4
 */
static mrb_value
folemethod_size_opt_params(mrb_state *mrb, mrb_value self)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    return ole_method_size_opt_params(mrb, pmethod->pTypeInfo, pmethod->index);
}

static mrb_value
ole_method_params(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index)
{
    FUNCDESC *pFuncDesc;
    HRESULT hr;
    BSTR *bstrs;
    UINT len, i;
    mrb_value param;
    mrb_value params = mrb_ary_new(mrb);
    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if (FAILED(hr))
        return params;

    len = 0;
    bstrs = ALLOCA_N(BSTR, pFuncDesc->cParams + 1);
    hr = pTypeInfo->lpVtbl->GetNames(pTypeInfo, pFuncDesc->memid,
                                     bstrs, pFuncDesc->cParams + 1,
                                     &len);
    if (FAILED(hr)) {
        pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
        return params;
    }
    SysFreeString(bstrs[0]);
    if (pFuncDesc->cParams > 0) {
        for(i = 1; i < len; i++) {
            param = create_win32ole_param(mrb, pTypeInfo, method_index, i-1, WC2VSTR(mrb, bstrs[i]));
            mrb_ary_push(mrb, params, param);
         }
     }
     pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
     return params;
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#params
 *
 *  returns array of WIN32OLE_PARAM object corresponding with method parameters.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbook')
 *     method = WIN32OLE_METHOD.new(tobj, 'SaveAs')
 *     p method.params # => [Filename, FileFormat, Password, WriteResPassword,
 *                           ReadOnlyRecommended, CreateBackup, AccessMode,
 *                           ConflictResolution, AddToMru, TextCodepage,
 *                           TextVisualLayout]
 */
static mrb_value
folemethod_params(mrb_state *mrb, mrb_value self)
{
    struct olemethoddata *pmethod;
    Data_Get_Struct(mrb, self, &olemethod_datatype, pmethod);
    return ole_method_params(mrb, pmethod->pTypeInfo, pmethod->index);
}

/*
 *  call-seq:
 *     WIN32OLE_METHOD#inspect -> String
 *
 *  Returns the method name with class name.
 *
 */
static mrb_value
folemethod_inspect(mrb_state *mrb, mrb_value self)
{
    return default_inspect(mrb, self, "WIN32OLE_METHOD");
}

void Init_win32ole_method(mrb_state *mrb)
{
    struct RClass *cWIN32OLE_METHOD = mrb_define_class(mrb, "WIN32OLE_METHOD", mrb->object_class);
    MRB_SET_INSTANCE_TT(cWIN32OLE_METHOD, MRB_TT_DATA);
    mrb_define_method(mrb, cWIN32OLE_METHOD, "initialize", folemethod_initialize, MRB_ARGS_REQ(2));
    mrb_define_method(mrb, cWIN32OLE_METHOD, "name", folemethod_name, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "return_type", folemethod_return_type, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "return_vtype", folemethod_return_vtype, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "return_type_detail", folemethod_return_type_detail, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "invoke_kind", folemethod_invoke_kind, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "invkind", folemethod_invkind, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "visible?", folemethod_visible, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "event?", folemethod_event, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "event_interface", folemethod_event_interface, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "helpstring", folemethod_helpstring, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "helpfile", folemethod_helpfile, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "helpcontext", folemethod_helpcontext, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "dispid", folemethod_dispid, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "offset_vtbl", folemethod_offset_vtbl, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "size_params", folemethod_size_params, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "size_opt_params", folemethod_size_opt_params, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_METHOD, "params", folemethod_params, MRB_ARGS_NONE());
    mrb_define_alias(mrb, cWIN32OLE_METHOD, "to_s", "name");
    mrb_define_method(mrb, cWIN32OLE_METHOD, "inspect", folemethod_inspect, MRB_ARGS_NONE());
}
