#include "win32ole.h"

struct oletypedata {
    ITypeInfo *pTypeInfo;
};

static void oletype_free(mrb_state *mrb, void *ptr);
static mrb_value foletype_s_ole_classes(mrb_state *mrb, mrb_value self);
static mrb_value foletype_s_typelibs(mrb_state *mrb, mrb_value self);
static mrb_value foletype_s_progids(mrb_state *mrb, mrb_value self);
static mrb_value oletype_set_member(mrb_state *mrb, mrb_value self, ITypeInfo *pTypeInfo, mrb_value name);
static struct oletypedata *oletypedata_alloc(mrb_state *mrb);
static mrb_value foletype_s_allocate(mrb_state *mrb, struct RClass *klass);
static mrb_value oleclass_from_typelib(mrb_state *mrb, mrb_value self, ITypeLib *pTypeLib, mrb_value oleclass);
static mrb_value foletype_initialize(mrb_state *mrb, mrb_value self);
static mrb_value foletype_name(mrb_state *mrb, mrb_value self);
static mrb_value ole_ole_type(mrb_state *mrb, ITypeInfo *pTypeInfo);
static mrb_value foletype_ole_type(mrb_state *mrb, mrb_value self);
static mrb_value ole_type_guid(mrb_state *mrb, ITypeInfo *pTypeInfo);
static mrb_value foletype_guid(mrb_state *mrb, mrb_value self);
static mrb_value ole_type_progid(mrb_state *mrb, ITypeInfo *pTypeInfo);
static mrb_value foletype_progid(mrb_state *mrb, mrb_value self);
static mrb_value ole_type_visible(mrb_state *mrb, ITypeInfo *pTypeInfo);
static mrb_value foletype_visible(mrb_state *mrb, mrb_value self);
static mrb_value ole_type_major_version(mrb_state *mrb, ITypeInfo *pTypeInfo);
static mrb_value foletype_major_version(mrb_state *mrb, mrb_value self);
static mrb_value ole_type_minor_version(mrb_state *mrb, ITypeInfo *pTypeInfo);
static mrb_value foletype_minor_version(mrb_state *mrb, mrb_value self);
static mrb_value ole_type_typekind(mrb_state *mrb, ITypeInfo *pTypeInfo);
static mrb_value foletype_typekind(mrb_state *mrb, mrb_value self);
static mrb_value ole_type_helpstring(mrb_state *mrb, ITypeInfo *pTypeInfo);
static mrb_value foletype_helpstring(mrb_state *mrb, mrb_value self);
static mrb_value ole_type_src_type(mrb_state *mrb, ITypeInfo *pTypeInfo);
static mrb_value foletype_src_type(mrb_state *mrb, mrb_value self);
static mrb_value ole_type_helpfile(mrb_state *mrb, ITypeInfo *pTypeInfo);
static mrb_value foletype_helpfile(mrb_state *mrb, mrb_value self);
static mrb_value ole_type_helpcontext(mrb_state *mrb, ITypeInfo *pTypeInfo);
static mrb_value foletype_helpcontext(mrb_state *mrb, mrb_value self);
static mrb_value ole_variables(mrb_state *mrb, ITypeInfo *pTypeInfo);
static mrb_value foletype_variables(mrb_state *mrb, mrb_value self);
static mrb_value foletype_methods(mrb_state *mrb, mrb_value self);
static mrb_value foletype_ole_typelib(mrb_state *mrb, mrb_value self);
static mrb_value ole_type_impl_ole_types(mrb_state *mrb, ITypeInfo *pTypeInfo, int implflags);
static mrb_value foletype_impl_ole_types(mrb_state *mrb, mrb_value self);
static mrb_value foletype_source_ole_types(mrb_state *mrb, mrb_value self);
static mrb_value foletype_default_event_sources(mrb_state *mrb, mrb_value self);
static mrb_value foletype_default_ole_types(mrb_state *mrb, mrb_value self);
static mrb_value foletype_inspect(mrb_state *mrb, mrb_value self);

static const mrb_data_type oletype_datatype = {
    "win32ole_type",
    oletype_free
};

/*
 * Document-class: WIN32OLE_TYPE
 *
 *   <code>WIN32OLE_TYPE</code> objects represent OLE type libarary information.
 */

static void
oletype_free(mrb_state *mrb, void *ptr)
{
    struct oletypedata *poletype = ptr;
    OLE_FREE(poletype->pTypeInfo);
    mrb_free(mrb, poletype);
}

ITypeInfo *itypeinfo(mrb_state *mrb, mrb_value self)
{
    struct oletypedata *ptype;
    Data_Get_Struct(mrb, self, &oletype_datatype, ptype);
    return ptype->pTypeInfo;
}

mrb_value
ole_type_from_itypeinfo(mrb_state *mrb, ITypeInfo *pTypeInfo)
{
    ITypeLib *pTypeLib;
    mrb_value type = mrb_nil_value();
    HRESULT hr;
    unsigned int index;
    BSTR bstr;

    hr = pTypeInfo->lpVtbl->GetContainingTypeLib( pTypeInfo, &pTypeLib, &index );
    if(FAILED(hr)) {
        return mrb_nil_value();
    }
    hr = pTypeLib->lpVtbl->GetDocumentation( pTypeLib, index,
                                             &bstr, NULL, NULL, NULL);
    OLE_RELEASE(pTypeLib);
    if (FAILED(hr)) {
        return mrb_nil_value();
    }
    type = create_win32ole_type(mrb, pTypeInfo, WC2VSTR(mrb, bstr));
    return type;
}


/*
 *   call-seq:
 *      WIN32OLE_TYPE.ole_classes(typelib)
 *
 *   Returns array of WIN32OLE_TYPE objects defined by the <i>typelib</i> type library.
 *   This method will be OBSOLETE. Use WIN32OLE_TYPELIB.new(typelib).ole_classes instead.
 */
static mrb_value
foletype_s_ole_classes(mrb_state *mrb, mrb_value self)
{
    mrb_value typelib;
    mrb_value obj;

    mrb_get_args(mrb, "o", &typelib);

    /*
    rb_warn("%s is obsolete; use %s instead.",
            "WIN32OLE_TYPE.ole_classes",
            "WIN32OLE_TYPELIB.new(typelib).ole_types");
    */
    obj = mrb_funcall(mrb, mrb_obj_value(C_WIN32OLE_TYPELIB), "new", 1, typelib);
    return mrb_funcall(mrb, obj, "ole_types", 0);
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE.typelibs
 *
 *  Returns array of type libraries.
 *  This method will be OBSOLETE. Use WIN32OLE_TYPELIB.typelibs.collect{|t| t.name} instead.
 *
 */
static mrb_value
foletype_s_typelibs(mrb_state *mrb, mrb_value self)
{
    /*
    rb_warn("%s is obsolete. use %s instead.",
            "WIN32OLE_TYPE.typelibs",
            "WIN32OLE_TYPELIB.typelibs.collect{t|t.name}");
    */
    mrb_value typelibs = mrb_funcall(mrb, mrb_obj_value(C_WIN32OLE_TYPELIB), "typelibs", 0);
    int len = RARRAY_LEN(typelibs);
    int i;
    mrb_value typelibnames = mrb_ary_new_capa(mrb, len);
    for (i = 0; i < len; i++) {
        mrb_value val = RARRAY_PTR(typelibs)[i];
        mrb_value name = mrb_funcall(mrb, val, "name", 0);
        mrb_ary_push(mrb, typelibnames, name);
    }
    return typelibnames;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE.progids
 *
 *  Returns array of ProgID.
 */
static mrb_value
foletype_s_progids(mrb_state *mrb, mrb_value self)
{
    HKEY hclsids, hclsid;
    DWORD i;
    LONG err;
    mrb_value clsid;
    mrb_value v = mrb_str_new_lit(mrb, "");
    mrb_value progids = mrb_ary_new(mrb);

    err = reg_open_key(HKEY_CLASSES_ROOT, "CLSID", &hclsids);
    if(err != ERROR_SUCCESS) {
        return progids;
    }
    for(i = 0; ; i++) {
        clsid = reg_enum_key(mrb, hclsids, i);
        if (mrb_nil_p(clsid))
            break;
        err = reg_open_vkey(mrb, hclsids, clsid, &hclsid);
        if (err != ERROR_SUCCESS)
            continue;
        if (!mrb_nil_p(v = reg_get_val2(mrb, hclsid, "ProgID")))
            mrb_ary_push(mrb, progids, v);
        if (!mrb_nil_p(v = reg_get_val2(mrb, hclsid, "VersionIndependentProgID")))
            mrb_ary_push(mrb, progids, v);
        RegCloseKey(hclsid);
    }
    RegCloseKey(hclsids);
    return progids;
}

static mrb_value
oletype_set_member(mrb_state *mrb, mrb_value self, ITypeInfo *pTypeInfo, mrb_value name)
{
    struct oletypedata *ptype;
    Data_Get_Struct(mrb, self, &oletype_datatype, ptype);
    mrb_iv_set(mrb, self, mrb_intern_lit(mrb, "name"), name);
    ptype->pTypeInfo = pTypeInfo;
    OLE_ADDREF(pTypeInfo);
    return self;
}

static struct oletypedata *
oletypedata_alloc(mrb_state *mrb)
{
    struct oletypedata *poletype = ALLOC(struct oletypedata);
    ole_initialize(mrb);
    poletype->pTypeInfo = NULL;
    return poletype;
}

static mrb_value
foletype_s_allocate(mrb_state *mrb, struct RClass *klass)
{
    mrb_value obj;
    struct oletypedata *poletype = oletypedata_alloc(mrb);
    obj = mrb_obj_value(Data_Wrap_Struct(mrb, klass, &oletype_datatype, poletype));
    return obj;
}

mrb_value
create_win32ole_type(mrb_state *mrb, ITypeInfo *pTypeInfo, mrb_value name)
{
    mrb_value obj = foletype_s_allocate(mrb, C_WIN32OLE_TYPE);
    oletype_set_member(mrb, obj, pTypeInfo, name);
    return obj;
}

static mrb_value
oleclass_from_typelib(mrb_state *mrb, mrb_value self, ITypeLib *pTypeLib, mrb_value oleclass)
{

    long count;
    int i;
    HRESULT hr;
    BSTR bstr;
    mrb_value typelib;
    ITypeInfo *pTypeInfo;

    mrb_value found = mrb_false_value();

    count = pTypeLib->lpVtbl->GetTypeInfoCount(pTypeLib);
    for (i = 0; i < count && !mrb_bool(found); i++) {
        hr = pTypeLib->lpVtbl->GetTypeInfo(pTypeLib, i, &pTypeInfo);
        if (FAILED(hr))
            continue;
        hr = pTypeLib->lpVtbl->GetDocumentation(pTypeLib, i,
                                                &bstr, NULL, NULL, NULL);
        if (FAILED(hr))
            continue;
        typelib = WC2VSTR(mrb, bstr);
        if (mrb_str_cmp(mrb, oleclass, typelib) == 0) {
            oletype_set_member(mrb, self, pTypeInfo, typelib);
            found = mrb_true_value();
        }
        OLE_RELEASE(pTypeInfo);
    }
    return found;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE.new(typelib, ole_class) -> WIN32OLE_TYPE object
 *
 *  Returns a new WIN32OLE_TYPE object.
 *  The first argument <i>typelib</i> specifies OLE type library name.
 *  The second argument specifies OLE class name.
 *
 *      WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Application')
 *          # => WIN32OLE_TYPE object of Application class of Excel.
 */
static mrb_value
foletype_initialize(mrb_state *mrb, mrb_value self)
{
    mrb_value typelib;
    mrb_value oleclass;
    mrb_value file;
    BSTR pbuf;
    ITypeLib *pTypeLib;
    HRESULT hr;
    struct oletypedata *poletype;

    poletype = (struct oletypedata *)DATA_PTR(self);
    if (poletype) {
        mrb_free(mrb, poletype);
    }
    mrb_data_init(self, NULL, &oletype_datatype);

    mrb_get_args(mrb, "SS", &typelib, &oleclass);

    poletype = oletypedata_alloc(mrb);
    mrb_data_init(self, poletype, &oletype_datatype);

    file = typelib_file(mrb, typelib);
    if (mrb_nil_p(file)) {
        file = typelib;
    }
    pbuf = ole_vstr2wc(mrb, file);
    hr = LoadTypeLibEx(pbuf, REGKIND_NONE, &pTypeLib);
    if (FAILED(hr))
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to LoadTypeLibEx");
    SysFreeString(pbuf);
    if (!mrb_bool(oleclass_from_typelib(mrb, self, pTypeLib, oleclass))) {
        OLE_RELEASE(pTypeLib);
        mrb_raisef(mrb, E_WIN32OLE_RUNTIME_ERROR, "not found `%S` in `%S`",
                 oleclass, typelib);
    }
    OLE_RELEASE(pTypeLib);

    return self;
}

/*
 * call-seq:
 *    WIN32OLE_TYPE#name #=> OLE type name
 *
 * Returns OLE type name.
 *    tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Application')
 *    puts tobj.name  # => Application
 */
static mrb_value
foletype_name(mrb_state *mrb, mrb_value self)
{
    return mrb_iv_get(mrb, self, mrb_intern_lit(mrb, "name"));
}

static mrb_value
ole_ole_type(mrb_state *mrb, ITypeInfo *pTypeInfo)
{
    HRESULT hr;
    TYPEATTR *pTypeAttr;
    mrb_value type = mrb_nil_value();
    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if(FAILED(hr)){
        return type;
    }
    switch(pTypeAttr->typekind) {
    case TKIND_ENUM:
        type = mrb_str_new_lit(mrb, "Enum");
        break;
    case TKIND_RECORD:
        type = mrb_str_new_lit(mrb, "Record");
        break;
    case TKIND_MODULE:
        type = mrb_str_new_lit(mrb, "Module");
        break;
    case TKIND_INTERFACE:
        type = mrb_str_new_lit(mrb, "Interface");
        break;
    case TKIND_DISPATCH:
        type = mrb_str_new_lit(mrb, "Dispatch");
        break;
    case TKIND_COCLASS:
        type = mrb_str_new_lit(mrb, "Class");
        break;
    case TKIND_ALIAS:
        type = mrb_str_new_lit(mrb, "Alias");
        break;
    case TKIND_UNION:
        type = mrb_str_new_lit(mrb, "Union");
        break;
    case TKIND_MAX:
        type = mrb_str_new_lit(mrb, "Max");
        break;
    default:
        type = mrb_nil_value();
        break;
    }
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
    return type;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE#ole_type #=> OLE type string.
 *
 *  returns type of OLE class.
 *    tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Application')
 *    puts tobj.ole_type  # => Class
 */
static mrb_value
foletype_ole_type(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_ole_type(mrb, pTypeInfo);
}

static mrb_value
ole_type_guid(mrb_state *mrb, ITypeInfo *pTypeInfo)
{
    HRESULT hr;
    TYPEATTR *pTypeAttr;
    int len;
    OLECHAR bstr[80];
    mrb_value guid = mrb_nil_value();
    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if (FAILED(hr))
        return guid;
    len = StringFromGUID2(&pTypeAttr->guid, bstr, sizeof(bstr)/sizeof(OLECHAR));
    if (len > 3) {
        guid = ole_wc2vstr(mrb, bstr, FALSE);
    }
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
    return guid;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE#guid  #=> GUID
 *
 *  Returns GUID.
 *    tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Application')
 *    puts tobj.guid  # => {00024500-0000-0000-C000-000000000046}
 */
static mrb_value
foletype_guid(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_type_guid(mrb, pTypeInfo);
}

static mrb_value
ole_type_progid(mrb_state *mrb, ITypeInfo *pTypeInfo)
{
    HRESULT hr;
    TYPEATTR *pTypeAttr;
    OLECHAR *pbuf;
    mrb_value progid = mrb_nil_value();
    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if (FAILED(hr))
        return progid;
    hr = ProgIDFromCLSID(&pTypeAttr->guid, &pbuf);
    if (SUCCEEDED(hr)) {
        progid = ole_wc2vstr(mrb, pbuf, FALSE);
        CoTaskMemFree(pbuf);
    }
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
    return progid;
}

/*
 * call-seq:
 *    WIN32OLE_TYPE#progid  #=> ProgID
 *
 * Returns ProgID if it exists. If not found, then returns nil.
 *    tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Application')
 *    puts tobj.progid  # =>   Excel.Application.9
 */
static mrb_value
foletype_progid(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_type_progid(mrb, pTypeInfo);
}


static mrb_value
ole_type_visible(mrb_state *mrb, ITypeInfo *pTypeInfo)
{
    HRESULT hr;
    TYPEATTR *pTypeAttr;
    mrb_value visible;
    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if (FAILED(hr))
        return mrb_true_value();
    if (pTypeAttr->wTypeFlags & (TYPEFLAG_FHIDDEN | TYPEFLAG_FRESTRICTED)) {
        visible = mrb_false_value();
    } else {
        visible = mrb_true_value();
    }
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
    return visible;
}

/*
 *  call-seq:
 *    WIN32OLE_TYPE#visible?  #=> true or false
 *
 *  Returns true if the OLE class is public.
 *    tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Application')
 *    puts tobj.visible  # => true
 */
static mrb_value
foletype_visible(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_type_visible(mrb, pTypeInfo);
}

static mrb_value
ole_type_major_version(mrb_state *mrb, ITypeInfo *pTypeInfo)
{
    mrb_value ver;
    TYPEATTR *pTypeAttr;
    HRESULT hr;
    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if (FAILED(hr))
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to GetTypeAttr");
    ver = INT2FIX(pTypeAttr->wMajorVerNum);
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
    return ver;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE#major_version
 *
 *  Returns major version.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Word 10.0 Object Library', 'Documents')
 *     puts tobj.major_version # => 8
 */
static mrb_value
foletype_major_version(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_type_major_version(mrb, pTypeInfo);
}

static mrb_value
ole_type_minor_version(mrb_state *mrb, ITypeInfo *pTypeInfo)
{
    mrb_value ver;
    TYPEATTR *pTypeAttr;
    HRESULT hr;
    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if (FAILED(hr))
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to GetTypeAttr");
    ver = mrb_fixnum_value(pTypeAttr->wMinorVerNum);
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
    return ver;
}

/*
 *  call-seq:
 *    WIN32OLE_TYPE#minor_version #=> OLE minor version
 *
 *  Returns minor version.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Word 10.0 Object Library', 'Documents')
 *     puts tobj.minor_version # => 2
 */
static mrb_value
foletype_minor_version(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_type_minor_version(mrb, pTypeInfo);
}

static mrb_value
ole_type_typekind(mrb_state *mrb, ITypeInfo *pTypeInfo)
{
    mrb_value typekind;
    TYPEATTR *pTypeAttr;
    HRESULT hr;
    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if (FAILED(hr))
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to GetTypeAttr");
    typekind = INT2FIX(pTypeAttr->typekind);
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
    return typekind;
}

/*
 *  call-seq:
 *    WIN32OLE_TYPE#typekind #=> number of type.
 *
 *  Returns number which represents type.
 *    tobj = WIN32OLE_TYPE.new('Microsoft Word 10.0 Object Library', 'Documents')
 *    puts tobj.typekind # => 4
 *
 */
static mrb_value
foletype_typekind(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_type_typekind(mrb, pTypeInfo);
}

static mrb_value
ole_type_helpstring(mrb_state *mrb, ITypeInfo *pTypeInfo)
{
    HRESULT hr;
    BSTR bhelpstr;
    hr = ole_docinfo_from_type(mrb, pTypeInfo, NULL, &bhelpstr, NULL, NULL);
    if(FAILED(hr)) {
        return mrb_nil_value();
    }
    return WC2VSTR(mrb, bhelpstr);
}

/*
 *  call-seq:
 *    WIN32OLE_TYPE#helpstring #=> help string.
 *
 *  Returns help string.
 *    tobj = WIN32OLE_TYPE.new('Microsoft Internet Controls', 'IWebBrowser')
 *    puts tobj.helpstring # => Web Browser interface
 */
static mrb_value
foletype_helpstring(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_type_helpstring(mrb, pTypeInfo);
}

static mrb_value
ole_type_src_type(mrb_state *mrb, ITypeInfo *pTypeInfo)
{
    HRESULT hr;
    TYPEATTR *pTypeAttr;
    mrb_value alias = mrb_nil_value();
    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if (FAILED(hr))
        return alias;
    if(pTypeAttr->typekind != TKIND_ALIAS) {
        OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
        return alias;
    }
    alias = ole_typedesc2val(mrb, pTypeInfo, &(pTypeAttr->tdescAlias), mrb_nil_value());
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
    return alias;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE#src_type #=> OLE source class
 *
 *  Returns source class when the OLE class is 'Alias'.
 *     tobj =  WIN32OLE_TYPE.new('Microsoft Office 9.0 Object Library', 'MsoRGBType')
 *     puts tobj.src_type # => I4
 *
 */
static mrb_value
foletype_src_type(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_type_src_type(mrb, pTypeInfo);
}

static mrb_value
ole_type_helpfile(mrb_state *mrb, ITypeInfo *pTypeInfo)
{
    HRESULT hr;
    BSTR bhelpfile;
    hr = ole_docinfo_from_type(mrb, pTypeInfo, NULL, NULL, NULL, &bhelpfile);
    if(FAILED(hr)) {
        return mrb_nil_value();
    }
    return WC2VSTR(mrb, bhelpfile);
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE#helpfile
 *
 *  Returns helpfile path. If helpfile is not found, then returns nil.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Worksheet')
 *     puts tobj.helpfile # => C:\...\VBAXL9.CHM
 *
 */
static mrb_value
foletype_helpfile(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_type_helpfile(mrb, pTypeInfo);
}

static mrb_value
ole_type_helpcontext(mrb_state *mrb, ITypeInfo *pTypeInfo)
{
    HRESULT hr;
    DWORD helpcontext;
    hr = ole_docinfo_from_type(mrb, pTypeInfo, NULL, NULL,
                               &helpcontext, NULL);
    if(FAILED(hr))
        return mrb_nil_value();
    return INT2FIX(helpcontext);
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE#helpcontext
 *
 *  Returns helpcontext. If helpcontext is not found, then returns nil.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Worksheet')
 *     puts tobj.helpfile # => 131185
 */
static mrb_value
foletype_helpcontext(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_type_helpcontext(mrb, pTypeInfo);
}

static mrb_value
ole_variables(mrb_state *mrb, ITypeInfo *pTypeInfo)
{
    HRESULT hr;
    TYPEATTR *pTypeAttr;
    WORD i;
    UINT len;
    BSTR bstr;
    VARDESC *pVarDesc;
    mrb_value var;
    mrb_value variables = mrb_ary_new(mrb);
    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to GetTypeAttr");
    }

    for(i = 0; i < pTypeAttr->cVars; i++) {
        hr = pTypeInfo->lpVtbl->GetVarDesc(pTypeInfo, i, &pVarDesc);
        if(FAILED(hr))
            continue;
        len = 0;
        hr = pTypeInfo->lpVtbl->GetNames(pTypeInfo, pVarDesc->memid, &bstr,
                                         1, &len);
        if(FAILED(hr) || len == 0 || !bstr)
            continue;

        var = create_win32ole_variable(mrb, pTypeInfo, i, WC2VSTR(mrb, bstr));
        mrb_ary_push(mrb, variables, var);

        pTypeInfo->lpVtbl->ReleaseVarDesc(pTypeInfo, pVarDesc);
        pVarDesc = NULL;
    }
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
    return variables;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE#variables
 *
 *  Returns array of WIN32OLE_VARIABLE objects which represent variables
 *  defined in OLE class.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'XlSheetType')
 *     vars = tobj.variables
 *     vars.each do |v|
 *       puts "#{v.name} = #{v.value}"
 *     end
 *
 *     The result of above sample script is follows:
 *       xlChart = -4109
 *       xlDialogSheet = -4116
 *       xlExcel4IntlMacroSheet = 4
 *       xlExcel4MacroSheet = 3
 *       xlWorksheet = -4167
 *
 */
static mrb_value
foletype_variables(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_variables(mrb, pTypeInfo);
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE#ole_methods # the array of WIN32OLE_METHOD objects.
 *
 *  Returns array of WIN32OLE_METHOD objects which represent OLE method defined in
 *  OLE type library.
 *    tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Worksheet')
 *    methods = tobj.ole_methods.collect{|m|
 *      m.name
 *    }
 *    # => ['Activate', 'Copy', 'Delete',....]
 */
static mrb_value
foletype_methods(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_methods_from_typeinfo(mrb, pTypeInfo, INVOKE_FUNC | INVOKE_PROPERTYGET | INVOKE_PROPERTYPUT | INVOKE_PROPERTYPUTREF);
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE#ole_typelib
 *
 *  Returns the WIN32OLE_TYPELIB object which is including the WIN32OLE_TYPE
 *  object. If it is not found, then returns nil.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Worksheet')
 *     puts tobj.ole_typelib # => 'Microsoft Excel 9.0 Object Library'
 */
static mrb_value
foletype_ole_typelib(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_typelib_from_itypeinfo(mrb, pTypeInfo);
}

static mrb_value
ole_type_impl_ole_types(mrb_state *mrb, ITypeInfo *pTypeInfo, int implflags)
{
    HRESULT hr;
    ITypeInfo *pRefTypeInfo;
    HREFTYPE href;
    WORD i;
    mrb_value type;
    TYPEATTR *pTypeAttr;
    int flags;

    mrb_value types = mrb_ary_new(mrb);
    hr = OLE_GET_TYPEATTR(pTypeInfo, &pTypeAttr);
    if (FAILED(hr)) {
        return types;
    }
    for (i = 0; i < pTypeAttr->cImplTypes; i++) {
        hr = pTypeInfo->lpVtbl->GetImplTypeFlags(pTypeInfo, i, &flags);
        if (FAILED(hr))
            continue;

        hr = pTypeInfo->lpVtbl->GetRefTypeOfImplType(pTypeInfo, i, &href);
        if (FAILED(hr))
            continue;
        hr = pTypeInfo->lpVtbl->GetRefTypeInfo(pTypeInfo, href, &pRefTypeInfo);
        if (FAILED(hr))
            continue;

        if ((flags & implflags) == implflags) {
            type = ole_type_from_itypeinfo(mrb, pRefTypeInfo);
            if (!mrb_nil_p(type)) {
                mrb_ary_push(mrb, types, type);
            }
        }

        OLE_RELEASE(pRefTypeInfo);
    }
    OLE_RELEASE_TYPEATTR(pTypeInfo, pTypeAttr);
    return types;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE#implemented_ole_types
 *
 *  Returns the array of WIN32OLE_TYPE object which is implemented by the WIN32OLE_TYPE
 *  object.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Worksheet')
 *     p tobj.implemented_ole_types # => [_Worksheet, DocEvents]
 */
static mrb_value
foletype_impl_ole_types(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_type_impl_ole_types(mrb, pTypeInfo, 0);
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE#source_ole_types
 *
 *  Returns the array of WIN32OLE_TYPE object which is implemented by the WIN32OLE_TYPE
 *  object and having IMPLTYPEFLAG_FSOURCE.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Internet Controls', "InternetExplorer")
 *     p tobj.source_ole_types
 *     # => [#<WIN32OLE_TYPE:DWebBrowserEvents2>, #<WIN32OLE_TYPE:DWebBrowserEvents>]
 */
static mrb_value
foletype_source_ole_types(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_type_impl_ole_types(mrb, pTypeInfo, IMPLTYPEFLAG_FSOURCE);
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE#default_event_sources
 *
 *  Returns the array of WIN32OLE_TYPE object which is implemented by the WIN32OLE_TYPE
 *  object and having IMPLTYPEFLAG_FSOURCE and IMPLTYPEFLAG_FDEFAULT.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Internet Controls', "InternetExplorer")
 *     p tobj.default_event_sources  # => [#<WIN32OLE_TYPE:DWebBrowserEvents2>]
 */
static mrb_value
foletype_default_event_sources(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_type_impl_ole_types(mrb, pTypeInfo, IMPLTYPEFLAG_FSOURCE|IMPLTYPEFLAG_FDEFAULT);
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE#default_ole_types
 *
 *  Returns the array of WIN32OLE_TYPE object which is implemented by the WIN32OLE_TYPE
 *  object and having IMPLTYPEFLAG_FDEFAULT.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Internet Controls', "InternetExplorer")
 *     p tobj.default_ole_types
 *     # => [#<WIN32OLE_TYPE:IWebBrowser2>, #<WIN32OLE_TYPE:DWebBrowserEvents2>]
 */
static mrb_value
foletype_default_ole_types(mrb_state *mrb, mrb_value self)
{
    ITypeInfo *pTypeInfo = itypeinfo(mrb, self);
    return ole_type_impl_ole_types(mrb, pTypeInfo, IMPLTYPEFLAG_FDEFAULT);
}

/*
 *  call-seq:
 *     WIN32OLE_TYPE#inspect -> String
 *
 *  Returns the type name with class name.
 *
 *     ie = WIN32OLE.new('InternetExplorer.Application')
 *     ie.ole_type.inspect => #<WIN32OLE_TYPE:IWebBrowser2>
 */
static mrb_value
foletype_inspect(mrb_state *mrb, mrb_value self)
{
    return default_inspect(mrb, self, "WIN32OLE_TYPE");
}

void Init_win32ole_type(mrb_state *mrb)
{
    struct RClass *cWIN32OLE_TYPE = mrb_define_class(mrb, "WIN32OLE_TYPE", mrb->object_class);
    mrb_define_class_method(mrb, cWIN32OLE_TYPE, "ole_classes", foletype_s_ole_classes, MRB_ARGS_REQ(1));
    mrb_define_class_method(mrb, cWIN32OLE_TYPE, "typelibs", foletype_s_typelibs, MRB_ARGS_NONE());
    mrb_define_class_method(mrb, cWIN32OLE_TYPE, "progids", foletype_s_progids, MRB_ARGS_NONE());
    MRB_SET_INSTANCE_TT(cWIN32OLE_TYPE, MRB_TT_DATA);
    mrb_define_method(mrb, cWIN32OLE_TYPE, "initialize", foletype_initialize, MRB_ARGS_REQ(2));
    mrb_define_method(mrb, cWIN32OLE_TYPE, "name", foletype_name, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "ole_type", foletype_ole_type, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "guid", foletype_guid, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "progid", foletype_progid, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "visible?", foletype_visible, MRB_ARGS_NONE());
    mrb_define_alias(mrb, cWIN32OLE_TYPE, "to_s", "name");
    mrb_define_method(mrb, cWIN32OLE_TYPE, "major_version", foletype_major_version, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "minor_version", foletype_minor_version, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "typekind", foletype_typekind, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "helpstring", foletype_helpstring, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "src_type", foletype_src_type, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "helpfile", foletype_helpfile, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "helpcontext", foletype_helpcontext, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "variables", foletype_variables, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "ole_methods", foletype_methods, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "ole_typelib", foletype_ole_typelib, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "implemented_ole_types", foletype_impl_ole_types, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "source_ole_types", foletype_source_ole_types, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "default_event_sources", foletype_default_event_sources, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "default_ole_types", foletype_default_ole_types, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPE, "inspect", foletype_inspect, MRB_ARGS_NONE());
}
