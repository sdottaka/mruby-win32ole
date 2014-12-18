#include "win32ole.h"

struct oletypelibdata {
    ITypeLib *pTypeLib;
};

static mrb_value reg_get_typelib_file_path(mrb_state *mrb, HKEY hkey);
static mrb_value oletypelib_path(mrb_state *mrb, mrb_value guid, mrb_value version);
static HRESULT oletypelib_from_guid(mrb_state *mrb, mrb_value guid, mrb_value version, ITypeLib **ppTypeLib);
static mrb_value foletypelib_s_typelibs(mrb_state *mrb, mrb_value self);
static mrb_value oletypelib_set_member(mrb_state *mrb, mrb_value self, ITypeLib *pTypeLib);
static void oletypelib_free(mrb_state *mrb, void *ptr);
static mrb_value foletypelib_s_allocate(mrb_state *mrb, struct RClass *klass);
static mrb_value oletypelib_search_registry(mrb_state *mrb, mrb_value self, mrb_value typelib);
static void oletypelib_get_libattr(mrb_state *mrb, ITypeLib *pTypeLib, TLIBATTR **ppTLibAttr);
static mrb_value oletypelib_search_registry2(mrb_state *mrb, mrb_value self, mrb_value typelib, mrb_value major_ver, mrb_value minor_ver);
static mrb_value foletypelib_initialize(mrb_state *mrb, mrb_value self);
static mrb_value foletypelib_guid(mrb_state *mrb, mrb_value self);
static mrb_value foletypelib_name(mrb_state *mrb, mrb_value self);
static mrb_value make_version_str(mrb_state *mrb, mrb_value major, mrb_value minor);
static mrb_value foletypelib_version(mrb_state *mrb, mrb_value self);
static mrb_value foletypelib_major_version(mrb_state *mrb, mrb_value self);
static mrb_value foletypelib_minor_version(mrb_state *mrb, mrb_value self);
static mrb_value foletypelib_path(mrb_state *mrb, mrb_value self);
static mrb_value foletypelib_visible(mrb_state *mrb, mrb_value self);
static mrb_value foletypelib_library_name(mrb_state *mrb, mrb_value self);
static mrb_value ole_types_from_typelib(mrb_state *mrb, ITypeLib *pTypeLib, mrb_value classes);
static mrb_value typelib_file_from_typelib(mrb_state *mrb, mrb_value ole);
static mrb_value typelib_file_from_clsid(mrb_state *mrb, mrb_value ole);
static mrb_value foletypelib_ole_types(mrb_state *mrb, mrb_value self);
static mrb_value foletypelib_inspect(mrb_state *mrb, mrb_value self);

static const mrb_data_type oletypelib_datatype = {
    "win32ole_typelib",
    oletypelib_free
};

static mrb_value
reg_get_typelib_file_path(mrb_state *mrb, HKEY hkey)
{
    mrb_value path = mrb_nil_value();
    path = reg_get_val2(mrb, hkey, "win64");
    if (!mrb_nil_p(path)) {
        return path;
    }
    path = reg_get_val2(mrb, hkey, "win32");
    if (!mrb_nil_p(path)) {
        return path;
    }
    path = reg_get_val2(mrb, hkey, "win16");
    return path;
}

static mrb_value
oletypelib_path(mrb_state *mrb, mrb_value guid, mrb_value version)
{
    int k;
    LONG err;
    HKEY hkey;
    HKEY hlang;
    mrb_value lang;
    mrb_value path = mrb_nil_value();
    int ai;
    mrb_value key = mrb_str_new_lit(mrb, "TypeLib\\");
    mrb_str_concat(mrb, key, guid);
    mrb_str_cat_lit(mrb, key, "\\");
    mrb_str_concat(mrb, key, version);

    err = reg_open_vkey(mrb, HKEY_CLASSES_ROOT, key, &hkey);
    if (err != ERROR_SUCCESS) {
        return mrb_nil_value();
    }
    ai = mrb_gc_arena_save(mrb);
    for(k = 0; mrb_nil_p(path); k++) {
        lang = reg_enum_key(mrb, hkey, k);
        if (mrb_nil_p(lang)) {
            mrb_gc_arena_restore(mrb, ai);
            break;
        }
        err = reg_open_vkey(mrb, hkey, lang, &hlang);
        mrb_gc_arena_restore(mrb, ai);
        if (err == ERROR_SUCCESS) {
            path = reg_get_typelib_file_path(mrb, hlang);
            RegCloseKey(hlang);
        }
    }
    RegCloseKey(hkey);
    return path;
}

static HRESULT
oletypelib_from_guid(mrb_state *mrb, mrb_value guid, mrb_value version, ITypeLib **ppTypeLib)
{
    mrb_value path;
    OLECHAR *pBuf;
    HRESULT hr;
    path = oletypelib_path(mrb, guid, version);
    if (mrb_nil_p(path)) {
        return E_UNEXPECTED;
    }
    pBuf = ole_vstr2wc(mrb, path);
    hr = LoadTypeLibEx(pBuf, REGKIND_NONE, ppTypeLib);
    SysFreeString(pBuf);
    return hr;
}

ITypeLib *
itypelib(mrb_state *mrb, mrb_value self)
{
    struct oletypelibdata *ptlib;
    Data_Get_Struct(mrb, self, &oletypelib_datatype, ptlib);
    return ptlib->pTypeLib;
}

mrb_value
ole_typelib_from_itypeinfo(mrb_state *mrb, ITypeInfo *pTypeInfo)
{
    HRESULT hr;
    ITypeLib *pTypeLib;
    unsigned int index;
    mrb_value retval = mrb_nil_value();

    hr = pTypeInfo->lpVtbl->GetContainingTypeLib(pTypeInfo, &pTypeLib, &index);
    if(FAILED(hr)) {
        return mrb_nil_value();
    }
    retval = create_win32ole_typelib(mrb, pTypeLib);
    return retval;
}

/*
 * Document-class: WIN32OLE_TYPELIB
 *
 *   <code>WIN32OLE_TYPELIB</code> objects represent OLE tyblib information.
 */

/*
 *  call-seq:
 *
 *     WIN32OLE_TYPELIB.typelibs
 *
 *  Returns the array of WIN32OLE_TYPELIB object.
 *
 *     tlibs = WIN32OLE_TYPELIB.typelibs
 *
 */
static mrb_value
foletypelib_s_typelibs(mrb_state *mrb, mrb_value self)
{
    HKEY htypelib, hguid;
    DWORD i, j;
    LONG err;
    mrb_value guid;
    mrb_value version;
    mrb_value name = mrb_nil_value();
    mrb_value typelibs = mrb_ary_new(mrb);
    mrb_value typelib = mrb_nil_value();
    HRESULT hr;
    ITypeLib *pTypeLib;
    int ai;

    err = reg_open_key(HKEY_CLASSES_ROOT, "TypeLib", &htypelib);
    if(err != ERROR_SUCCESS) {
        return typelibs;
    }
    ai = mrb_gc_arena_save(mrb);
    for(i = 0; ; i++) {
        guid = reg_enum_key(mrb, htypelib, i);
        if (mrb_nil_p(guid)) {
            mrb_gc_arena_restore(mrb, ai);
            break;
        }
        err = reg_open_vkey(mrb, htypelib, guid, &hguid);
        if (err != ERROR_SUCCESS) {
            mrb_gc_arena_restore(mrb, ai);
            continue;
        }
        for(j = 0; ; j++) {
            version = reg_enum_key(mrb, hguid, j);
            if (mrb_nil_p(version)) {
                mrb_gc_arena_restore(mrb, ai);
                break;
            }
            if ( !mrb_nil_p((name = reg_get_val2(mrb, hguid, ole_obj_to_cstr(mrb, version)))) ) {
                hr = oletypelib_from_guid(mrb, guid, version, &pTypeLib);
                if (SUCCEEDED(hr)) {
                    typelib = create_win32ole_typelib(mrb, pTypeLib);
                    mrb_ary_push(mrb, typelibs, typelib);
                }
            }
            mrb_gc_arena_restore(mrb, ai);
        }
        mrb_gc_arena_restore(mrb, ai);
        RegCloseKey(hguid);
    }
    RegCloseKey(htypelib);
    return typelibs;
}

static mrb_value
oletypelib_set_member(mrb_state *mrb, mrb_value self, ITypeLib *pTypeLib)
{
    struct oletypelibdata *ptlib;
    Data_Get_Struct(mrb, self, &oletypelib_datatype, ptlib);
    ptlib->pTypeLib = pTypeLib;
    return self;
}

static void
oletypelib_free(mrb_state *mrb, void *ptr)
{
    struct oletypelibdata *poletypelib = ptr;
    if (!ptr) return;
    OLE_FREE(poletypelib->pTypeLib);
    mrb_free(mrb, poletypelib);
}

static struct oletypelibdata *
oletypelibdata_alloc(mrb_state *mrb)
{
    struct oletypelibdata *poletypelib = ALLOC(struct oletypelibdata);
    ole_initialize(mrb);
    poletypelib->pTypeLib = NULL;
    return poletypelib;
}

static mrb_value
foletypelib_s_allocate(mrb_state *mrb, struct RClass *klass)
{
    mrb_value obj;
    struct oletypelibdata *poletypelib = oletypelibdata_alloc(mrb);
    obj = mrb_obj_value(Data_Wrap_Struct(mrb, klass, &oletypelib_datatype, poletypelib));
    return obj;
}

mrb_value
create_win32ole_typelib(mrb_state *mrb, ITypeLib *pTypeLib)
{
    mrb_value obj = foletypelib_s_allocate(mrb, C_WIN32OLE_TYPELIB);
    oletypelib_set_member(mrb, obj, pTypeLib);
    return obj;
}

static mrb_value
oletypelib_search_registry(mrb_state *mrb, mrb_value self, mrb_value typelib)
{
    HKEY htypelib, hguid, hversion;
    DWORD i, j;
    LONG err;
    mrb_value found = mrb_false_value();
    mrb_value tlib;
    mrb_value guid;
    mrb_value ver;
    HRESULT hr;
    ITypeLib *pTypeLib;
    int ai;

    err = reg_open_key(HKEY_CLASSES_ROOT, "TypeLib", &htypelib);
    if(err != ERROR_SUCCESS) {
        return mrb_false_value();
    }
    ai = mrb_gc_arena_save(mrb);
    for(i = 0; !mrb_bool(found); i++) {
        guid = reg_enum_key(mrb, htypelib, i);
        if (mrb_nil_p(guid)) {
            mrb_gc_arena_restore(mrb, ai);
            break;
        }
        err = reg_open_vkey(mrb, htypelib, guid, &hguid);
        if (err != ERROR_SUCCESS) {
            mrb_gc_arena_restore(mrb, ai);
            continue;
        }
        for(j = 0; !mrb_bool(found); j++) {
            ver = reg_enum_key(mrb, hguid, j);
            if (mrb_nil_p(ver)) {
                mrb_gc_arena_restore(mrb, ai);
                break;
            }
            err = reg_open_vkey(mrb, hguid, ver, &hversion);
            if (err != ERROR_SUCCESS) {
                mrb_gc_arena_restore(mrb, ai);
                continue;
            }
            tlib = reg_get_val(mrb, hversion, NULL);
            if (mrb_nil_p(tlib)) {
                mrb_gc_arena_restore(mrb, ai);
                RegCloseKey(hversion);
                continue;
            }
            if (mrb_str_cmp(mrb, typelib, tlib) == 0) {
                hr = oletypelib_from_guid(mrb, guid, ver, &pTypeLib);
                if (SUCCEEDED(hr)) {
                    oletypelib_set_member(mrb, self, pTypeLib);
                    found = mrb_true_value();
                }
            }
            mrb_gc_arena_restore(mrb, ai);
            RegCloseKey(hversion);
        }
        mrb_gc_arena_restore(mrb, ai);
        RegCloseKey(hguid);
    }
    RegCloseKey(htypelib);
    return  found;
}

static void
oletypelib_get_libattr(mrb_state *mrb, ITypeLib *pTypeLib, TLIBATTR **ppTLibAttr)
{
    HRESULT hr;
    hr = pTypeLib->lpVtbl->GetLibAttr(pTypeLib, ppTLibAttr);
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR,
		  "failed to get library attribute(TLIBATTR) from ITypeLib");
    }
}

static mrb_value
oletypelib_search_registry2(mrb_state *mrb, mrb_value self, mrb_value typelibname, mrb_value major_ver, mrb_value minor_ver)
{
    HKEY htypelib, hguid, hversion;
    double fver;
    DWORD j;
    LONG err;
    mrb_value found = mrb_false_value();
    mrb_value tlib;
    mrb_value ver;
    mrb_value version_str;
    mrb_value version = mrb_nil_value();
    mrb_value typelib = mrb_nil_value();
    HRESULT hr;
    ITypeLib *pTypeLib;
    int ai;
    mrb_value guid = typelibname;

    ai = mrb_gc_arena_save(mrb);
    version_str = make_version_str(mrb, major_ver, minor_ver);

    err = reg_open_key(HKEY_CLASSES_ROOT, "TypeLib", &htypelib);
    if(err != ERROR_SUCCESS) {
        mrb_gc_arena_restore(mrb, ai);
        return mrb_false_value();
    }
    err = reg_open_vkey(mrb, htypelib, guid, &hguid);
    if (err != ERROR_SUCCESS) {
        mrb_gc_arena_restore(mrb, ai);
        RegCloseKey(htypelib);
        return mrb_false_value();
    }
    if (!mrb_nil_p(version_str)) {
        err = reg_open_vkey(mrb, hguid, version_str, &hversion);
        if (err == ERROR_SUCCESS) {
            tlib = reg_get_val(mrb, hversion, NULL);
            if (!mrb_nil_p(tlib)) {
                typelib = tlib;
                version = version_str;
            }
        }
        RegCloseKey(hversion);
    } else {
        fver = 0.0;
        for(j = 0; ;j++) {
            ver = reg_enum_key(mrb, hguid, j);
            if (mrb_nil_p(ver))
                break;
            err = reg_open_vkey(mrb, hguid, ver, &hversion);
            if (err != ERROR_SUCCESS)
                continue;
            tlib = reg_get_val(mrb, hversion, NULL);
            if (mrb_nil_p(tlib)) {
                RegCloseKey(hversion);
                continue;
            }
            if (fver < atof(ole_obj_to_cstr(mrb, ver))) {
                fver = atof(ole_obj_to_cstr(mrb, ver));
                version = ver;
                typelib = tlib;
            }
            RegCloseKey(hversion);
        }
    }
    RegCloseKey(hguid);
    RegCloseKey(htypelib);
    if (!mrb_nil_p(typelib)) {
        hr = oletypelib_from_guid(mrb, guid, version, &pTypeLib);
        if (SUCCEEDED(hr)) {
            found = mrb_true_value();
            oletypelib_set_member(mrb, self, pTypeLib);
        }
    }
    mrb_gc_arena_restore(mrb, ai);
    return found;
}


/*
 * call-seq:
 *    WIN32OLE_TYPELIB.new(typelib [, version1, version2]) -> WIN32OLE_TYPELIB object
 *
 * Returns a new WIN32OLE_TYPELIB object.
 *
 * The first argument <i>typelib</i>  specifies OLE type library name or GUID or
 * OLE library file.
 * The second argument is major version or version of the type library.
 * The third argument is minor version.
 * The second argument and third argument are optional.
 * If the first argument is type library name, then the second and third argument
 * are ignored.
 *
 *     tlib1 = WIN32OLE_TYPELIB.new('Microsoft Excel 9.0 Object Library')
 *     tlib2 = WIN32OLE_TYPELIB.new('{00020813-0000-0000-C000-000000000046}')
 *     tlib3 = WIN32OLE_TYPELIB.new('{00020813-0000-0000-C000-000000000046}', 1.3)
 *     tlib4 = WIN32OLE_TYPELIB.new('{00020813-0000-0000-C000-000000000046}', 1, 3)
 *     tlib5 = WIN32OLE_TYPELIB.new("C:\\WINNT\\SYSTEM32\\SHELL32.DLL")
 *     puts tlib1.name  # -> 'Microsoft Excel 9.0 Object Library'
 *     puts tlib2.name  # -> 'Microsoft Excel 9.0 Object Library'
 *     puts tlib3.name  # -> 'Microsoft Excel 9.0 Object Library'
 *     puts tlib4.name  # -> 'Microsoft Excel 9.0 Object Library'
 *     puts tlib5.name  # -> 'Microsoft Shell Controls And Automation'
 *
 */
static mrb_value
foletypelib_initialize(mrb_state *mrb, mrb_value self)
{
    mrb_int argc;
    mrb_value argv;
    mrb_value major_ver = mrb_nil_value();
    mrb_value minor_ver = mrb_nil_value();
    mrb_value found = mrb_false_value();
    mrb_value typelib = mrb_nil_value();
    OLECHAR * pbuf;
    ITypeLib *pTypeLib;
    HRESULT hr = S_OK;

    struct oletypelibdata *ptlib = (struct oletypelibdata *)DATA_PTR(self);
    if (ptlib) {
        mrb_free(mrb, ptlib);
    }
    mrb_data_init(self, NULL, &oletypelib_datatype);

    mrb_get_args(mrb, "*", &argv, &argc);
    if (argc < 1 || argc > 3) {
        mrb_raisef(mrb, E_ARGUMENT_ERROR, "wrong number of arguments (%S for 1..3)", INT2FIX(argc));
    }
    mrb_get_args(mrb, "S|oo", &typelib, &major_ver, &minor_ver);

    ptlib = oletypelibdata_alloc(mrb);
    mrb_data_init(self, ptlib, &oletypelib_datatype);

    found = oletypelib_search_registry(mrb, self, typelib);
    if (!mrb_bool(found)) {
        found = oletypelib_search_registry2(mrb, self, typelib, major_ver, minor_ver);
    }
    if (!mrb_bool(found)) {
        pbuf = ole_vstr2wc(mrb, typelib);
        hr = LoadTypeLibEx(pbuf, REGKIND_NONE, &pTypeLib);
        SysFreeString(pbuf);
        if (SUCCEEDED(hr)) {
            found = mrb_true_value();
            oletypelib_set_member(mrb, self, pTypeLib);
        }
    }

    if (!mrb_bool(found)) {
        mrb_raisef(mrb, E_WIN32OLE_RUNTIME_ERROR, "not found type library `%S`",
                   typelib);
    }
    return self;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPELIB#guid -> The guid string.
 *
 *  Returns guid string which specifies type library.
 *
 *     tlib = WIN32OLE_TYPELIB.new('Microsoft Excel 9.0 Object Library')
 *     guid = tlib.guid # -> '{00020813-0000-0000-C000-000000000046}'
 */
static mrb_value
foletypelib_guid(mrb_state *mrb, mrb_value self)
{
    ITypeLib *pTypeLib;
    OLECHAR bstr[80];
    mrb_value guid = mrb_nil_value();
    int len;
    TLIBATTR *pTLibAttr;

    pTypeLib = itypelib(mrb, self);
    oletypelib_get_libattr(mrb, pTypeLib, &pTLibAttr);
    len = StringFromGUID2(&pTLibAttr->guid, bstr, sizeof(bstr)/sizeof(OLECHAR));
    if (len > 3) {
        guid = ole_wc2vstr(mrb, bstr, FALSE);
    }
    pTypeLib->lpVtbl->ReleaseTLibAttr(pTypeLib, pTLibAttr);
    return guid;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPELIB#name -> The type library name
 *
 *  Returns the type library name.
 *
 *     tlib = WIN32OLE_TYPELIB.new('Microsoft Excel 9.0 Object Library')
 *     name = tlib.name # -> 'Microsoft Excel 9.0 Object Library'
 */
static mrb_value
foletypelib_name(mrb_state *mrb, mrb_value self)
{
    ITypeLib *pTypeLib;
    HRESULT hr;
    BSTR bstr;
    mrb_value name;
    pTypeLib = itypelib(mrb, self);
    hr = pTypeLib->lpVtbl->GetDocumentation(pTypeLib, -1,
                                            NULL, &bstr, NULL, NULL);

    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to get name from ITypeLib");
    }
    name = WC2VSTR(mrb, bstr);
    return name;
}

static mrb_value
make_version_str(mrb_state *mrb, mrb_value major, mrb_value minor)
{
    mrb_value version_str = mrb_nil_value();
    mrb_value minor_str = mrb_nil_value();
    if (mrb_nil_p(major)) {
        return mrb_nil_value();
    }
    version_str = mrb_str_to_str(mrb, major);
    if (!mrb_nil_p(minor)) {
        minor_str = mrb_str_to_str(mrb, minor);
        mrb_str_cat_lit(mrb, version_str, ".");
        mrb_str_append(mrb, version_str, minor_str);
    }
    return version_str;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPELIB#version -> The type library version String object.
 *
 *  Returns the type library version.
 *
 *     tlib = WIN32OLE_TYPELIB.new('Microsoft Excel 9.0 Object Library')
 *     puts tlib.version #-> "1.3"
 */
static mrb_value
foletypelib_version(mrb_state *mrb, mrb_value self)
{
    TLIBATTR *pTLibAttr;
    ITypeLib *pTypeLib;
    mrb_value version;

    pTypeLib = itypelib(mrb, self);
    oletypelib_get_libattr(mrb, pTypeLib, &pTLibAttr);
    version = mrb_format(mrb, "%S.%S", mrb_fixnum_value(pTLibAttr->wMajorVerNum), mrb_fixnum_value(pTLibAttr->wMinorVerNum));
    pTypeLib->lpVtbl->ReleaseTLibAttr(pTypeLib, pTLibAttr);
    return version;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPELIB#major_version -> The type library major version.
 *
 *  Returns the type library major version.
 *
 *     tlib = WIN32OLE_TYPELIB.new('Microsoft Excel 9.0 Object Library')
 *     puts tlib.major_version # -> 1
 */
static mrb_value
foletypelib_major_version(mrb_state *mrb, mrb_value self)
{
    TLIBATTR *pTLibAttr;
    mrb_value major;
    ITypeLib *pTypeLib;
    pTypeLib = itypelib(mrb, self);
    oletypelib_get_libattr(mrb, pTypeLib, &pTLibAttr);

    major =  INT2NUM(pTLibAttr->wMajorVerNum);
    pTypeLib->lpVtbl->ReleaseTLibAttr(pTypeLib, pTLibAttr);
    return major;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPELIB#minor_version -> The type library minor version.
 *
 *  Returns the type library minor version.
 *
 *     tlib = WIN32OLE_TYPELIB.new('Microsoft Excel 9.0 Object Library')
 *     puts tlib.minor_version # -> 3
 */
static mrb_value
foletypelib_minor_version(mrb_state *mrb, mrb_value self)
{
    TLIBATTR *pTLibAttr;
    mrb_value minor;
    ITypeLib *pTypeLib;
    pTypeLib = itypelib(mrb, self);
    oletypelib_get_libattr(mrb, pTypeLib, &pTLibAttr);
    minor =  INT2NUM(pTLibAttr->wMinorVerNum);
    pTypeLib->lpVtbl->ReleaseTLibAttr(pTypeLib, pTLibAttr);
    return minor;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPELIB#path -> The type library file path.
 *
 *  Returns the type library file path.
 *
 *     tlib = WIN32OLE_TYPELIB.new('Microsoft Excel 9.0 Object Library')
 *     puts tlib.path #-> 'C:\...\EXCEL9.OLB'
 */
static mrb_value
foletypelib_path(mrb_state *mrb, mrb_value self)
{
    TLIBATTR *pTLibAttr;
    HRESULT hr = S_OK;
    BSTR bstr;
    LCID lcid = cWIN32OLE_lcid;
    mrb_value path;
    ITypeLib *pTypeLib;

    pTypeLib = itypelib(mrb, self);
    oletypelib_get_libattr(mrb, pTypeLib, &pTLibAttr);
    hr = QueryPathOfRegTypeLib(&pTLibAttr->guid,
	                       pTLibAttr->wMajorVerNum,
			       pTLibAttr->wMinorVerNum,
			       lcid,
			       &bstr);
    if (FAILED(hr)) {
	pTypeLib->lpVtbl->ReleaseTLibAttr(pTypeLib, pTLibAttr);
	ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to QueryPathOfRegTypeTypeLib");
    }

    pTypeLib->lpVtbl->ReleaseTLibAttr(pTypeLib, pTLibAttr);
    path = WC2VSTR(mrb, bstr);
    return path;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPELIB#visible?
 *
 *  Returns true if the type library information is not hidden.
 *  If wLibFlags of TLIBATTR is 0 or LIBFLAG_FRESTRICTED or LIBFLAG_FHIDDEN,
 *  the method returns false, otherwise, returns true.
 *  If the method fails to access the TLIBATTR information, then
 *  WIN32OLERuntimeError is raised.
 *
 *     tlib = WIN32OLE_TYPELIB.new('Microsoft Excel 9.0 Object Library')
 *     tlib.visible? # => true
 */
static mrb_value
foletypelib_visible(mrb_state *mrb, mrb_value self)
{
    ITypeLib *pTypeLib = NULL;
    mrb_value visible = mrb_true_value();
    TLIBATTR *pTLibAttr;

    pTypeLib = itypelib(mrb, self);
    oletypelib_get_libattr(mrb, pTypeLib, &pTLibAttr);

    if ((pTLibAttr->wLibFlags == 0) ||
        (pTLibAttr->wLibFlags & LIBFLAG_FRESTRICTED) ||
        (pTLibAttr->wLibFlags & LIBFLAG_FHIDDEN)) {
        visible = mrb_false_value();
    }
    pTypeLib->lpVtbl->ReleaseTLibAttr(pTypeLib, pTLibAttr);
    return visible;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPELIB#library_name
 *
 *  Returns library name.
 *  If the method fails to access library name, WIN32OLERuntimeError is raised.
 *
 *     tlib = WIN32OLE_TYPELIB.new('Microsoft Excel 9.0 Object Library')
 *     tlib.library_name # => Excel
 */
static mrb_value
foletypelib_library_name(mrb_state *mrb, mrb_value self)
{
    HRESULT hr;
    ITypeLib *pTypeLib = NULL;
    mrb_value libname = mrb_nil_value();
    BSTR bstr;

    pTypeLib = itypelib(mrb, self);
    hr = pTypeLib->lpVtbl->GetDocumentation(pTypeLib, -1,
                                            &bstr, NULL, NULL, NULL);
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to get library name");
    }
    libname = WC2VSTR(mrb, bstr);
    return libname;
}

static mrb_value
ole_types_from_typelib(mrb_state *mrb, ITypeLib *pTypeLib, mrb_value classes)
{
    long count;
    int i;
    HRESULT hr;
    BSTR bstr;
    ITypeInfo *pTypeInfo;
    mrb_value type;
    int ai;

    count = pTypeLib->lpVtbl->GetTypeInfoCount(pTypeLib);
    for (i = 0; i < count; i++) {
        hr = pTypeLib->lpVtbl->GetDocumentation(pTypeLib, i,
                                                &bstr, NULL, NULL, NULL);
        if (FAILED(hr))
            continue;

        hr = pTypeLib->lpVtbl->GetTypeInfo(pTypeLib, i, &pTypeInfo);
        if (FAILED(hr))
            continue;

        ai = mrb_gc_arena_save(mrb);
        type = create_win32ole_type(mrb, pTypeInfo, WC2VSTR(mrb, bstr));

        mrb_ary_push(mrb, classes, type);
        OLE_RELEASE(pTypeInfo);
        mrb_gc_arena_restore(mrb, ai);
    }
    return classes;
}

static mrb_value
typelib_file_from_typelib(mrb_state *mrb, mrb_value ole)
{
    HKEY htypelib, hclsid, hversion, hlang;
    double fver;
    DWORD i, j, k;
    LONG err;
    BOOL found = FALSE;
    mrb_value typelib;
    mrb_value file = mrb_nil_value();
    mrb_value clsid;
    mrb_value ver;
    mrb_value lang;
    int ai;

    err = reg_open_key(HKEY_CLASSES_ROOT, "TypeLib", &htypelib);
    if(err != ERROR_SUCCESS) {
        return mrb_nil_value();
    }
    ai = mrb_gc_arena_save(mrb);
    for(i = 0; !found; i++) {
        mrb_gc_arena_restore(mrb, ai);
        clsid = reg_enum_key(mrb, htypelib, i);
        if (mrb_nil_p(clsid))
            break;
        err = reg_open_vkey(mrb, htypelib, clsid, &hclsid);
        if (err != ERROR_SUCCESS)
            continue;
        fver = 0;
        for(j = 0; !found; j++) {
            mrb_gc_arena_restore(mrb, ai);
            ver = reg_enum_key(mrb, hclsid, j);
            if (mrb_nil_p(ver))
                break;
            err = reg_open_vkey(mrb, hclsid, ver, &hversion);
			if (err != ERROR_SUCCESS || fver > atof(ole_obj_to_cstr(mrb, ver)))
                continue;
            fver = atof(ole_obj_to_cstr(mrb, ver));
            typelib = reg_get_val(mrb, hversion, NULL);
            if (mrb_nil_p(typelib))
                continue;
            if (mrb_str_cmp(mrb, typelib, ole) == 0) {
                for(k = 0; !found; k++) {
                    mrb_gc_arena_restore(mrb, ai);
                    lang = reg_enum_key(mrb, hversion, k);
                    if (mrb_nil_p(lang))
                        break;
                    err = reg_open_vkey(mrb, hversion, lang, &hlang);
                    if (err == ERROR_SUCCESS) {
                        if (!mrb_nil_p((file = reg_get_typelib_file_path(mrb, hlang))))
                            found = TRUE;
                        RegCloseKey(hlang);
                    }
                }
            }
            RegCloseKey(hversion);
        }
        RegCloseKey(hclsid);
    }
    mrb_gc_arena_restore(mrb, ai);
    RegCloseKey(htypelib);
    return  file;
}

static mrb_value
typelib_file_from_clsid(mrb_state *mrb, mrb_value ole)
{
    HKEY hroot, hclsid;
    LONG err;
    mrb_value typelib;
    char path[MAX_PATH + 1];

    err = reg_open_key(HKEY_CLASSES_ROOT, "CLSID", &hroot);
    if (err != ERROR_SUCCESS) {
        return mrb_nil_value();
    }
    err = reg_open_key(hroot, ole_obj_to_cstr(mrb, ole), &hclsid);
    if (err != ERROR_SUCCESS) {
        RegCloseKey(hroot);
        return mrb_nil_value();
    }
    typelib = reg_get_val2(mrb, hclsid, "InprocServer32");
    RegCloseKey(hroot);
    RegCloseKey(hclsid);
    if (!mrb_nil_p(typelib)) {
        ExpandEnvironmentStringsA(ole_obj_to_cstr(mrb, typelib), path, sizeof(path));
        path[MAX_PATH] = '\0';
        typelib = mrb_str_new_cstr(mrb, path);
    }
    return typelib;
}

mrb_value
typelib_file(mrb_state *mrb, mrb_value ole)
{
    mrb_value file = typelib_file_from_clsid(mrb, ole);
    if (!mrb_nil_p(file)) {
        return file;
    }
    return typelib_file_from_typelib(mrb, ole);
}


/*
 *  call-seq:
 *     WIN32OLE_TYPELIB#ole_types -> The array of WIN32OLE_TYPE object included the type library.
 *
 *  Returns the type library file path.
 *
 *     tlib = WIN32OLE_TYPELIB.new('Microsoft Excel 9.0 Object Library')
 *     classes = tlib.ole_types.collect{|k| k.name} # -> ['AddIn', 'AddIns' ...]
 */
static mrb_value
foletypelib_ole_types(mrb_state *mrb, mrb_value self)
{
    ITypeLib *pTypeLib = NULL;
    mrb_value classes = mrb_ary_new(mrb);
    pTypeLib = itypelib(mrb, self);
    ole_types_from_typelib(mrb, pTypeLib, classes);
    return classes;
}

/*
 *  call-seq:
 *     WIN32OLE_TYPELIB#inspect -> String
 *
 *  Returns the type library name with class name.
 *
 *     tlib = WIN32OLE_TYPELIB.new('Microsoft Excel 9.0 Object Library')
 *     tlib.inspect # => "<#WIN32OLE_TYPELIB:Microsoft Excel 9.0 Object Library>"
 */
static mrb_value
foletypelib_inspect(mrb_state *mrb, mrb_value self)
{
    return default_inspect(mrb, self, "WIN32OLE_TYPELIB");
}

void
Init_win32ole_typelib(mrb_state *mrb)
{
    struct RClass *cWIN32OLE_TYPELIB = mrb_define_class(mrb, "WIN32OLE_TYPELIB", mrb->object_class);
    mrb_define_class_method(mrb, cWIN32OLE_TYPELIB, "typelibs", foletypelib_s_typelibs, MRB_ARGS_NONE());
    MRB_SET_INSTANCE_TT(cWIN32OLE_TYPELIB, MRB_TT_DATA);
    mrb_define_method(mrb, cWIN32OLE_TYPELIB, "initialize", foletypelib_initialize, MRB_ARGS_ANY());
    mrb_define_method(mrb, cWIN32OLE_TYPELIB, "guid", foletypelib_guid, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPELIB, "name", foletypelib_name, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPELIB, "version", foletypelib_version, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPELIB, "major_version", foletypelib_major_version, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPELIB, "minor_version", foletypelib_minor_version, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPELIB, "path", foletypelib_path, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPELIB, "ole_types", foletypelib_ole_types, MRB_ARGS_NONE());
    mrb_define_alias(mrb, cWIN32OLE_TYPELIB, "ole_classes", "ole_types");
    mrb_define_method(mrb, cWIN32OLE_TYPELIB, "visible?", foletypelib_visible, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_TYPELIB, "library_name", foletypelib_library_name, MRB_ARGS_NONE());
    mrb_define_alias(mrb, cWIN32OLE_TYPELIB, "to_s", "name");
    mrb_define_method(mrb, cWIN32OLE_TYPELIB, "inspect", foletypelib_inspect, MRB_ARGS_NONE());
}
