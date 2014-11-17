#include "win32ole.h"

struct olerecorddata {
    IRecordInfo *pri;
    void *pdata;
};

static HRESULT recordinfo_from_itypelib(mrb_state *mrb, ITypeLib *pTypeLib, mrb_value name, IRecordInfo **ppri);
static int hash2olerec(mrb_state *mrb, mrb_value key, mrb_value val, mrb_value rec);
static void olerecord_free(mrb_state *mrb, void *ptr);
static mrb_value folerecord_s_allocate(mrb_state *mrb, struct RClass *klass);
static mrb_value folerecord_initialize(mrb_state *mrb, mrb_value self);
static mrb_value folerecord_to_h(mrb_state *mrb, mrb_value self);
static mrb_value folerecord_typename(mrb_state *mrb, mrb_value self);
static mrb_value olerecord_ivar_get(mrb_state *mrb, mrb_value self, mrb_value name);
static mrb_value olerecord_ivar_set(mrb_state *mrb, mrb_value self, mrb_value name, mrb_value val);
static mrb_value folerecord_method_missing(mrb_state *mrb, mrb_value self);
static mrb_value folerecord_ole_instance_variable_get(mrb_state *mrb, mrb_value self);
static mrb_value folerecord_ole_instance_variable_set(mrb_state *mrb, mrb_value self);
static mrb_value folerecord_inspect(mrb_state *mrb, mrb_value self);

static const mrb_data_type olerecord_datatype = {
    "win32ole_record",
    olerecord_free,
};

static HRESULT
recordinfo_from_itypelib(mrb_state *mrb, ITypeLib *pTypeLib, mrb_value name, IRecordInfo **ppri)
{

    unsigned int count;
    unsigned int i;
    ITypeInfo *pTypeInfo;
    HRESULT hr = OLE_E_LAST;
    BSTR bstr;

    count = pTypeLib->lpVtbl->GetTypeInfoCount(pTypeLib);
    for (i = 0; i < count; i++) {
        hr = pTypeLib->lpVtbl->GetDocumentation(pTypeLib, i,
                                                &bstr, NULL, NULL, NULL);
        if (FAILED(hr))
            continue;

        hr = pTypeLib->lpVtbl->GetTypeInfo(pTypeLib, i, &pTypeInfo);
        if (FAILED(hr))
            continue;

        if (mrb_str_cmp(mrb, WC2VSTR(mrb, bstr), name) == 0) {
            hr = GetRecordInfoFromTypeInfo(pTypeInfo, ppri);
            OLE_RELEASE(pTypeInfo);
            return hr;
        }
        OLE_RELEASE(pTypeInfo);
    }
    hr = OLE_E_LAST;
    return hr;
}

static int
hash2olerec(mrb_state *mrb, mrb_value key, mrb_value val, mrb_value rec)
{
    VARIANT var;
    OLECHAR *pbuf;
    struct olerecorddata *prec;
    IRecordInfo *pri;
    HRESULT hr;

    if (!mrb_nil_p(val)) {
        Data_Get_Struct(mrb, rec, &olerecord_datatype, prec);
        pri = prec->pri;
        VariantInit(&var);
        ole_val2variant(mrb, val, &var);
        pbuf = ole_vstr2wc(mrb, key);
        hr = pri->lpVtbl->PutField(pri, INVOKE_PROPERTYPUT, prec->pdata, pbuf, &var);
        SysFreeString(pbuf);
        VariantClear(&var);
        if (FAILED(hr)) {
            ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to putfield of `%s`", mrb_string_value_ptr(mrb, key));
        }
    }
    return 0;
}

void
ole_rec2variant(mrb_state *mrb, mrb_value rec, VARIANT *var)
{
    struct olerecorddata *prec;
    ULONG size = 0;
    IRecordInfo *pri;
    HRESULT hr;
    mrb_value fields;
    mrb_value keys;
    int i, l;
    Data_Get_Struct(mrb, rec, &olerecord_datatype, prec);
    pri = prec->pri;
    if (pri) {
        hr = pri->lpVtbl->GetSize(pri, &size);
        if (FAILED(hr)) {
            ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to get size for allocation of VT_RECORD object");
        }
        if (prec->pdata) {
            free(prec->pdata);
        }
        prec->pdata = ALLOC_N(char, size);
        if (!prec->pdata) {
            mrb_raisef(mrb, E_RUNTIME_ERROR, "failed to memory allocation of %S bytes", mrb_fixnum_value(size));
        }
        hr = pri->lpVtbl->RecordInit(pri, prec->pdata);
        if (FAILED(hr)) {
            ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to initialize VT_RECORD object");
        }
        fields = folerecord_to_h(mrb, rec);
        keys = mrb_hash_keys(mrb, fields);
        l = RARRAY_LEN(keys);
        for (i = 0; i < l; i++) {
            mrb_value key = RARRAY_PTR(keys)[i];
            hash2olerec(mrb, key, mrb_hash_get(mrb, fields, key), rec); 
        }
        V_RECORDINFO(var) = pri;
        V_RECORD(var) = prec->pdata;
        V_VT(var) = VT_RECORD;
    } else {
        mrb_raise(mrb, E_WIN32OLE_RUNTIME_ERROR, "failed to retrieve IRecordInfo interface");
    }
}

void
olerecord_set_ivar(mrb_state *mrb, mrb_value obj, IRecordInfo *pri, void *prec)
{
    HRESULT hr;
    BSTR bstr;
    BSTR *bstrs;
    ULONG count = 0;
    ULONG i;
    mrb_value fields;
    mrb_value val;
    VARIANT var;
    void *pdata = NULL;
    struct olerecorddata *pvar;

    Data_Get_Struct(mrb, obj, &olerecord_datatype, pvar);
    OLE_ADDREF(pri);
    OLE_RELEASE(pvar->pri);
    pvar->pri = pri;

    hr = pri->lpVtbl->GetName(pri, &bstr);
    if (SUCCEEDED(hr)) {
        mrb_iv_set(mrb, obj, mrb_intern_lit(mrb, "typename"), WC2VSTR(mrb, bstr));
    }

    hr = pri->lpVtbl->GetFieldNames(pri, &count, NULL);
    if (FAILED(hr) || count == 0)
        return;
    bstrs = ALLOCA_N(BSTR, count);
    hr = pri->lpVtbl->GetFieldNames(pri, &count, bstrs);
    if (FAILED(hr)) {
        return;
    }

    /* TODO: should use mrb_gc_arena_save()? */
    fields = mrb_hash_new(mrb);
    mrb_iv_set(mrb, obj, mrb_intern_lit(mrb, "fields"), fields);
    for (i = 0; i < count; i++) {
        pdata = NULL;
        VariantInit(&var);
        val = mrb_nil_value();
        if (prec) {
            hr = pri->lpVtbl->GetFieldNoCopy(pri, prec, bstrs[i], &var, &pdata);
            if (SUCCEEDED(hr)) {
                val = ole_variant2val(mrb, &var);
            }
        }
        mrb_hash_set(mrb, fields, WC2VSTR(mrb, bstrs[i]), val);
    }
}

mrb_value
create_win32ole_record(mrb_state *mrb, IRecordInfo *pri, void *prec)
{
    mrb_value obj = folerecord_s_allocate(mrb, C_WIN32OLE_RECORD);
    olerecord_set_ivar(mrb, obj, pri, prec);
    return obj;
}

/*
 * Document-class: WIN32OLE_RECORD
 *
 *   <code>WIN32OLE_RECORD</code> objects represents VT_RECORD OLE variant.
 *   Win32OLE returns WIN32OLE_RECORD object if the result value of invoking
 *   OLE methods.
 *
 *   If COM server in VB.NET ComServer project is the following:
 *
 *     Imports System.Runtime.InteropServices
 *     Public Class ComClass
 *         Public Structure Book
 *             <MarshalAs(UnmanagedType.BStr)> _
 *             Public title As String
 *             Public cost As Integer
 *         End Structure
 *         Public Function getBook() As Book
 *             Dim book As New Book
 *             book.title = "The Ruby Book"
 *             book.cost = 20
 *             Return book
 *         End Function
 *     End Class
 *
 *   then, you can retrieve getBook return value from the following
 *   Ruby script:
 *
 *     require 'win32ole'
 *     obj = WIN32OLE.new('ComServer.ComClass')
 *     book = obj.getBook
 *     book.class # => WIN32OLE_RECORD
 *     book.title # => "The Ruby Book"
 *     book.cost  # => 20
 *
 */

static void
olerecord_free(mrb_state *mrb, void *ptr) {
    struct olerecorddata *pvar = ptr;
    OLE_FREE(pvar->pri);
    if (pvar->pdata) {
        mrb_free(mrb, pvar->pdata);
    }
    mrb_free(mrb, pvar);
}

static struct olerecorddata *olerecorddata_alloc(mrb_state *mrb)
{
    struct olerecorddata *pvar = ALLOC(struct olerecorddata);
    pvar->pri = NULL;
    pvar->pdata = NULL;
    return pvar;
}

static mrb_value
folerecord_s_allocate(mrb_state *mrb, struct RClass *klass) {
    mrb_value obj = mrb_nil_value();
    struct olerecorddata *pvar = olerecorddata_alloc(mrb);
    obj = mrb_obj_value(Data_Wrap_Struct(mrb, klass, &olerecord_datatype, pvar));
    return obj;
}

/*
 * call-seq:
 *    WIN32OLE_RECORD.new(typename, obj) -> WIN32OLE_RECORD object
 *
 * Returns WIN32OLE_RECORD object. The first argument is struct name (String
 * or Symbol).
 * The second parameter obj should be WIN32OLE object or WIN32OLE_TYPELIB object.
 * If COM server in VB.NET ComServer project is the following:
 *
 *   Imports System.Runtime.InteropServices
 *   Public Class ComClass
 *       Public Structure Book
 *           <MarshalAs(UnmanagedType.BStr)> _
 *           Public title As String
 *           Public cost As Integer
 *       End Structure
 *   End Class
 *
 * then, you can create WIN32OLE_RECORD object is as following:
 *
 *   require 'win32ole'
 *   obj = WIN32OLE.new('ComServer.ComClass')
 *   book1 = WIN32OLE_RECORD.new('Book', obj) # => WIN32OLE_RECORD object
 *   tlib = obj.ole_typelib
 *   book2 = WIN32OLE_RECORD.new('Book', tlib) # => WIN32OLE_RECORD object
 *
 */
static mrb_value
folerecord_initialize(mrb_state *mrb, mrb_value self) {
    mrb_value typename, oleobj;
    HRESULT hr;
    ITypeLib *pTypeLib = NULL;
    IRecordInfo *pri = NULL;
    struct olerecorddata *prec;

    prec = (struct olerecorddata *)DATA_PTR(self);
    if (prec) {
        mrb_free(mrb, prec);
    }
    mrb_data_init(self, NULL, &olerecord_datatype);

    mrb_get_args(mrb, "oo", &typename, &oleobj);

    if (!mrb_string_p(typename) && !mrb_symbol_p(typename)) {
        mrb_raise(mrb, E_ARGUMENT_ERROR, "1st argument should be String or Symbol");
    }
    if (mrb_symbol_p(typename)) {
        typename = mrb_sym2str(mrb, mrb_symbol(typename));
    }

    hr = S_OK;
    if(mrb_obj_is_kind_of(mrb, oleobj, C_WIN32OLE)) {
        hr = typelib_from_val(mrb, oleobj, &pTypeLib);
    } else if (mrb_obj_is_kind_of(mrb, oleobj, C_WIN32OLE_TYPELIB)) {
        pTypeLib = itypelib(mrb, oleobj);
        OLE_ADDREF(pTypeLib);
        if (pTypeLib) {
            hr = S_OK;
        } else {
            hr = E_FAIL;
        }
    } else {
        mrb_raise(mrb, E_ARGUMENT_ERROR, "2nd argument should be WIN32OLE object or WIN32OLE_TYPELIB object");
    }

    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "fail to query ITypeLib interface");
    }

    hr = recordinfo_from_itypelib(mrb, pTypeLib, typename, &pri);
    OLE_RELEASE(pTypeLib);
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "fail to query IRecordInfo interface for `%s'", mrb_string_value_ptr(mrb, typename));
    }

    prec = olerecorddata_alloc(mrb);
    mrb_data_init(self, prec, &olerecord_datatype);

    olerecord_set_ivar(mrb, self, pri, NULL);

    return self;
}

/*
 *  call-seq:
 *     WIN32OLE_RECORD#to_h #=> Ruby Hash object.
 *
 *  Returns Ruby Hash object which represents VT_RECORD variable.
 *  The keys of Hash object are member names of VT_RECORD OLE variable and
 *  the values of Hash object are values of VT_RECORD OLE variable.
 *
 *  If COM server in VB.NET ComServer project is the following:
 *
 *     Imports System.Runtime.InteropServices
 *     Public Class ComClass
 *         Public Structure Book
 *             <MarshalAs(UnmanagedType.BStr)> _
 *             Public title As String
 *             Public cost As Integer
 *         End Structure
 *         Public Function getBook() As Book
 *             Dim book As New Book
 *             book.title = "The Ruby Book"
 *             book.cost = 20
 *             Return book
 *         End Function
 *     End Class
 *
 *  then, the result of WIN32OLE_RECORD#to_h is the following:
 *
 *     require 'win32ole'
 *     obj = WIN32OLE.new('ComServer.ComClass')
 *     book = obj.getBook
 *     book.to_h # => {"title"=>"The Ruby Book", "cost"=>20}
 *
 */
static mrb_value
folerecord_to_h(mrb_state *mrb, mrb_value self)
{
    return mrb_iv_get(mrb, self, mrb_intern_lit(mrb, "fields"));
}

/*
 *  call-seq:
 *     WIN32OLE_RECORD#typename #=> String object
 *
 *  Returns the type name of VT_RECORD OLE variable.
 *
 *  If COM server in VB.NET ComServer project is the following:
 *
 *     Imports System.Runtime.InteropServices
 *     Public Class ComClass
 *         Public Structure Book
 *             <MarshalAs(UnmanagedType.BStr)> _
 *             Public title As String
 *             Public cost As Integer
 *         End Structure
 *         Public Function getBook() As Book
 *             Dim book As New Book
 *             book.title = "The Ruby Book"
 *             book.cost = 20
 *             Return book
 *         End Function
 *     End Class
 *
 *  then, the result of WIN32OLE_RECORD#typename is the following:
 *
 *     require 'win32ole'
 *     obj = WIN32OLE.new('ComServer.ComClass')
 *     book = obj.getBook
 *     book.typename # => "Book"
 *
 */
static mrb_value
folerecord_typename(mrb_state *mrb, mrb_value self)
{
    return mrb_iv_get(mrb, self, mrb_intern_lit(mrb, "typename"));
}

static mrb_value
olerecord_ivar_get(mrb_state *mrb, mrb_value self, mrb_value name)
{
    mrb_value fields;
    mrb_value val;
    fields = mrb_iv_get(mrb, self, mrb_intern_lit(mrb, "fields"));
    val = mrb_hash_fetch(mrb, fields, name, mrb_undef_value());
    if (mrb_undef_p(val))
        mrb_raisef(mrb, E_KEY_ERROR, "key not found:%S", name);
    return val;
}

static mrb_value
olerecord_ivar_set(mrb_state *mrb, mrb_value self, mrb_value name, mrb_value val)
{
    long len;
    char *p;
    mrb_value fields;
    len  = RSTRING_LEN(name);
    p = RSTRING_PTR(name);
    if (p[len-1] == '=') {
        name = mrb_str_substr(mrb, name, 0, len-1);
    }
    fields = mrb_iv_get(mrb, self, mrb_intern_lit(mrb, "fields"));
    if (mrb_undef_p(mrb_hash_fetch(mrb, fields, name, mrb_undef_value())))
        mrb_raisef(mrb, E_KEY_ERROR, "key not found:%S", name);
    mrb_hash_set(mrb, fields, name, val);
    return val;
}

/*
 *  call-seq:
 *     WIN32OLE_RECORD#method_missing(name)
 *
 *  Returns value specified by the member name of VT_RECORD OLE variable.
 *  Or sets value specified by the member name of VT_RECORD OLE variable.
 *  If the member name is not correct, KeyError exception is raised.
 *
 *  If COM server in VB.NET ComServer project is the following:
 *
 *     Imports System.Runtime.InteropServices
 *     Public Class ComClass
 *         Public Structure Book
 *             <MarshalAs(UnmanagedType.BStr)> _
 *             Public title As String
 *             Public cost As Integer
 *         End Structure
 *     End Class
 *
 *  Then getting/setting value from Ruby is as the following:
 *
 *     obj = WIN32OLE.new('ComServer.ComClass')
 *     book = WIN32OLE_RECORD.new('Book', obj)
 *     book.title # => nil ( book.method_missing(:title) is invoked. )
 *     book.title = "Ruby" # ( book.method_missing(:title=, "Ruby") is invoked. )
 */
static mrb_value
folerecord_method_missing(mrb_state *mrb, mrb_value self)
{
    int argc;
    mrb_value *argv;
    mrb_value name;
    mrb_get_args(mrb, "*", &argv, &argc);
    if (argc < 1 || argc > 2)
        mrb_raisef(mrb, E_ARGUMENT_ERROR, "wrong number of arguments (%S for 1..2)", mrb_fixnum_value(argc));

    name = mrb_sym2str(mrb, mrb_symbol(argv[0]));

#if SIZEOF_SIZE_T > SIZEOF_LONG
    {
        size_t n = strlen(mrb_string_value_cstr(mrb, name));
        if (n >= LONG_MAX) {
            mrb_raise(mrb, E_RUNTIME_ERROR, "too long member name");
        }
    }
#endif

    if (argc == 1) {
        return olerecord_ivar_get(mrb, self, name);
    } else if (argc == 2) {
        return olerecord_ivar_set(mrb, self, name, argv[1]);
    }
    return mrb_nil_value();
}

/*
 *  call-seq:
 *     WIN32OLE_RECORD#ole_instance_variable_get(name)
 *
 *  Returns value specified by the member name of VT_RECORD OLE object.
 *  If the member name is not correct, KeyError exception is raised.
 *  If you can't access member variable of VT_RECORD OLE object directly,
 *  use this method.
 *
 *  If COM server in VB.NET ComServer project is the following:
 *
 *     Imports System.Runtime.InteropServices
 *     Public Class ComClass
 *         Public Structure ComObject
 *             Public object_id As Ineger
 *         End Structure
 *     End Class
 *
 *  and Ruby Object class has title attribute:
 *
 *  then accessing object_id of ComObject from Ruby is as the following:
 *
 *     srver = WIN32OLE.new('ComServer.ComClass')
 *     obj = WIN32OLE_RECORD.new('ComObject', server)
 *     # obj.object_id returns Ruby Object#object_id
 *     obj.ole_instance_variable_get(:object_id) # => nil
 *
 */
static mrb_value
folerecord_ole_instance_variable_get(mrb_state *mrb, mrb_value self)
{
    mrb_value name;
    mrb_value sname;
    mrb_get_args(mrb, "o", &name);
    if(!mrb_string_p(name) && !mrb_symbol_p(name)) {
        mrb_raise(mrb, E_TYPE_ERROR, "wrong argument type (expected String or Symbol)");
    }
    sname = name;
    if (mrb_symbol_p(name)) {
        sname = mrb_sym2str(mrb, mrb_symbol(name));
    }
    return olerecord_ivar_get(mrb, self, sname);
}

/*
 *  call-seq:
 *     WIN32OLE_RECORD#ole_instance_variable_set(name, val)
 *
 *  Sets value specified by the member name of VT_RECORD OLE object.
 *  If the member name is not correct, KeyError exception is raised.
 *  If you can't set value of member of VT_RECORD OLE object directly,
 *  use this method.
 *
 *  If COM server in VB.NET ComServer project is the following:
 *
 *     Imports System.Runtime.InteropServices
 *     Public Class ComClass
 *         <MarshalAs(UnmanagedType.BStr)> _
 *         Public title As String
 *         Public cost As Integer
 *     End Class
 *
 *  then setting value of the `title' member is as following:
 *
 *     srver = WIN32OLE.new('ComServer.ComClass')
 *     obj = WIN32OLE_RECORD.new('Book', server)
 *     obj.ole_instance_variable_set(:title, "The Ruby Book")
 *
 */
static mrb_value
folerecord_ole_instance_variable_set(mrb_state *mrb, mrb_value self)
{
    mrb_value sname;
    mrb_value name, val;
    mrb_get_args(mrb, "oo", &name, &val);
    if(!mrb_string_p(name) && !mrb_symbol_p(name)) {
        mrb_raise(mrb, E_TYPE_ERROR, "wrong argument type (expected String or Symbol)");
    }
    sname = name;
    if (mrb_symbol_p(name)) {
        sname = mrb_sym2str(mrb, mrb_symbol(name));
    }
    return olerecord_ivar_set(mrb, self, sname, val);
}

/*
 *  call-seq:
 *     WIN32OLE_RECORD#inspect -> String
 *
 *  Returns the OLE struct name and member name and the value of member
 *
 *  If COM server in VB.NET ComServer project is the following:
 *
 *     Imports System.Runtime.InteropServices
 *     Public Class ComClass
 *         <MarshalAs(UnmanagedType.BStr)> _
 *         Public title As String
 *         Public cost As Integer
 *     End Class
 *
 *  then
 *
 *     srver = WIN32OLE.new('ComServer.ComClass')
 *     obj = WIN32OLE_RECORD.new('Book', server)
 *     obj.inspect # => <WIN32OLE_RECORD(ComClass) {"title" => nil, "cost" => nil}>
 *
 */
static mrb_value
folerecord_inspect(mrb_state *mrb, mrb_value self)
{
    mrb_value tname;
    mrb_value field;
    tname = folerecord_typename(mrb, self);
    if (mrb_nil_p(tname)) {
        tname = mrb_inspect(mrb, tname);
    }
    field = mrb_inspect(mrb, folerecord_to_h(mrb, self));
    return mrb_format(mrb, "#<WIN32OLE_RECORD(%S) %S>",
                      tname,
                      field);
}

void
Init_win32ole_record(mrb_state *mrb)
{
    struct RClass *cWIN32OLE_RECORD = mrb_define_class(mrb, "WIN32OLE_RECORD", mrb->object_class);
    MRB_SET_INSTANCE_TT(cWIN32OLE_RECORD, MRB_TT_DATA);
    mrb_define_method(mrb, cWIN32OLE_RECORD, "initialize", folerecord_initialize, MRB_ARGS_REQ(2));
    mrb_define_method(mrb, cWIN32OLE_RECORD, "to_h", folerecord_to_h, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_RECORD, "typename", folerecord_typename, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_RECORD, "method_missing", folerecord_method_missing, MRB_ARGS_ANY());
    mrb_define_method(mrb, cWIN32OLE_RECORD, "ole_instance_variable_get", folerecord_ole_instance_variable_get, MRB_ARGS_REQ(1));
    mrb_define_method(mrb, cWIN32OLE_RECORD, "ole_instance_variable_set", folerecord_ole_instance_variable_set, MRB_ARGS_REQ(2));
    mrb_define_method(mrb, cWIN32OLE_RECORD, "inspect", folerecord_inspect, MRB_ARGS_NONE());
}
