#include "win32ole.h"

struct olevariantdata {
    VARIANT realvar;
    VARIANT var;
};

static void  olevariant_free(mrb_state *mrb, void *ptr);
static void ole_val2olevariantdata(mrb_state *mrb, mrb_value val, VARTYPE vt, struct olevariantdata *pvar);
static void ole_val2variant_err(mrb_state *mrb, mrb_value val, VARIANT *var);
static void ole_set_byref(mrb_state *mrb, VARIANT *realvar, VARIANT *var,  VARTYPE vt);
static struct olevariantdata *olevariantdata_alloc(mrb_state *mrb);
static mrb_value folevariant_s_allocate(mrb_state *mrb, struct RClass *klass);
static mrb_value folevariant_s_array(mrb_state *mrb, mrb_value klass);
static void check_type_val2variant(mrb_state *mrb, mrb_value val);
static mrb_value folevariant_initialize(mrb_state *mrb, mrb_value self);
static LONG *ary2safe_array_index(mrb_state *mrb, int ary_size, mrb_value *ary, SAFEARRAY *psa);
static void unlock_safe_array(mrb_state *mrb, SAFEARRAY *psa);
static SAFEARRAY *get_locked_safe_array(mrb_state *mrb, mrb_value val);
static mrb_value folevariant_ary_aref(mrb_state *mrb, mrb_value self);
static mrb_value folevariant_ary_aset(mrb_state *mrb, mrb_value self);
static mrb_value folevariant_value(mrb_state *mrb, mrb_value self);
static mrb_value folevariant_vartype(mrb_state *mrb, mrb_value self);
static mrb_value folevariant_set_value(mrb_state *mrb, mrb_value self);

static const mrb_data_type olevariant_datatype = {
    "win32ole_variant",
    olevariant_free
};

static void
olevariant_free(mrb_state *mrb, void *ptr)
{
    struct olevariantdata *pvar = ptr;
    if (!ptr) return;
    VariantClear(&(pvar->realvar));
    VariantClear(&(pvar->var));
    mrb_free(mrb, pvar);
}

static void
ole_val2olevariantdata(mrb_state *mrb, mrb_value val, VARTYPE vt, struct olevariantdata *pvar)
{
    HRESULT hr = S_OK;

    if (((vt & ~VT_BYREF) ==  (VT_ARRAY | VT_UI1)) && mrb_string_p(val)) {
        long len = RSTRING_LEN(val);
        void *pdest = NULL;
        SAFEARRAY *p = NULL;
        SAFEARRAY *psa = SafeArrayCreateVector(VT_UI1, 0, len);
        if (!psa) {
            mrb_raise(mrb, E_RUNTIME_ERROR, "fail to SafeArrayCreateVector");
        }
        hr = SafeArrayAccessData(psa, &pdest);
        if (SUCCEEDED(hr)) {
            memcpy(pdest, RSTRING_PTR(val), len);
            SafeArrayUnaccessData(psa);
            V_VT(&(pvar->realvar)) = (vt & ~VT_BYREF);
            p = V_ARRAY(&(pvar->realvar));
            if (p != NULL) {
                SafeArrayDestroy(p);
            }
            V_ARRAY(&(pvar->realvar)) = psa;
            if (vt & VT_BYREF) {
                V_VT(&(pvar->var)) = vt;
                V_ARRAYREF(&(pvar->var)) = &(V_ARRAY(&(pvar->realvar)));
            } else {
                hr = VariantCopy(&(pvar->var), &(pvar->realvar));
            }
        } else {
            if (psa)
                SafeArrayDestroy(psa);
        }
    } else if (vt & VT_ARRAY) {
        if (mrb_nil_p(val)) {
            V_VT(&(pvar->var)) = vt;
            if (vt & VT_BYREF) {
                V_ARRAYREF(&(pvar->var)) = &(V_ARRAY(&(pvar->realvar)));
            }
        } else {
            hr = ole_val_ary2variant_ary(mrb, val, &(pvar->realvar), (VARTYPE)(vt & ~VT_BYREF));
            if (SUCCEEDED(hr)) {
                if (vt & VT_BYREF) {
                    V_VT(&(pvar->var)) = vt;
                    V_ARRAYREF(&(pvar->var)) = &(V_ARRAY(&(pvar->realvar)));
                } else {
                    hr = VariantCopy(&(pvar->var), &(pvar->realvar));
                }
            }
        }
#if (_MSC_VER >= 1300) || defined(__CYGWIN__) || defined(__MINGW32__)
    } else if ( (vt & ~VT_BYREF) == VT_I8 || (vt & ~VT_BYREF) == VT_UI8) {
        ole_val2variant_ex(mrb, val, &(pvar->realvar), (vt & ~VT_BYREF));
        ole_val2variant_ex(mrb, val, &(pvar->var), (vt & ~VT_BYREF));
        V_VT(&(pvar->var)) = vt;
        if (vt & VT_BYREF) {
            ole_set_byref(mrb, &(pvar->realvar), &(pvar->var), vt);
        }
#endif
    } else if ( (vt & ~VT_BYREF) == VT_ERROR) {
        ole_val2variant_err(mrb, val, &(pvar->realvar));
        if (vt & VT_BYREF) {
            ole_set_byref(mrb, &(pvar->realvar), &(pvar->var), vt);
        } else {
            hr = VariantCopy(&(pvar->var), &(pvar->realvar));
        }
    } else {
        if (mrb_nil_p(val)) {
            V_VT(&(pvar->var)) = vt;
            if (vt == (VT_BYREF | VT_VARIANT)) {
                ole_set_byref(mrb, &(pvar->realvar), &(pvar->var), vt);
            } else {
                V_VT(&(pvar->realvar)) = vt & ~VT_BYREF;
                if (vt & VT_BYREF) {
                    ole_set_byref(mrb, &(pvar->realvar), &(pvar->var), vt);
                }
            }
        } else {
            ole_val2variant_ex(mrb, val, &(pvar->realvar), (VARTYPE)(vt & ~VT_BYREF));
            if (vt == (VT_BYREF | VT_VARIANT)) {
                ole_set_byref(mrb, &(pvar->realvar), &(pvar->var), vt);
            } else if (vt & VT_BYREF) {
                if ( (vt & ~VT_BYREF) != V_VT(&(pvar->realvar))) {
                    hr = VariantChangeTypeEx(&(pvar->realvar), &(pvar->realvar),
                            cWIN32OLE_lcid, 0, (VARTYPE)(vt & ~VT_BYREF));
                }
                if (SUCCEEDED(hr)) {
                    ole_set_byref(mrb, &(pvar->realvar), &(pvar->var), vt);
                }
            } else {
                if (vt == V_VT(&(pvar->realvar))) {
                    hr = VariantCopy(&(pvar->var), &(pvar->realvar));
                } else {
                    hr = VariantChangeTypeEx(&(pvar->var), &(pvar->realvar),
                            cWIN32OLE_lcid, 0, vt);
                }
            }
        }
    }
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to change type");
    }
}

static void
ole_val2variant_err(mrb_state *mrb, mrb_value val, VARIANT *var)
{
    mrb_value v = val;
    if (mrb_obj_is_kind_of(mrb, v, C_WIN32OLE_VARIANT)) {
        v = folevariant_value(mrb, v);
    }
    if (!mrb_fixnum_p(v) && /* FIXME: TYPE(v) != T_BIGNUM &&*/ !mrb_nil_p(v)) {
        mrb_raisef(mrb, E_WIN32OLE_RUNTIME_ERROR, "failed to convert VT_ERROR VARIANT:`%S'", mrb_inspect(mrb, v));
    }
    V_VT(var) = VT_ERROR;
    if (!mrb_nil_p(v)) {
        V_ERROR(var) = NUM2LONG(val);
    } else {
        V_ERROR(var) = 0;
    }
}

static void
ole_set_byref(mrb_state *mrb, VARIANT *realvar, VARIANT *var,  VARTYPE vt)
{
    V_VT(var) = vt;
    if (vt == (VT_VARIANT|VT_BYREF)) {
        V_VARIANTREF(var) = realvar;
    } else {
        if (V_VT(realvar) != (vt & ~VT_BYREF)) {
            mrb_raise(mrb, E_WIN32OLE_RUNTIME_ERROR, "variant type mismatch");
        }
        switch(vt & ~VT_BYREF) {
        case VT_I1:
            V_I1REF(var) = &V_I1(realvar);
            break;
        case VT_UI1:
            V_UI1REF(var) = &V_UI1(realvar);
            break;
        case VT_I2:
            V_I2REF(var) = &V_I2(realvar);
            break;
        case VT_UI2:
            V_UI2REF(var) = &V_UI2(realvar);
            break;
        case VT_I4:
            V_I4REF(var) = &V_I4(realvar);
            break;
        case VT_UI4:
            V_UI4REF(var) = &V_UI4(realvar);
            break;
        case VT_R4:
            V_R4REF(var) = &V_R4(realvar);
            break;
        case VT_R8:
            V_R8REF(var) = &V_R8(realvar);
            break;

#if (_MSC_VER >= 1300) || defined(__CYGWIN__) || defined(__MINGW32__)
#ifdef V_I8REF
        case VT_I8:
            V_I8REF(var) = &V_I8(realvar);
            break;
#endif
#ifdef V_UI8REF
        case VT_UI8:
            V_UI8REF(var) = &V_UI8(realvar);
            break;
#endif
#endif
        case VT_INT:
            V_INTREF(var) = &V_INT(realvar);
            break;

        case VT_UINT:
            V_UINTREF(var) = &V_UINT(realvar);
            break;

        case VT_CY:
            V_CYREF(var) = &V_CY(realvar);
            break;
        case VT_DATE:
            V_DATEREF(var) = &V_DATE(realvar);
            break;
        case VT_BSTR:
            V_BSTRREF(var) = &V_BSTR(realvar);
            break;
        case VT_DISPATCH:
            V_DISPATCHREF(var) = &V_DISPATCH(realvar);
            break;
        case VT_ERROR:
            V_ERRORREF(var) = &V_ERROR(realvar);
            break;
        case VT_BOOL:
            V_BOOLREF(var) = &V_BOOL(realvar);
            break;
        case VT_UNKNOWN:
            V_UNKNOWNREF(var) = &V_UNKNOWN(realvar);
            break;
        case VT_ARRAY:
            V_ARRAYREF(var) = &V_ARRAY(realvar);
            break;
        default:
            mrb_raisef(mrb, E_WIN32OLE_RUNTIME_ERROR, "unknown type specified(setting BYREF):%S", INT2FIX(vt));
            break;
        }
    }
}

static struct olevariantdata *olevariantdata_alloc(mrb_state *mrb)
{
    struct olevariantdata *pvar = ALLOC(struct olevariantdata);
    ole_initialize(mrb);
    VariantInit(&(pvar->var));
    VariantInit(&(pvar->realvar));
    return pvar;
}

static mrb_value
folevariant_s_allocate(mrb_state *mrb, struct RClass *klass)
{
    mrb_value obj;
    struct olevariantdata *pvar = olevariantdata_alloc(mrb);
    obj = mrb_obj_value(Data_Wrap_Struct(mrb, klass, &olevariant_datatype, pvar));
    return obj;
}

/*
 *  call-seq:
 *     WIN32OLE_VARIANT.array(ary, vt)
 *
 *  Returns Ruby object wrapping OLE variant whose variant type is VT_ARRAY.
 *  The first argument should be Array object which specifies dimensions
 *  and each size of dimensions of OLE array.
 *  The second argument specifies variant type of the element of OLE array.
 *
 *  The following create 2 dimensions OLE array. The first dimensions size
 *  is 3, and the second is 4.
 *
 *     ole_ary = WIN32OLE_VARIANT.array([3,4], VT_I4)
 *     ruby_ary = ole_ary.value # => [[0, 0, 0, 0], [0, 0, 0, 0], [0, 0, 0, 0]]
 *
 */
static mrb_value
folevariant_s_array(mrb_state *mrb, mrb_value klass)
{
    mrb_value elems;
    mrb_value obj = mrb_nil_value();
    VARTYPE vt;
    struct olevariantdata *pvar;
    SAFEARRAYBOUND *psab = NULL;
    SAFEARRAY *psa = NULL;
    UINT dim = 0;
    UINT i = 0;

    mrb_get_args(mrb, "Ai", &elems, &vt);

    ole_initialize(mrb);

    vt = (vt | VT_ARRAY);
    obj = folevariant_s_allocate(mrb, mrb_class_ptr(klass));

    Data_Get_Struct(mrb, obj, &olevariant_datatype, pvar);
    dim = RARRAY_LEN(elems);

    psab = ALLOC_N(SAFEARRAYBOUND, dim);

    if(!psab) {
        mrb_raise(mrb, E_RUNTIME_ERROR, "memory allocation error");
    }

    for (i = 0; i < dim; i++) {
        psab[i].cElements = FIX2INT(mrb_ary_entry(elems, i));
        psab[i].lLbound = 0;
    }

    psa = SafeArrayCreate((VARTYPE)(vt & VT_TYPEMASK), dim, psab);
    if (psa == NULL) {
        if (psab) free(psab);
        mrb_raise(mrb, E_RUNTIME_ERROR, "memory allocation error(SafeArrayCreate)");
    }

    V_VT(&(pvar->var)) = vt;
    if (vt & VT_BYREF) {
        V_VT(&(pvar->realvar)) = (vt & ~VT_BYREF);
        V_ARRAY(&(pvar->realvar)) = psa;
        V_ARRAYREF(&(pvar->var)) = &(V_ARRAY(&(pvar->realvar)));
    } else {
        V_ARRAY(&(pvar->var)) = psa;
    }
    if (psab) free(psab);
    return obj;
}

static void
check_type_val2variant(mrb_state *mrb, mrb_value val)
{
    mrb_value elem;
    int len = 0;
    int i = 0;
    if(!mrb_obj_is_kind_of(mrb, val, C_WIN32OLE) &&
       !mrb_obj_is_kind_of(mrb, val, C_WIN32OLE_VARIANT) &&
       !mrb_obj_is_kind_of(mrb, val, mrb_class_get(mrb, "Time"))) {
        switch (mrb_type(val)) {
        case MRB_TT_ARRAY:
            len = RARRAY_LEN(val);
            for(i = 0; i < len; i++) {
                elem = mrb_ary_entry(val, i);
                check_type_val2variant(mrb, elem);
            }
            break;
        case MRB_TT_STRING:
        case MRB_TT_FIXNUM:
        /* FIXME: case MRB_TT_BIGNUM: */
        case MRB_TT_FLOAT:
        case MRB_TT_TRUE:
        case MRB_TT_FALSE:
            break;
        default:
            if (mrb_nil_p(val))
                break;
            mrb_raisef(mrb, E_TYPE_ERROR, "can not convert WIN32OLE_VARIANT from type %S",
                       mrb_str_new_cstr(mrb, mrb_obj_classname(mrb, val)));
        }
    }
}

/*
 * Document-class: WIN32OLE_VARIANT
 *
 *   <code>WIN32OLE_VARIANT</code> objects represents OLE variant.
 *
 *   Win32OLE converts Ruby object into OLE variant automatically when
 *   invoking OLE methods. If OLE method requires the argument which is
 *   different from the variant by automatic conversion of Win32OLE, you
 *   can convert the specfied variant type by using WIN32OLE_VARIANT class.
 *
 *     param = WIN32OLE_VARIANT.new(10, WIN32OLE::VARIANT::VT_R4)
 *     oleobj.method(param)
 *
 *   WIN32OLE_VARIANT does not support VT_RECORD variant. Use WIN32OLE_RECORD
 *   class instead of WIN32OLE_VARIANT if the VT_RECORD variant is needed.
 */

/*
 *  call-seq:
 *     WIN32OLE_VARIANT.new(val, vartype) #=> WIN32OLE_VARIANT object.
 *
 *  Returns Ruby object wrapping OLE variant.
 *  The first argument specifies Ruby object to convert OLE variant variable.
 *  The second argument specifies VARIANT type.
 *  In some situation, you need the WIN32OLE_VARIANT object to pass OLE method
 *
 *     shell = WIN32OLE.new("Shell.Application")
 *     folder = shell.NameSpace("C:\\Windows")
 *     item = folder.ParseName("tmp.txt")
 *     # You can't use Ruby String object to call FolderItem.InvokeVerb.
 *     # Instead, you have to use WIN32OLE_VARIANT object to call the method.
 *     shortcut = WIN32OLE_VARIANT.new("Create Shortcut(\&S)")
 *     item.invokeVerb(shortcut)
 *
 */
static mrb_value
folevariant_initialize(mrb_state *mrb, mrb_value self)
{
    mrb_int argc;
    mrb_value *argv;
    int len = 0;
    VARIANT var;
    mrb_value val;
    mrb_value vvt;
    VARTYPE vt;
    struct olevariantdata *pvar;

    pvar = (struct olevariantdata *)DATA_PTR(self);
    if (pvar) {
        mrb_free(mrb, pvar);
    }
    mrb_data_init(self, NULL, &olevariant_datatype);

    mrb_get_args(mrb, "*", &argv, &argc);

    len = argc;
    if (len < 1 || len > 3) {
        mrb_raisef(mrb, E_ARGUMENT_ERROR, "wrong number of arguments (%S for 1..3)", INT2FIX(len));
    }
    VariantInit(&var);
    val = argv[0];

    check_type_val2variant(mrb, val);

    pvar = olevariantdata_alloc(mrb);
    mrb_data_init(self, pvar, &olevariant_datatype);

    Data_Get_Struct(mrb, self, &olevariant_datatype, pvar);
    if (len == 1) {
        ole_val2variant(mrb, val, &(pvar->var));
    } else {
        vvt = argv[1];
        vt = NUM2INT(vvt);
        if ((vt & VT_TYPEMASK) == VT_RECORD) {
            mrb_raise(mrb, E_ARGUMENT_ERROR, "not supported VT_RECORD WIN32OLE_VARIANT object");
        }
        ole_val2olevariantdata(mrb, val, vt, pvar);
    }
    return self;
}

static SAFEARRAY *
get_locked_safe_array(mrb_state *mrb, mrb_value val)
{
    struct olevariantdata *pvar;
    SAFEARRAY *psa = NULL;
    HRESULT hr;
    Data_Get_Struct(mrb, val, &olevariant_datatype, pvar);
    if (!(V_VT(&(pvar->var)) & VT_ARRAY)) {
        mrb_raise(mrb, E_TYPE_ERROR, "variant type is not VT_ARRAY.");
    }
    psa = V_ISBYREF(&(pvar->var)) ? *V_ARRAYREF(&(pvar->var)) : V_ARRAY(&(pvar->var));
    if (psa == NULL) {
        return psa;
    }
    hr = SafeArrayLock(psa);
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_RUNTIME_ERROR, "failed to SafeArrayLock");
    }
    return psa;
}

static LONG *
ary2safe_array_index(mrb_state *mrb, int ary_size, mrb_value *ary, SAFEARRAY *psa)
{
    long dim;
    LONG *pid;
    long i;
    dim = SafeArrayGetDim(psa);
    if (dim != ary_size) {
        mrb_raise(mrb, E_ARGUMENT_ERROR, "unmatch number of indices");
    }
    pid = ALLOC_N(LONG, dim);
    if (pid == NULL) {
        mrb_raise(mrb, E_RUNTIME_ERROR, "failed to allocate memory for indices");
    }
    for (i = 0; i < dim; i++) {
        pid[i] = NUM2INT(ary[i]);
    }
    return pid;
}

static void
unlock_safe_array(mrb_state *mrb, SAFEARRAY *psa)
{
    HRESULT hr;
    hr = SafeArrayUnlock(psa);
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_RUNTIME_ERROR, "failed to SafeArrayUnlock");
    }
}

/*
 *  call-seq:
 *     WIN32OLE_VARIANT[i,j,...] #=> element of OLE array.
 *
 *  Returns the element of WIN32OLE_VARIANT object(OLE array).
 *  This method is available only when the variant type of
 *  WIN32OLE_VARIANT object is VT_ARRAY.
 *
 *  REMARK:
 *     The all indicies should be 0 or natural number and
 *     lower than or equal to max indicies.
 *     (This point is different with Ruby Array indicies.)
 *
 *     obj = WIN32OLE_VARIANT.new([[1,2,3],[4,5,6]])
 *     p obj[0,0] # => 1
 *     p obj[1,0] # => 4
 *     p obj[2,0] # => WIN32OLERuntimeError
 *     p obj[0, -1] # => WIN32OLERuntimeError
 *
 */
static mrb_value
folevariant_ary_aref(mrb_state *mrb, mrb_value self)
{
    mrb_int argc;
    mrb_value *argv;
    struct olevariantdata *pvar;
    SAFEARRAY *psa;
    mrb_value val = mrb_nil_value();
    VARIANT variant;
    LONG *pid;
    HRESULT hr;

    mrb_get_args(mrb, "*", &argv, &argc);

    Data_Get_Struct(mrb, self, &olevariant_datatype, pvar);
    if (!V_ISARRAY(&(pvar->var))) {
        mrb_raise(mrb, E_WIN32OLE_RUNTIME_ERROR,
                 "`[]' is not available for this variant type object");
    }
    psa = get_locked_safe_array(mrb, self);
    if (psa == NULL) {
        return val;
    }

    pid = ary2safe_array_index(mrb, argc, argv, psa);

    VariantInit(&variant);
    V_VT(&variant) = (V_VT(&(pvar->var)) & ~VT_ARRAY) | VT_BYREF;
    hr = SafeArrayPtrOfIndex(psa, pid, &V_BYREF(&variant));
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to SafeArrayPtrOfIndex");
    }
    val = ole_variant2val(mrb, &variant);

    unlock_safe_array(mrb, psa);
    if (pid) free(pid);
    return val;
}

/*
 *  call-seq:
 *     WIN32OLE_VARIANT[i,j,...] = val #=> set the element of OLE array
 *
 *  Set the element of WIN32OLE_VARIANT object(OLE array) to val.
 *  This method is available only when the variant type of
 *  WIN32OLE_VARIANT object is VT_ARRAY.
 *
 *  REMARK:
 *     The all indicies should be 0 or natural number and
 *     lower than or equal to max indicies.
 *     (This point is different with Ruby Array indicies.)
 *
 *     obj = WIN32OLE_VARIANT.new([[1,2,3],[4,5,6]])
 *     obj[0,0] = 7
 *     obj[1,0] = 8
 *     p obj.value # => [[7,2,3], [8,5,6]]
 *     obj[2,0] = 9 # => WIN32OLERuntimeError
 *     obj[0, -1] = 9 # => WIN32OLERuntimeError
 *
 */
static mrb_value
folevariant_ary_aset(mrb_state *mrb, mrb_value self)
{
    mrb_int argc;
    mrb_value *argv;
    struct olevariantdata *pvar;
    SAFEARRAY *psa;
    VARIANT var;
    VARTYPE vt;
    LONG *pid;
    HRESULT hr;
    VOID *p = NULL;

    mrb_get_args(mrb, "*", &argv, &argc);

    Data_Get_Struct(mrb, self, &olevariant_datatype, pvar);
    if (!V_ISARRAY(&(pvar->var))) {
        mrb_raise(mrb, E_WIN32OLE_RUNTIME_ERROR,
                 "`[]' is not available for this variant type object");
    }
    psa = get_locked_safe_array(mrb, self);
    if (psa == NULL) {
        mrb_raise(mrb, E_RUNTIME_ERROR, "failed to get SafeArray pointer");
    }

    pid = ary2safe_array_index(mrb, argc-1, argv, psa);

    VariantInit(&var);
    vt = (V_VT(&(pvar->var)) & ~VT_ARRAY);
    p = val2variant_ptr(mrb, argv[argc-1], &var, vt);
    if ((V_VT(&var) == VT_DISPATCH && V_DISPATCH(&var) == NULL) ||
        (V_VT(&var) == VT_UNKNOWN && V_UNKNOWN(&var) == NULL)) {
        mrb_raise(mrb, E_WIN32OLE_RUNTIME_ERROR, "argument does not have IDispatch or IUnknown Interface");
    }
    hr = SafeArrayPutElement(psa, pid, p);
    if (FAILED(hr)) {
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to SafeArrayPutElement");
    }

    unlock_safe_array(mrb, psa);
    if (pid) free(pid);
    return argv[argc-1];
}

/*
 *  call-seq:
 *     WIN32OLE_VARIANT.value #=> Ruby object.
 *
 *  Returns Ruby object value from OLE variant.
 *     obj = WIN32OLE_VARIANT.new(1, WIN32OLE::VARIANT::VT_BSTR)
 *     obj.value # => "1" (not Fixnum object, but String object "1")
 *
 */
static mrb_value
folevariant_value(mrb_state *mrb, mrb_value self)
{
    struct olevariantdata *pvar;
    mrb_value val = mrb_nil_value();
    VARTYPE vt;
    int dim;
    SAFEARRAY *psa;
    Data_Get_Struct(mrb, self, &olevariant_datatype, pvar);

    val = ole_variant2val(mrb, &(pvar->var));
    vt = V_VT(&(pvar->var));

    if ((vt & ~VT_BYREF) == (VT_UI1|VT_ARRAY)) {
        if (vt & VT_BYREF) {
            psa = *V_ARRAYREF(&(pvar->var));
        } else {
            psa  = V_ARRAY(&(pvar->var));
        }
        if (!psa) {
            return val;
        }
        dim = SafeArrayGetDim(psa);
        if (dim == 1) {
            val = mrb_funcall(mrb, val, "pack", 1, mrb_str_new_lit(mrb, "C*"));
        }
    }
    return val;
}

/*
 *  call-seq:
 *     WIN32OLE_VARIANT.vartype #=> OLE variant type.
 *
 *  Returns OLE variant type.
 *     obj = WIN32OLE_VARIANT.new("string")
 *     obj.vartype # => WIN32OLE::VARIANT::VT_BSTR
 *
 */
static mrb_value
folevariant_vartype(mrb_state *mrb, mrb_value self)
{
    struct olevariantdata *pvar;
    Data_Get_Struct(mrb, self, &olevariant_datatype, pvar);
    return INT2FIX(V_VT(&pvar->var));
}

/*
 *  call-seq:
 *     WIN32OLE_VARIANT.value = val #=> set WIN32OLE_VARIANT value to val.
 *
 *  Sets variant value to val. If the val type does not match variant value
 *  type(vartype), then val is changed to match variant value type(vartype)
 *  before setting val.
 *  Thie method is not available when vartype is VT_ARRAY(except VT_UI1|VT_ARRAY).
 *  If the vartype is VT_UI1|VT_ARRAY, the val should be String object.
 *
 *     obj = WIN32OLE_VARIANT.new(1) # obj.vartype is WIN32OLE::VARIANT::VT_I4
 *     obj.value = 3.2 # 3.2 is changed to 3 when setting value.
 *     p obj.value # => 3
 */
static mrb_value
folevariant_set_value(mrb_state *mrb, mrb_value self)
{
    mrb_value val;
    struct olevariantdata *pvar;
    VARTYPE vt;
    mrb_get_args(mrb, "o", &val);
    Data_Get_Struct(mrb, self, &olevariant_datatype, pvar);
    vt = V_VT(&(pvar->var));
    if (V_ISARRAY(&(pvar->var)) && ((vt & ~VT_BYREF) != (VT_UI1|VT_ARRAY) || !mrb_string_p(val))) {
        mrb_raise(mrb, E_WIN32OLE_RUNTIME_ERROR,
                 "`value=' is not available for this variant type object");
    }
    ole_val2olevariantdata(mrb, val, vt, pvar);
    return mrb_nil_value();
}

void
ole_variant2variant(mrb_state *mrb, mrb_value val, VARIANT *var)
{
    struct olevariantdata *pvar;
    Data_Get_Struct(mrb, val, &olevariant_datatype, pvar);
    VariantCopy(var, &(pvar->var));
}

void
Init_win32ole_variant(mrb_state *mrb)
{
    struct RClass *cWIN32OLE_VARIANT = mrb_define_class(mrb, "WIN32OLE_VARIANT", mrb->object_class);
    MRB_SET_INSTANCE_TT(cWIN32OLE_VARIANT, MRB_TT_DATA);

    mrb_define_class_method(mrb, cWIN32OLE_VARIANT, "array", folevariant_s_array, MRB_ARGS_REQ(2));
    mrb_define_method(mrb, cWIN32OLE_VARIANT, "initialize", folevariant_initialize, MRB_ARGS_ANY());
    mrb_define_method(mrb, cWIN32OLE_VARIANT, "value", folevariant_value, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_VARIANT, "value=", folevariant_set_value, MRB_ARGS_REQ(1));
    mrb_define_method(mrb, cWIN32OLE_VARIANT, "vartype", folevariant_vartype, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_VARIANT, "[]", folevariant_ary_aref, MRB_ARGS_ANY());
    mrb_define_method(mrb, cWIN32OLE_VARIANT, "[]=", folevariant_ary_aset, MRB_ARGS_ANY());

    /*
     * represents VT_EMPTY OLE object.
     */
    mrb_define_const(mrb, cWIN32OLE_VARIANT, "Empty",
            mrb_funcall(mrb, mrb_obj_value(cWIN32OLE_VARIANT), "new", 2, mrb_nil_value(), INT2FIX(VT_EMPTY)));

    /*
     * represents VT_NULL OLE object.
     */
    mrb_define_const(mrb, cWIN32OLE_VARIANT, "Null",
            mrb_funcall(mrb, mrb_obj_value(cWIN32OLE_VARIANT), "new", 2, mrb_nil_value(), INT2FIX(VT_NULL)));

    /*
     * represents Nothing of VB.NET or VB.
     */
    mrb_define_const(mrb, cWIN32OLE_VARIANT, "Nothing",
            mrb_funcall(mrb, mrb_obj_value(cWIN32OLE_VARIANT), "new", 2, mrb_nil_value(), INT2FIX(VT_DISPATCH)));

    /*
     * represents VT_ERROR variant with DISP_E_PARAMNOTFOUND.
     * This constants is used for not specified parameter.
     *
     *  fso = WIN32OLE.new("Scripting.FileSystemObject")
     *  fso.openTextFile(filename, WIN32OLE_VARIANT::NoParam, false)
     */
    mrb_define_const(mrb, cWIN32OLE_VARIANT, "NoParam",
            mrb_funcall(mrb, mrb_obj_value(cWIN32OLE_VARIANT), "new", 2, INT2NUM(DISP_E_PARAMNOTFOUND), INT2FIX(VT_ERROR)));
}
