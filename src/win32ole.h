#ifndef WIN32OLE_H
#define WIN32OLE_H 1
#include "mruby.h"
#include "mruby/string.h"
#include "mruby/error.h"
#include "mruby/variable.h"
#include "mruby/array.h"
#include "mruby/data.h"
#include "mruby/class.h"
#include "mruby/khash.h"
#include "mruby/hash.h"
#include "mruby/numeric.h"
#include <malloc.h>

#ifndef __CYGWIN__
#define strcasecmp   _stricmp
#endif
#define FIX2INT(val) mrb_fixnum((val))
#define INT2FIX(val) mrb_fixnum_value((val))
#define INT2NUM(val) (FIXABLE((val)) ? mrb_fixnum_value((mrb_int)(val)) : mrb_float_value(mrb, (mrb_float)(val)))
#define UINT2NUM(val) INT2NUM((unsigned mrb_int)(val))
#define NUM2CHR(val) ((mrb_string_p(val) && (RSTRING_LEN(val)>=1)) ? RSTRING_PTR(val)[0] : (char)(NUM2INT(val) & 0xff))
#define TO_PDT(val) ((mrb_type((val)) == MRB_TT_FLOAT) ? mrb_float((val)) : mrb_int(mrb, (val)))
#define NUM2INT(val) ((mrb_int)TO_PDT((val)))
#define NUM2UINT(val) ((unsigned mrb_int)TO_PDT((val)))
#define NUM2LONG(val) (mrb_int(mrb, (val)))
#define LL2NUM(val) INT2NUM((val))
#define ULL2NUM(val) INT2NUM((unsigned __int64)(val))
#define NUM2LL(val) ((__int64)(TO_PDT((val))))
#define NUM2ULL(val) ((unsigned __int64)(TO_PDT((val))))
#define NUM2DBL(val) (mrb_to_flo(mrb, (val)))
#define ALLOC(type) ((type*)mrb_malloc(mrb, sizeof(type)))
#define ALLOC_N(type, n) ((type *)mrb_malloc(mrb, sizeof(type) * (n)))
#define ALLOCA_N(type, n) ((type *)alloca(sizeof(type) * (n)))
#define MEMCMP(p1,p2,type,n) memcmp((p1), (p2), sizeof(type)*(n))

#define GNUC_OLDER_3_4_4 \
    ((__GNUC__ < 3) || \
     ((__GNUC__ <= 3) && (__GNUC_MINOR__ < 4)) || \
     ((__GNUC__ <= 3) && (__GNUC_MINOR__ <= 4) && (__GNUC_PATCHLEVEL__ <= 4)))

#if (defined(__GNUC__)) && (GNUC_OLDER_3_4_4)
#ifndef NONAMELESSUNION
#define NONAMELESSUNION 1
#endif
#endif

#include <ctype.h>

#include <windows.h>
#include <ocidl.h>
#include <olectl.h>
#include <ole2.h>
#if defined(HAVE_TYPE_IMULTILANGUAGE2) || defined(HAVE_TYPE_IMULTILANGUAGE)
#include <mlang.h>
#endif
#include <stdlib.h>
#include <math.h>
/*#ifdef HAVE_STDARG_PROTOTYPES*/
#if 1
#include <stdarg.h>
#define va_init_list(a,b) va_start(a,b)
#else
#include <varargs.h>
#define va_init_list(a,b) va_start(a)
#endif
#include <objidl.h>

#define DOUT fprintf(stderr,"%s(%d)\n", __FILE__, __LINE__)
#define DOUTS(x) fprintf(stderr,"%s(%d):" #x "=%s\n",__FILE__, __LINE__,x)
#define DOUTMSG(x) fprintf(stderr, "%s(%d):" #x "\n",__FILE__, __LINE__)
#define DOUTI(x) fprintf(stderr, "%s(%d):" #x "=%d\n",__FILE__, __LINE__,x)
#define DOUTD(x) fprintf(stderr, "%s(%d):" #x "=%f\n",__FILE__, __LINE__,x)

#if (defined(__GNUC__)) && (GNUC_OLDER_3_4_4)
#define V_UNION1(X, Y) ((X)->u.Y)
#else
#define V_UNION1(X, Y) ((X)->Y)
#endif

#if (defined(__GNUC__)) && (GNUC_OLDER_3_4_4)
#undef V_UNION
#define V_UNION(X,Y) ((X)->n1.n2.n3.Y)

#undef V_VT
#define V_VT(X) ((X)->n1.n2.vt)

#undef V_BOOL
#define V_BOOL(X) V_UNION(X,boolVal)
#endif

#ifndef V_I1REF
#define V_I1REF(X) V_UNION(X, pcVal)
#endif

#ifndef V_UI2REF
#define V_UI2REF(X) V_UNION(X, puiVal)
#endif

#ifndef V_INT
#define V_INT(X) V_UNION(X, intVal)
#endif

#ifndef V_INTREF
#define V_INTREF(X) V_UNION(X, pintVal)
#endif

#ifndef V_UINT
#define V_UINT(X) V_UNION(X, uintVal)
#endif

#ifndef V_UINTREF
#define V_UINTREF(X) V_UNION(X, puintVal)
#endif

#ifndef V_RECORD
#define V_RECORD(X)     ((X)->pvRecord)
#endif

#ifndef V_RECORDINFO
#define V_RECORDINFO(X) ((X)->pRecInfo)
#endif

#if defined(HAVE_LONG_LONG) || defined(__int64) || defined(_MSC_VER)
#define I8_2_NUM LL2NUM
#define UI8_2_NUM ULL2NUM
#define NUM2I8  NUM2LL
#define NUM2UI8 NUM2ULL
#else
#define I8_2_NUM INT2NUM
#define UI8_2_NUM UINT2NUM
#define NUM2I8  NUM2INT
#define NUM2UI8 NUM2UINT
#endif

#define OLE_ADDREF(X) (X) ? ((X)->lpVtbl->AddRef(X)) : 0
#define OLE_RELEASE(X) (X) ? ((X)->lpVtbl->Release(X)) : 0
#define OLE_FREE(x) {\
    if(ole_initialized() == TRUE) {\
        if(x) {\
            OLE_RELEASE(x);\
            (x) = 0;\
        }\
    }\
}

#define OLE_GET_TYPEATTR(X, Y) ((X)->lpVtbl->GetTypeAttr((X), (Y)))
#define OLE_RELEASE_TYPEATTR(X, Y) ((X)->lpVtbl->ReleaseTypeAttr((X), (Y)))

struct oledata {
    IDispatch *pDispatch;
};

extern const mrb_data_type ole_datatype;

#define OLEData_Get_Struct(mrb, obj, pole) {\
    Data_Get_Struct(mrb, obj, &ole_datatype, pole);\
    if(!pole->pDispatch) {\
        mrb_raise(mrb, E_RUNTIME_ERROR, "failed to get Dispatch Interface");\
    }\
}

#define C_WIN32OLE (mrb_class_get(mrb, "WIN32OLE"))
LCID cWIN32OLE_lcid;

BSTR ole_vstr2wc(mrb_state *mrb, mrb_value vstr);
LONG reg_open_key(HKEY hkey, const char *name, HKEY *phkey);
LONG reg_open_vkey(mrb_state *mrb, HKEY hkey, mrb_value key, HKEY *phkey);
mrb_value reg_enum_key(mrb_state *mrb, HKEY hkey, DWORD i);
mrb_value reg_get_val(mrb_state *mrb, HKEY hkey, const char *subkey);
mrb_value reg_get_val2(mrb_state *mrb, HKEY hkey, const char *subkey);
void ole_initialize(mrb_state *mrb);
mrb_value default_inspect(mrb_state *mrb, mrb_value self, const char *class_name);
char *ole_wc2mb(mrb_state *mrb, LPWSTR pw);
mrb_value ole_wc2vstr(mrb_state *mrb, LPWSTR pw, BOOL isfree);

#define WC2VSTR(mrb, x) ole_wc2vstr((mrb), (x), TRUE)

BOOL ole_initialized(void);
HRESULT ole_docinfo_from_type(mrb_state *mrb, ITypeInfo *pTypeInfo, BSTR *name, BSTR *helpstr, DWORD *helpcontext, BSTR *helpfile);
mrb_value ole_typedesc2val(mrb_state *mrb, ITypeInfo *pTypeInfo, TYPEDESC *pTypeDesc, mrb_value typedetails);
mrb_value make_inspect(mrb_state *mrb, const char *class_name, mrb_value detail);
void ole_val2variant(mrb_state *mrb, mrb_value val, VARIANT *var);
void ole_val2variant2(mrb_state *mrb, mrb_value val, VARIANT *var);
void ole_val2variant_ex(mrb_state *mrb, mrb_value val, VARIANT *var, VARTYPE vt);
mrb_value ole_variant2val(mrb_state *mrb, VARIANT *pvar);
HRESULT ole_val_ary2variant_ary(mrb_state *mrb, mrb_value val, VARIANT *var, VARTYPE vt);
VOID *val2variant_ptr(mrb_state *mrb, mrb_value val, VARIANT *var, VARTYPE vt);
HRESULT typelib_from_val(mrb_state *mrb, mrb_value obj, ITypeLib **pTypeLib);
const char *ole_obj_to_cstr(mrb_state *mrb, mrb_value obj);

#include "win32ole_variant_m.h"
#include "win32ole_typelib.h"
#include "win32ole_type.h"
#include "win32ole_variable.h"
#include "win32ole_method.h"
#include "win32ole_param.h"
#include "win32ole_event.h"
#include "win32ole_variant.h"
#include "win32ole_record.h"
#include "win32ole_error.h"

#endif
