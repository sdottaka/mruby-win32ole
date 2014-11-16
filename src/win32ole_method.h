#ifndef WIN32OLE_METHOD_H
#define WIN32OLE_METHOD_H 1

struct olemethoddata {
    ITypeInfo *pOwnerTypeInfo;
    ITypeInfo *pTypeInfo;
    UINT index;
};

#define C_WIN32OLE_METHOD (mrb_class_get(mrb, "WIN32OLE_METHOD"))
mrb_value folemethod_s_allocate(mrb_state *mrb, struct RClass *klass);
mrb_value ole_methods_from_typeinfo(mrb_state *mrb, ITypeInfo *pTypeInfo, int mask);
mrb_value create_win32ole_method(mrb_state *mrb, ITypeInfo *pTypeInfo, mrb_value name);
struct olemethoddata *olemethod_data_get_struct(mrb_state *mrb, mrb_value obj);
void Init_win32ole_method(mrb_state *mrb);
#endif
