#ifndef WIN32OLE_TYPELIB_H
#define WIN32OLE_TYPELIB_H 1

#define C_WIN32OLE_TYPELIB (mrb_class_get(mrb, "WIN32OLE_TYPELIB"))

void Init_win32ole_typelib(mrb_state *mrb);
ITypeLib * itypelib(mrb_state *mrb, mrb_value self);
mrb_value typelib_file(mrb_state *mrb, mrb_value ole);
mrb_value create_win32ole_typelib(mrb_state *mrb, ITypeLib *pTypeLib);
mrb_value ole_typelib_from_itypeinfo(mrb_state *mrb, ITypeInfo *pTypeInfo);
#endif
