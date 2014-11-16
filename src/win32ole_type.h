#ifndef WIN32OLE_TYPE_H
#define WIN32OLE_TYPE_H 1
#define C_WIN32OLE_TYPE (mrb_class_get(mrb, "WIN32OLE_TYPE"))
mrb_value create_win32ole_type(mrb_state *mrb, ITypeInfo *pTypeInfo, mrb_value name);
ITypeInfo *itypeinfo(mrb_state *mrb, mrb_value self);
mrb_value ole_type_from_itypeinfo(mrb_state *mrb, ITypeInfo *pTypeInfo);
void Init_win32ole_type(mrb_state *mrb);
#endif
