#ifndef WIN32OLE_VARIABLE_H
#define WIN32OLE_VARIABLE_H 1

#define C_WIN32OLE_VARIABLE (mrb_class_get(mrb, "WIN32OLE_VARIABLE"))
mrb_value create_win32ole_variable(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT index, mrb_value name);
void Init_win32ole_variable(mrb_state *mrb);

#endif
