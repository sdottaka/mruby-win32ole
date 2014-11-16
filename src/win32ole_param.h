#ifndef WIN32OLE_PARAM_H
#define WIN32OLE_PARAM_H

#define C_WIN32OLE_PARAM (mrb_class_get(mrb, "WIN32OLE_PARAM"))
mrb_value create_win32ole_param(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index, UINT index, mrb_value name);
void Init_win32ole_param(mrb_state *mrb);

#endif

