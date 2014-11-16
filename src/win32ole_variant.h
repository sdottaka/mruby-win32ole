#ifndef WIN32OLE_VARIANT_H
#define WIN32OLE_VARIANT_H 1

#define C_WIN32OLE_VARIANT (mrb_class_get(mrb, "WIN32OLE_VARIANT"))
void ole_variant2variant(mrb_state *mrb, mrb_value val, VARIANT *var);
void Init_win32ole_variant(mrb_state *mrb);

#endif

