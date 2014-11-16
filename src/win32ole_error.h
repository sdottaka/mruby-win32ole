#ifndef WIN32OLE_ERROR_H
#define WIN32OLE_ERROR_H 1

#define E_WIN32OLE_RUNTIME_ERROR (mrb_class_get(mrb, "WIN32OLERuntimeError"))
void ole_raise(mrb_state *mrb, HRESULT hr, struct RClass *ecs, const char *fmt, ...);
void Init_win32ole_error(mrb_state *mrb);

#endif
