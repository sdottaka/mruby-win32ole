#ifndef WIN32OLE_RECORD_H
#define WIN32OLE_RECORD_H 1

#define C_WIN32OLE_RECORD (mrb_class_get(mrb, "WIN32OLE_RECORD"))
void ole_rec2variant(mrb_state *mrb, mrb_value rec, VARIANT *var);
void olerecord_set_ivar(mrb_state *mrb, mrb_value obj, IRecordInfo *pri, void *prec);
mrb_value create_win32ole_record(mrb_state *mrb, IRecordInfo *pri, void *prec);
void Init_win32ole_record(mrb_state *mrb);

#endif
