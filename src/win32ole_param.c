#include "win32ole.h"

struct oleparamdata {
    ITypeInfo *pTypeInfo;
    UINT method_index;
    UINT index;
};

static void oleparam_free(mrb_state *mrb, void *ptr);
static struct oleparamdata *oleparamdata_alloc(mrb_state *mrb);
static mrb_value foleparam_s_allocate(mrb_state *mrb, struct RClass *klass);
static mrb_value oleparam_ole_param_from_index(mrb_state *mrb, mrb_value self, ITypeInfo *pTypeInfo, UINT method_index, int param_index);
static mrb_value oleparam_ole_param(mrb_state *mrb, mrb_value self, mrb_value olemethod, int n);
static mrb_value foleparam_initialize(mrb_state *mrb, mrb_value self);
static mrb_value foleparam_name(mrb_state *mrb, mrb_value self);
static mrb_value ole_param_ole_type(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index, UINT index);
static mrb_value foleparam_ole_type(mrb_state *mrb, mrb_value self);
static mrb_value ole_param_ole_type_detail(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index, UINT index);
static mrb_value foleparam_ole_type_detail(mrb_state *mrb, mrb_value self);
static mrb_value ole_param_flag_mask(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index, UINT index, USHORT mask);
static mrb_value foleparam_input(mrb_state *mrb, mrb_value self);
static mrb_value foleparam_output(mrb_state *mrb, mrb_value self);
static mrb_value foleparam_optional(mrb_state *mrb, mrb_value self);
static mrb_value foleparam_retval(mrb_state *mrb, mrb_value self);
static mrb_value ole_param_default(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index, UINT index);
static mrb_value foleparam_default(mrb_state *mrb, mrb_value self);
static mrb_value foleparam_inspect(mrb_state *mrb, mrb_value self);

static const struct mrb_data_type oleparam_datatype = {
    "win32ole_param",
    oleparam_free
};

static void
oleparam_free(mrb_state *mrb, void *ptr)
{
    struct oleparamdata *pole = ptr;
    if (!ptr) return;
    OLE_FREE(pole->pTypeInfo);
    mrb_free(mrb, pole);
}

mrb_value
create_win32ole_param(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index, UINT index, mrb_value name)
{
    struct oleparamdata *pparam;
    mrb_value obj = foleparam_s_allocate(mrb, C_WIN32OLE_PARAM);
    Data_Get_Struct(mrb, obj, &oleparam_datatype, pparam);

    pparam->pTypeInfo = pTypeInfo;
    OLE_ADDREF(pTypeInfo);
    pparam->method_index = method_index;
    pparam->index = index;
    mrb_iv_set(mrb, obj, mrb_intern_lit(mrb, "name"), name);
    return obj;
}

/*
 * Document-class: WIN32OLE_PARAM
 *
 *   <code>WIN32OLE_PARAM</code> objects represent param information of
 *   the OLE method.
 */
static struct oleparamdata *
oleparamdata_alloc(mrb_state *mrb)
{
    struct oleparamdata *pparam = ALLOC(struct oleparamdata);
    pparam->pTypeInfo = NULL;
    pparam->method_index = 0;
    pparam->index = 0;
    return pparam;
}

static mrb_value
foleparam_s_allocate(mrb_state *mrb, struct RClass *klass)
{
    mrb_value obj;
    struct oleparamdata *pparam = oleparamdata_alloc(mrb);
    obj = mrb_obj_value(Data_Wrap_Struct(mrb, klass,
                                &oleparam_datatype, pparam));
    return obj;
}

static mrb_value
oleparam_ole_param_from_index(mrb_state *mrb, mrb_value self, ITypeInfo *pTypeInfo, UINT method_index, int param_index)
{
    FUNCDESC *pFuncDesc;
    HRESULT hr;
    BSTR *bstrs;
    UINT len;
    struct oleparamdata *pparam;
    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if (FAILED(hr))
        ole_raise(mrb, hr, E_RUNTIME_ERROR, "fail to ITypeInfo::GetFuncDesc");

    len = 0;
    bstrs = ALLOCA_N(BSTR, pFuncDesc->cParams + 1);
    hr = pTypeInfo->lpVtbl->GetNames(pTypeInfo, pFuncDesc->memid,
                                     bstrs, pFuncDesc->cParams + 1,
                                     &len);
    if (FAILED(hr)) {
        pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
        ole_raise(mrb, hr, E_RUNTIME_ERROR, "fail to ITypeInfo::GetNames");
    }
    SysFreeString(bstrs[0]);
    if (param_index < 1 || len <= (UINT)param_index)
    {
        pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
        mrb_raisef(mrb, E_INDEX_ERROR, "index of param must be in 1..%S", mrb_fixnum_value(len));
    }

    Data_Get_Struct(mrb, self, &oleparam_datatype, pparam);
    pparam->pTypeInfo = pTypeInfo;
    OLE_ADDREF(pTypeInfo);
    pparam->method_index = method_index;
    pparam->index = param_index - 1;
    mrb_iv_set(mrb, self, mrb_intern_lit(mrb, "name"), WC2VSTR(mrb, bstrs[param_index]));

    pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
    return self;
}

static mrb_value
oleparam_ole_param(mrb_state *mrb, mrb_value self, mrb_value olemethod, int n)
{
    struct olemethoddata *pmethod = olemethod_data_get_struct(mrb, olemethod);
    return oleparam_ole_param_from_index(mrb, self, pmethod->pTypeInfo, pmethod->index, n);
}

/*
 * call-seq:
 *    WIN32OLE_PARAM.new(method, n) -> WIN32OLE_PARAM object
 *
 * Returns WIN32OLE_PARAM object which represents OLE parameter information.
 * 1st argument should be WIN32OLE_METHOD object.
 * 2nd argument `n' is n-th parameter of the method specified by 1st argument.
 *
 *    tobj = WIN32OLE_TYPE.new('Microsoft Scripting Runtime', 'IFileSystem')
 *    method = WIN32OLE_METHOD.new(tobj, 'CreateTextFile')
 *    param = WIN32OLE_PARAM.new(method, 2) # => #<WIN32OLE_PARAM:Overwrite=true>
 *
 */
static mrb_value
foleparam_initialize(mrb_state *mrb, mrb_value self)
{
    mrb_value olemethod;
    mrb_int idx;
    struct oleparamdata *pparam = (struct oleparamdata *)DATA_PTR(self);
    if (pparam) {
        mrb_free(mrb, pparam);
    }
    mrb_data_init(self, NULL, &oleparam_datatype);

    mrb_get_args(mrb, "oi", &olemethod, &idx);
    if (!mrb_obj_is_kind_of(mrb, olemethod, C_WIN32OLE_METHOD)) {
        mrb_raise(mrb, E_TYPE_ERROR, "1st parameter must be WIN32OLE_METHOD object");
    }

    pparam = oleparamdata_alloc(mrb);
    mrb_data_init(self, pparam, &oleparam_datatype);

    return oleparam_ole_param(mrb, self, olemethod, idx);
}

/*
 *  call-seq:
 *     WIN32OLE_PARAM#name
 *
 *  Returns name.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbook')
 *     method = WIN32OLE_METHOD.new(tobj, 'SaveAs')
 *     param1 = method.params[0]
 *     puts param1.name # => Filename
 */
static mrb_value
foleparam_name(mrb_state *mrb, mrb_value self)
{
    return mrb_iv_get(mrb, self, mrb_intern_lit(mrb, "name"));
}

static mrb_value
ole_param_ole_type(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index, UINT index)
{
    FUNCDESC *pFuncDesc;
    HRESULT hr;
    mrb_value type = mrb_str_new_lit(mrb, "unknown type");
    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if (FAILED(hr))
        return type;
    type = ole_typedesc2val(mrb, pTypeInfo,
                            &(pFuncDesc->lprgelemdescParam[index].tdesc), mrb_nil_value());
    pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
    return type;
}

/*
 *  call-seq:
 *     WIN32OLE_PARAM#ole_type
 *
 *  Returns OLE type of WIN32OLE_PARAM object(parameter of OLE method).
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbook')
 *     method = WIN32OLE_METHOD.new(tobj, 'SaveAs')
 *     param1 = method.params[0]
 *     puts param1.ole_type # => VARIANT
 */
static mrb_value
foleparam_ole_type(mrb_state *mrb, mrb_value self)
{
    struct oleparamdata *pparam;
    Data_Get_Struct(mrb, self, &oleparam_datatype, pparam);
    return ole_param_ole_type(mrb, pparam->pTypeInfo, pparam->method_index,
                              pparam->index);
}

static mrb_value
ole_param_ole_type_detail(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index, UINT index)
{
    FUNCDESC *pFuncDesc;
    HRESULT hr;
    mrb_value typedetail = mrb_ary_new(mrb);
    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if (FAILED(hr))
        return typedetail;
    ole_typedesc2val(mrb, pTypeInfo,
                     &(pFuncDesc->lprgelemdescParam[index].tdesc), typedetail);
    pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
    return typedetail;
}

/*
 *  call-seq:
 *     WIN32OLE_PARAM#ole_type_detail
 *
 *  Returns detail information of type of argument.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'IWorksheetFunction')
 *     method = WIN32OLE_METHOD.new(tobj, 'SumIf')
 *     param1 = method.params[0]
 *     p param1.ole_type_detail # => ["PTR", "USERDEFINED", "Range"]
 */
static mrb_value
foleparam_ole_type_detail(mrb_state *mrb, mrb_value self)
{
    struct oleparamdata *pparam;
    Data_Get_Struct(mrb, self, &oleparam_datatype, pparam);
    return ole_param_ole_type_detail(mrb, pparam->pTypeInfo, pparam->method_index,
                                     pparam->index);
}

static mrb_value
ole_param_flag_mask(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index, UINT index, USHORT mask)
{
    FUNCDESC *pFuncDesc;
    HRESULT hr;
    mrb_value ret = mrb_false_value();
    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if(FAILED(hr))
        return ret;
    if (V_UNION1((&(pFuncDesc->lprgelemdescParam[index])), paramdesc).wParamFlags &mask)
        ret = mrb_true_value();
    pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
    return ret;
}

/*
 *  call-seq:
 *     WIN32OLE_PARAM#input?
 *
 *  Returns true if the parameter is input.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbook')
 *     method = WIN32OLE_METHOD.new(tobj, 'SaveAs')
 *     param1 = method.params[0]
 *     puts param1.input? # => true
 */
static mrb_value
foleparam_input(mrb_state *mrb, mrb_value self)
{
    struct oleparamdata *pparam;
    Data_Get_Struct(mrb, self, &oleparam_datatype, pparam);
    return ole_param_flag_mask(mrb, pparam->pTypeInfo, pparam->method_index,
                               pparam->index, PARAMFLAG_FIN);
}

/*
 *  call-seq:
 *     WIN32OLE#output?
 *
 *  Returns true if argument is output.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Internet Controls', 'DWebBrowserEvents')
 *     method = WIN32OLE_METHOD.new(tobj, 'NewWindow')
 *     method.params.each do |param|
 *       puts "#{param.name} #{param.output?}"
 *     end
 *
 *     The result of above script is following:
 *       URL false
 *       Flags false
 *       TargetFrameName false
 *       PostData false
 *       Headers false
 *       Processed true
 */
static mrb_value
foleparam_output(mrb_state *mrb, mrb_value self)
{
    struct oleparamdata *pparam;
    Data_Get_Struct(mrb, self, &oleparam_datatype, pparam);
    return ole_param_flag_mask(mrb, pparam->pTypeInfo, pparam->method_index,
                               pparam->index, PARAMFLAG_FOUT);
}

/*
 *  call-seq:
 *     WIN32OLE_PARAM#optional?
 *
 *  Returns true if argument is optional.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbook')
 *     method = WIN32OLE_METHOD.new(tobj, 'SaveAs')
 *     param1 = method.params[0]
 *     puts "#{param1.name} #{param1.optional?}" # => Filename true
 */
static mrb_value
foleparam_optional(mrb_state *mrb, mrb_value self)
{
    struct oleparamdata *pparam;
    Data_Get_Struct(mrb, self, &oleparam_datatype, pparam);
    return ole_param_flag_mask(mrb, pparam->pTypeInfo, pparam->method_index,
                               pparam->index, PARAMFLAG_FOPT);
}

/*
 *  call-seq:
 *     WIN32OLE_PARAM#retval?
 *
 *  Returns true if argument is return value.
 *     tobj = WIN32OLE_TYPE.new('DirectX 7 for Visual Basic Type Library',
 *                              'DirectPlayLobbyConnection')
 *     method = WIN32OLE_METHOD.new(tobj, 'GetPlayerShortName')
 *     param = method.params[0]
 *     puts "#{param.name} #{param.retval?}"  # => name true
 */
static mrb_value
foleparam_retval(mrb_state *mrb, mrb_value self)
{
    struct oleparamdata *pparam;
    Data_Get_Struct(mrb, self, &oleparam_datatype, pparam);
    return ole_param_flag_mask(mrb, pparam->pTypeInfo, pparam->method_index,
                               pparam->index, PARAMFLAG_FRETVAL);
}

static mrb_value
ole_param_default(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT method_index, UINT index)
{
    FUNCDESC *pFuncDesc;
    ELEMDESC *pElemDesc;
    PARAMDESCEX * pParamDescEx;
    HRESULT hr;
    USHORT wParamFlags;
    USHORT mask = PARAMFLAG_FOPT|PARAMFLAG_FHASDEFAULT;
    mrb_value defval = mrb_nil_value();
    hr = pTypeInfo->lpVtbl->GetFuncDesc(pTypeInfo, method_index, &pFuncDesc);
    if (FAILED(hr))
        return defval;
    pElemDesc = &pFuncDesc->lprgelemdescParam[index];
    wParamFlags = V_UNION1(pElemDesc, paramdesc).wParamFlags;
    if ((wParamFlags & mask) == mask) {
         pParamDescEx = V_UNION1(pElemDesc, paramdesc).pparamdescex;
         defval = ole_variant2val(mrb, &pParamDescEx->varDefaultValue);
    }
    pTypeInfo->lpVtbl->ReleaseFuncDesc(pTypeInfo, pFuncDesc);
    return defval;
}

/*
 *  call-seq:
 *     WIN32OLE_PARAM#default
 *
 *  Returns default value. If the default value does not exist,
 *  this method returns nil.
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'Workbook')
 *     method = WIN32OLE_METHOD.new(tobj, 'SaveAs')
 *     method.params.each do |param|
 *       if param.default
 *         puts "#{param.name} (= #{param.default})"
 *       else
 *         puts "#{param}"
 *       end
 *     end
 *
 *     The above script result is following:
 *         Filename
 *         FileFormat
 *         Password
 *         WriteResPassword
 *         ReadOnlyRecommended
 *         CreateBackup
 *         AccessMode (= 1)
 *         ConflictResolution
 *         AddToMru
 *         TextCodepage
 *         TextVisualLayout
 */
static mrb_value
foleparam_default(mrb_state *mrb, mrb_value self)
{
    struct oleparamdata *pparam;
    Data_Get_Struct(mrb, self, &oleparam_datatype, pparam);
    return ole_param_default(mrb, pparam->pTypeInfo, pparam->method_index,
                             pparam->index);
}

/*
 *  call-seq:
 *     WIN32OLE_PARAM#inspect -> String
 *
 *  Returns the parameter name with class name. If the parameter has default value,
 *  then returns name=value string with class name.
 *
 */
static mrb_value
foleparam_inspect(mrb_state *mrb, mrb_value self)
{
    mrb_value detail = foleparam_name(mrb, self);
    mrb_value defval = foleparam_default(mrb, self);
    if (!mrb_nil_p(defval)) {
        mrb_str_cat_lit(mrb, detail, "=");
        mrb_str_concat(mrb, detail, mrb_inspect(mrb, defval));
    }
    return make_inspect(mrb, "WIN32OLE_PARAM", detail);
}

void
Init_win32ole_param(mrb_state *mrb)
{
    struct RClass *cWIN32OLE_PARAM = mrb_define_class(mrb, "WIN32OLE_PARAM", mrb->object_class);
    MRB_SET_INSTANCE_TT(cWIN32OLE_PARAM, MRB_TT_DATA);

    mrb_define_method(mrb, cWIN32OLE_PARAM, "initialize", foleparam_initialize, MRB_ARGS_REQ(2));
    mrb_define_method(mrb, cWIN32OLE_PARAM, "name", foleparam_name, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_PARAM, "ole_type", foleparam_ole_type, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_PARAM, "ole_type_detail", foleparam_ole_type_detail, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_PARAM, "input?", foleparam_input, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_PARAM, "output?", foleparam_output, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_PARAM, "optional?", foleparam_optional, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_PARAM, "retval?", foleparam_retval, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_PARAM, "default", foleparam_default, MRB_ARGS_NONE());
    mrb_define_alias(mrb, cWIN32OLE_PARAM, "to_s", "name");
    mrb_define_method(mrb, cWIN32OLE_PARAM, "inspect", foleparam_inspect, MRB_ARGS_NONE());
}
