#include "win32ole.h"

struct olevariabledata {
    ITypeInfo *pTypeInfo;
    UINT index;
};

static void olevariable_free(mrb_state *mrb, void *ptr);
static mrb_value folevariable_name(mrb_state *mrb, mrb_value self);
static mrb_value ole_variable_ole_type(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT var_index);
static mrb_value folevariable_ole_type(mrb_state *mrb, mrb_value self);
static mrb_value ole_variable_ole_type_detail(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT var_index);
static mrb_value folevariable_ole_type_detail(mrb_state *mrb, mrb_value self);
static mrb_value ole_variable_value(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT var_index);
static mrb_value folevariable_value(mrb_state *mrb, mrb_value self);
static mrb_value ole_variable_visible(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT var_index);
static mrb_value folevariable_visible(mrb_state *mrb, mrb_value self);
static mrb_value ole_variable_kind(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT var_index);
static mrb_value folevariable_variable_kind(mrb_state *mrb, mrb_value self);
static mrb_value ole_variable_varkind(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT var_index);
static mrb_value folevariable_varkind(mrb_state *mrb, mrb_value self);
static mrb_value folevariable_inspect(mrb_state *mrb, mrb_value self);

static const mrb_data_type olevariable_datatype = {
    "win32ole_variable",
    olevariable_free
};

static void
olevariable_free(mrb_state *mrb, void *ptr)
{
    struct olevariabledata *polevar = ptr;
    OLE_FREE(polevar->pTypeInfo);
    mrb_free(mrb, polevar);
}

/*
 * Document-class: WIN32OLE_VARIABLE
 *
 *   <code>WIN32OLE_VARIABLE</code> objects represent OLE variable information.
 */

mrb_value
create_win32ole_variable(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT index, mrb_value name)
{
    struct olevariabledata *pvar;
    struct RData *data;
    mrb_value obj;

    Data_Make_Struct(mrb, C_WIN32OLE_VARIABLE, struct olevariabledata,
                     &olevariable_datatype, pvar, data);
    obj = mrb_obj_value(data);
    pvar->pTypeInfo = pTypeInfo;
    OLE_ADDREF(pTypeInfo);
    pvar->index = index;
    mrb_iv_set(mrb, obj, mrb_intern_lit(mrb, "name"), name);
    return obj;
}

/*
 *  call-seq:
 *     WIN32OLE_VARIABLE#name
 *
 *  Returns the name of variable.
 *
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'XlSheetType')
 *     variables = tobj.variables
 *     variables.each do |variable|
 *       puts "#{variable.name}"
 *     end
 *
 *     The result of above script is following:
 *       xlChart
 *       xlDialogSheet
 *       xlExcel4IntlMacroSheet
 *       xlExcel4MacroSheet
 *       xlWorksheet
 *
 */
static mrb_value
folevariable_name(mrb_state *mrb, mrb_value self)
{
    return mrb_iv_get(mrb, self, mrb_intern_lit(mrb, "name"));
}

static mrb_value
ole_variable_ole_type(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT var_index)
{
    VARDESC *pVarDesc;
    HRESULT hr;
    mrb_value type;
    hr = pTypeInfo->lpVtbl->GetVarDesc(pTypeInfo, var_index, &pVarDesc);
    if (FAILED(hr))
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to GetVarDesc");
    type = ole_typedesc2val(mrb, pTypeInfo, &(pVarDesc->elemdescVar.tdesc), mrb_nil_value());
    pTypeInfo->lpVtbl->ReleaseVarDesc(pTypeInfo, pVarDesc);
    return type;
}

/*
 *   call-seq:
 *      WIN32OLE_VARIABLE#ole_type
 *
 *   Returns OLE type string.
 *
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'XlSheetType')
 *     variables = tobj.variables
 *     variables.each do |variable|
 *       puts "#{variable.ole_type} #{variable.name}"
 *     end
 *
 *     The result of above script is following:
 *       INT xlChart
 *       INT xlDialogSheet
 *       INT xlExcel4IntlMacroSheet
 *       INT xlExcel4MacroSheet
 *       INT xlWorksheet
 *
 */
static mrb_value
folevariable_ole_type(mrb_state *mrb, mrb_value self)
{
    struct olevariabledata *pvar;
    Data_Get_Struct(mrb, self, &olevariable_datatype, pvar);
    return ole_variable_ole_type(mrb, pvar->pTypeInfo, pvar->index);
}

static mrb_value
ole_variable_ole_type_detail(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT var_index)
{
    VARDESC *pVarDesc;
    HRESULT hr;
    mrb_value type = mrb_ary_new(mrb);
    hr = pTypeInfo->lpVtbl->GetVarDesc(pTypeInfo, var_index, &pVarDesc);
    if (FAILED(hr))
        ole_raise(mrb, hr, E_WIN32OLE_RUNTIME_ERROR, "failed to GetVarDesc");
    ole_typedesc2val(mrb, pTypeInfo, &(pVarDesc->elemdescVar.tdesc), type);
    pTypeInfo->lpVtbl->ReleaseVarDesc(pTypeInfo, pVarDesc);
    return type;
}

/*
 *  call-seq:
 *     WIN32OLE_VARIABLE#ole_type_detail
 *
 *  Returns detail information of type. The information is array of type.
 *
 *     tobj = WIN32OLE_TYPE.new('DirectX 7 for Visual Basic Type Library', 'D3DCLIPSTATUS')
 *     variable = tobj.variables.find {|variable| variable.name == 'lFlags'}
 *     tdetail  = variable.ole_type_detail
 *     p tdetail # => ["USERDEFINED", "CONST_D3DCLIPSTATUSFLAGS"]
 *
 */
static mrb_value
folevariable_ole_type_detail(mrb_state *mrb, mrb_value self)
{
    struct olevariabledata *pvar;
    Data_Get_Struct(mrb, self, &olevariable_datatype, pvar);
    return ole_variable_ole_type_detail(mrb, pvar->pTypeInfo, pvar->index);
}

static mrb_value
ole_variable_value(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT var_index)
{
    VARDESC *pVarDesc;
    HRESULT hr;
    mrb_value val = mrb_nil_value();
    hr = pTypeInfo->lpVtbl->GetVarDesc(pTypeInfo, var_index, &pVarDesc);
    if (FAILED(hr))
        return mrb_nil_value();
    if(pVarDesc->varkind == VAR_CONST)
        val = ole_variant2val(mrb, V_UNION1(pVarDesc, lpvarValue));
    pTypeInfo->lpVtbl->ReleaseVarDesc(pTypeInfo, pVarDesc);
    return val;
}

/*
 *  call-seq:
 *     WIN32OLE_VARIABLE#value
 *
 *  Returns value if value is exists. If the value does not exist,
 *  this method returns nil.
 *
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'XlSheetType')
 *     variables = tobj.variables
 *     variables.each do |variable|
 *       puts "#{variable.name} #{variable.value}"
 *     end
 *
 *     The result of above script is following:
 *       xlChart = -4109
 *       xlDialogSheet = -4116
 *       xlExcel4IntlMacroSheet = 4
 *       xlExcel4MacroSheet = 3
 *       xlWorksheet = -4167
 *
 */
static mrb_value
folevariable_value(mrb_state *mrb, mrb_value self)
{
    struct olevariabledata *pvar;
    Data_Get_Struct(mrb, self, &olevariable_datatype, pvar);
    return ole_variable_value(mrb, pvar->pTypeInfo, pvar->index);
}

static mrb_value
ole_variable_visible(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT var_index)
{
    VARDESC *pVarDesc;
    HRESULT hr;
    mrb_value visible = mrb_false_value();
    hr = pTypeInfo->lpVtbl->GetVarDesc(pTypeInfo, var_index, &pVarDesc);
    if (FAILED(hr))
        return visible;
    if (!(pVarDesc->wVarFlags & (VARFLAG_FHIDDEN |
                                 VARFLAG_FRESTRICTED |
                                 VARFLAG_FNONBROWSABLE))) {
        visible = mrb_true_value();
    }
    pTypeInfo->lpVtbl->ReleaseVarDesc(pTypeInfo, pVarDesc);
    return visible;
}

/*
 *  call-seq:
 *     WIN32OLE_VARIABLE#visible?
 *
 *  Returns true if the variable is public.
 *
 *     tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'XlSheetType')
 *     variables = tobj.variables
 *     variables.each do |variable|
 *       puts "#{variable.name} #{variable.visible?}"
 *     end
 *
 *     The result of above script is following:
 *       xlChart true
 *       xlDialogSheet true
 *       xlExcel4IntlMacroSheet true
 *       xlExcel4MacroSheet true
 *       xlWorksheet true
 *
 */
static mrb_value
folevariable_visible(mrb_state *mrb, mrb_value self)
{
    struct olevariabledata *pvar;
    Data_Get_Struct(mrb, self, &olevariable_datatype, pvar);
    return ole_variable_visible(mrb, pvar->pTypeInfo, pvar->index);
}

static mrb_value
ole_variable_kind(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT var_index)
{
    VARDESC *pVarDesc;
    HRESULT hr;
    mrb_value kind = mrb_str_new_lit(mrb, "UNKNOWN");
    hr = pTypeInfo->lpVtbl->GetVarDesc(pTypeInfo, var_index, &pVarDesc);
    if (FAILED(hr))
        return kind;
    switch(pVarDesc->varkind) {
    case VAR_PERINSTANCE:
        kind = mrb_str_new_lit(mrb, "PERINSTANCE");
        break;
    case VAR_STATIC:
        kind = mrb_str_new_lit(mrb, "STATIC");
        break;
    case VAR_CONST:
        kind = mrb_str_new_lit(mrb, "CONSTANT");
        break;
    case VAR_DISPATCH:
        kind = mrb_str_new_lit(mrb, "DISPATCH");
        break;
    default:
        break;
    }
    pTypeInfo->lpVtbl->ReleaseVarDesc(pTypeInfo, pVarDesc);
    return kind;
}

/*
 * call-seq:
 *   WIN32OLE_VARIABLE#variable_kind
 *
 * Returns variable kind string.
 *
 *    tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'XlSheetType')
 *    variables = tobj.variables
 *    variables.each do |variable|
 *      puts "#{variable.name} #{variable.variable_kind}"
 *    end
 *
 *    The result of above script is following:
 *      xlChart CONSTANT
 *      xlDialogSheet CONSTANT
 *      xlExcel4IntlMacroSheet CONSTANT
 *      xlExcel4MacroSheet CONSTANT
 *      xlWorksheet CONSTANT
 */
static mrb_value
folevariable_variable_kind(mrb_state *mrb, mrb_value self)
{
    struct olevariabledata *pvar;
    Data_Get_Struct(mrb, self, &olevariable_datatype, pvar);
    return ole_variable_kind(mrb, pvar->pTypeInfo, pvar->index);
}

static mrb_value
ole_variable_varkind(mrb_state *mrb, ITypeInfo *pTypeInfo, UINT var_index)
{
    VARDESC *pVarDesc;
    HRESULT hr;
    mrb_value kind = mrb_nil_value();
    hr = pTypeInfo->lpVtbl->GetVarDesc(pTypeInfo, var_index, &pVarDesc);
    if (FAILED(hr))
        return kind;
    pTypeInfo->lpVtbl->ReleaseVarDesc(pTypeInfo, pVarDesc);
    kind = INT2FIX(pVarDesc->varkind);
    return kind;
}

/*
 *  call-seq:
 *     WIN32OLE_VARIABLE#varkind
 *
 *  Returns the number which represents variable kind.
 *    tobj = WIN32OLE_TYPE.new('Microsoft Excel 9.0 Object Library', 'XlSheetType')
 *    variables = tobj.variables
 *    variables.each do |variable|
 *      puts "#{variable.name} #{variable.varkind}"
 *    end
 *
 *    The result of above script is following:
 *       xlChart 2
 *       xlDialogSheet 2
 *       xlExcel4IntlMacroSheet 2
 *       xlExcel4MacroSheet 2
 *       xlWorksheet 2
 */
static mrb_value
folevariable_varkind(mrb_state *mrb, mrb_value self)
{
    struct olevariabledata *pvar;
    Data_Get_Struct(mrb, self, &olevariable_datatype, pvar);
    return ole_variable_varkind(mrb, pvar->pTypeInfo, pvar->index);
}

/*
 *  call-seq:
 *     WIN32OLE_VARIABLE#inspect -> String
 *
 *  Returns the OLE variable name and the value with class name.
 *
 */
static mrb_value
folevariable_inspect(mrb_state *mrb, mrb_value self)
{
    mrb_value v = mrb_inspect(mrb, folevariable_value(mrb, self));
    mrb_value n = folevariable_name(mrb, self);
    mrb_value detail = mrb_format(mrb, "%S=%S", n, v);
    return make_inspect(mrb, "WIN32OLE_VARIABLE", detail);
}

void Init_win32ole_variable(mrb_state *mrb)
{
    struct RClass *cWIN32OLE_VARIABLE = mrb_define_class(mrb, "WIN32OLE_VARIABLE", mrb->object_class);
    MRB_SET_INSTANCE_TT(cWIN32OLE_VARIABLE, MRB_TT_DATA);
    mrb_define_method(mrb, cWIN32OLE_VARIABLE, "name", folevariable_name, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_VARIABLE, "ole_type", folevariable_ole_type, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_VARIABLE, "ole_type_detail", folevariable_ole_type_detail, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_VARIABLE, "value", folevariable_value, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_VARIABLE, "visible?", folevariable_visible, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_VARIABLE, "variable_kind", folevariable_variable_kind, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_VARIABLE, "varkind", folevariable_varkind, MRB_ARGS_NONE());
    mrb_define_method(mrb, cWIN32OLE_VARIABLE, "inspect", folevariable_inspect, MRB_ARGS_NONE());
    mrb_define_alias(mrb, cWIN32OLE_VARIABLE, "to_s", "name");
}
