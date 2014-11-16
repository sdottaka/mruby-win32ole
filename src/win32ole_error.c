#include "win32ole.h"

static mrb_value ole_hresult2msg(mrb_state *mrb, HRESULT hr);

static mrb_value
ole_hresult2msg(mrb_state *mrb, HRESULT hr)
{
    mrb_value msg = mrb_nil_value();
    char *p_msg = NULL;
    char *term = NULL;
    DWORD dwCount;

    char strhr[100];
    sprintf(strhr, "    HRESULT error code:0x%08x\n      ", (unsigned)hr);
    msg = mrb_str_new_cstr(mrb, strhr);
    dwCount = FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER |
                            FORMAT_MESSAGE_FROM_SYSTEM |
                            FORMAT_MESSAGE_IGNORE_INSERTS,
                            NULL, hr,
                            MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US),
                            (LPTSTR)&p_msg, 0, NULL);
    if (dwCount == 0) {
        dwCount = FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER |
                                FORMAT_MESSAGE_FROM_SYSTEM |
                                FORMAT_MESSAGE_IGNORE_INSERTS,
                                NULL, hr, cWIN32OLE_lcid,
                                (LPTSTR)&p_msg, 0, NULL);
    }
    if (dwCount > 0) {
        term = p_msg + strlen(p_msg);
        while (p_msg < term) {
            term--;
            if (*term == '\r' || *term == '\n')
                *term = '\0';
            else break;
        }
        if (p_msg[0] != '\0') {
            mrb_str_cat_cstr(mrb, msg, p_msg);
        }
    }
    LocalFree(p_msg);
    return msg;
}

void
ole_raise(mrb_state *mrb, HRESULT hr, struct RClass *ecs, const char *fmt, ...)
{
    va_list args;
    mrb_value msg;
    mrb_value err_msg;
    char cmsg[1024];
    va_init_list(args, fmt);
    vsprintf(cmsg, fmt, args);
    va_end(args);
    msg = mrb_str_new_cstr(mrb, cmsg);

    err_msg = ole_hresult2msg(mrb, hr);
    if(!mrb_nil_p(err_msg)) {
        mrb_str_cat_lit(mrb, msg, "\n");
        mrb_str_append(mrb, msg, err_msg);
    }
    mrb_exc_raise(mrb, mrb_exc_new_str(mrb, ecs, msg));
}

void
Init_win32ole_error(mrb_state *mrb)
{
    /*
     * Document-class: WIN32OLERuntimeError
     *
     * Raised when OLE processing failed.
     *
     * EX:
     *
     *   obj = WIN32OLE.new("NonExistProgID")
     *
     * raises the exception:
     *
     *   WIN32OLERuntimeError: unknown OLE server: `NonExistProgID'
     *       HRESULT error code:0x800401f3
     *         Invalid class string
     *
     */
    mrb_define_class(mrb, "WIN32OLERuntimeError", E_RUNTIME_ERROR);
}
