#include <napi.h>
#include <xlsxwriter.h>

#include "chart.h"
#include "error.h"
#include "format.h"
#include "workbook.h"
#include "worksheet.h"

Napi::Object Workbook::Init(Napi::Env env, Napi::Object exports) {
  Napi::Function func = DefineClass(
      env,
      "Workbook",
      {
          InstanceMethod<&Workbook::AddChart>("addChart", napi_default_method),
          InstanceMethod<&Workbook::AddFormat>("addFormat",
                                               napi_default_method),
          InstanceMethod<&Workbook::AddWorksheet>("addWorksheet",
                                                  napi_default_method),
          InstanceAccessor<&Workbook::GetDefaultURLFormat>("defaultURLFormat",
                                                           napi_configurable),
          InstanceMethod<&Workbook::Close>("close", napi_default_method),
      });
  exports["Workbook"] = func;
  return exports;
}

Workbook::Workbook(const Napi::CallbackInfo& info)
    : Napi::ObjectWrap<Workbook>(info) {
  lxw_workbook_options options = {};
  options.output_buffer = const_cast<const char**>(&output_buffer);
  options.output_buffer_size = &output_buffer_size;
  workbook = workbook_new_opt(NULL, &options);
  defaultURLFormat = Napi::Persistent(
      Format::New(info.Env(), workbook_get_default_url_format(workbook))
          .As<Napi::Object>());
}

Workbook::~Workbook() {
  if (workbook) lxw_workbook_free(workbook);
}

Napi::Value Workbook::AddChart(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  return Chart::New(
      env,
      workbook_add_chart(workbook, info[0].As<Napi::Number>().Uint32Value()));
}

Napi::Value Workbook::AddFormat(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  return Format::New(env, workbook_add_format(workbook));
}

Napi::Value Workbook::AddWorksheet(const Napi::CallbackInfo& info) {
  auto name = info[0].As<Napi::String>();
  auto env = info.Env();
  return Worksheet::New(
      env, workbook_add_worksheet(workbook, name.Utf8Value().c_str()));
}

Napi::Value Workbook::GetDefaultURLFormat(const Napi::CallbackInfo& info) {
  return defaultURLFormat.Value();
}

Napi::Value Workbook::Close(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  checkError(env, workbook_close(workbook));
  workbook = nullptr;
  return Napi::ArrayBuffer::New(
      env, output_buffer, output_buffer_size, [](Napi::Env env, void* data) {
        free(data);
      });
}
