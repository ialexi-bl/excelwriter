#include <napi.h>
#include <xlsxwriter.h>

#include "chart.h"
#include "format.h"
#include "worksheet.h"

Napi::Object Worksheet::Init(Napi::Env env, Napi::Object exports) {
  auto func = DefineClass(
      env,
      "Worksheet",
      {
          InstanceMethod<&Worksheet::FreezePanes>("freezePanes",
                                                  napi_default_method),
          InstanceMethod<&Worksheet::SplitPanes>("splitPanes",
                                                 napi_default_method),
          InstanceMethod<&Worksheet::InsertChart>("insertChart",
                                                  napi_default_method),
          InstanceMethod<&Worksheet::InsertImage>("insertImage",
                                                  napi_default_method),
          InstanceMethod<&Worksheet::MergeRange>("mergeRange",
                                                 napi_default_method),
          InstanceMethod<&Worksheet::SetColumn>("setColumn",
                                                napi_default_method),
          InstanceMethod<&Worksheet::SetRow>("setRow", napi_default_method),
          InstanceMethod<&Worksheet::SetFooter>("setFooter",
                                                napi_default_method),
          InstanceMethod<&Worksheet::SetHeader>("setHeader",
                                                napi_default_method),
          InstanceMethod<&Worksheet::SetSelection>("setSelection",
                                                   napi_default_method),
          InstanceMethod<&Worksheet::WriteBoolean>("writeBoolean",
                                                   napi_default_method),
          InstanceMethod<&Worksheet::WriteDatetime>("writeDatetime",
                                                    napi_default_method),
          InstanceMethod<&Worksheet::WriteFormula>("writeFormula",
                                                   napi_default_method),
          InstanceMethod<&Worksheet::WriteNumber>("writeNumber",
                                                  napi_default_method),
          InstanceMethod<&Worksheet::WriteString>("writeString",
                                                  napi_default_method),
          InstanceMethod<&Worksheet::WriteURL>("writeURL", napi_default_method),
          InstanceMethod<&Worksheet::WriteFormulaNum>("writeFormulaNum",
                                                      napi_default_method),
          InstanceMethod<&Worksheet::WriteFormulaStr>("writeFormulaStr",
                                                      napi_default_method),
          InstanceMethod<&Worksheet::Autofilter>("autofilter",
                                                 napi_default_method),
          InstanceMethod<&Worksheet::DataValidationCell>("dataValidationCell",
                                                         napi_default_method),
          InstanceMethod<&Worksheet::DataValidationRange>("dataValidationRange",
                                                          napi_default_method),
      });

  auto data = env.GetInstanceData<Napi::ObjectReference>();

  if (!data) {
    data = new Napi::ObjectReference();
    *data = Napi::Persistent(Napi::Object::New(env));
    env.SetInstanceData(data);
  }

  data->Set("WorksheetConstructor", func);

  return exports;
}

Worksheet::Worksheet(const Napi::CallbackInfo& info)
    : Napi::ObjectWrap<Worksheet>(info) {
  worksheet = info[0].As<Napi::External<lxw_worksheet>>().Data();
}

Napi::Value Worksheet::New(Napi::Env env, lxw_worksheet* worksheet) {
  return env.GetInstanceData<Napi::ObjectReference>()
      ->Get("WorksheetConstructor")
      .As<Napi::Function>()
      .New({Napi::External<lxw_worksheet>::New(env, worksheet)});
}

Napi::Value Worksheet::FreezePanes(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_freeze_panes(worksheet,
                         info[0].As<Napi::Number>(),
                         info[1].As<Napi::Number>().Uint32Value());
  return env.Undefined();
}

Napi::Value Worksheet::SplitPanes(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_split_panes(
      worksheet, info[0].As<Napi::Number>(), info[1].As<Napi::Number>());
  return env.Undefined();
}

Napi::Value Worksheet::InsertChart(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_insert_chart(worksheet,
                         info[0].As<Napi::Number>(),
                         info[1].As<Napi::Number>().Uint32Value(),
                         Chart::Get(info[2]));
  return env.Undefined();
}

Napi::Value Worksheet::InsertImage(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  auto buffer = info[2].As<Napi::Uint8Array>();
  worksheet_insert_image_buffer(worksheet,
                                info[0].As<Napi::Number>(),
                                info[1].As<Napi::Number>().Uint32Value(),
                                buffer.Data(),
                                buffer.ByteLength());
  return env.Undefined();
}

Napi::Value Worksheet::MergeRange(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_merge_range(worksheet,
                        info[0].As<Napi::Number>(),
                        info[1].As<Napi::Number>().Uint32Value(),
                        info[2].As<Napi::Number>(),
                        info[3].As<Napi::Number>().Uint32Value(),
                        info[4].As<Napi::String>().Utf8Value().c_str(),
                        Format::Get(info[5]));
  return env.Undefined();
}

Napi::Value Worksheet::SetColumn(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_set_column(worksheet,
                       info[0].As<Napi::Number>().Uint32Value(),
                       info[1].As<Napi::Number>().Uint32Value(),
                       info[2].As<Napi::Number>(),
                       Format::Get(info[3]));
  return env.Undefined();
}

Napi::Value Worksheet::SetRow(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_set_row(worksheet,
                    info[0].As<Napi::Number>(),
                    info[1].As<Napi::Number>(),
                    Format::Get(info[2]));
  return env.Undefined();
}

Napi::Value Worksheet::SetFooter(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_set_footer(worksheet,
                       info[0].As<Napi::String>().Utf8Value().c_str());
  return env.Undefined();
}

Napi::Value Worksheet::SetHeader(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_set_header(worksheet,
                       info[0].As<Napi::String>().Utf8Value().c_str());
  return env.Undefined();
}

Napi::Value Worksheet::SetSelection(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_set_selection(worksheet,
                          info[0].As<Napi::Number>(),
                          info[1].As<Napi::Number>().Uint32Value(),
                          info[2].As<Napi::Number>(),
                          info[3].As<Napi::Number>().Uint32Value());
  return env.Undefined();
}

Napi::Value Worksheet::WriteBoolean(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_write_boolean(worksheet,
                          info[0].As<Napi::Number>(),
                          info[1].As<Napi::Number>().Uint32Value(),
                          info[2].As<Napi::Boolean>(),
                          Format::Get(info[3]));
  return env.Undefined();
}

Napi::Value Worksheet::WriteDatetime(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  auto date = info[2].As<Napi::Object>();
  auto offset = date.Get("getTimezoneOffset")
                    .As<Napi::Function>()
                    .Call(date, {})
                    .As<Napi::Number>()
                    .Int32Value() *
                60;
  worksheet_write_unixtime(worksheet,
                           info[0].As<Napi::Number>(),
                           info[1].As<Napi::Number>().Uint32Value(),
                           info[2].As<Napi::Date>() / 1000 - offset,
                           Format::Get(info[3]));
  return env.Undefined();
}

Napi::Value Worksheet::WriteFormula(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_write_formula(worksheet,
                          info[0].As<Napi::Number>(),
                          info[1].As<Napi::Number>().Uint32Value(),
                          info[2].As<Napi::String>().Utf8Value().c_str(),
                          Format::Get(info[3]));
  return env.Undefined();
}

Napi::Value Worksheet::WriteNumber(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_write_number(worksheet,
                         info[0].As<Napi::Number>(),
                         info[1].As<Napi::Number>().Uint32Value(),
                         info[2].As<Napi::Number>(),
                         Format::Get(info[3]));
  return env.Undefined();
}

Napi::Value Worksheet::WriteString(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_write_string(worksheet,
                         info[0].As<Napi::Number>(),
                         info[1].As<Napi::Number>().Uint32Value(),
                         info[2].As<Napi::String>().Utf8Value().c_str(),
                         Format::Get(info[3]));
  return env.Undefined();
}

Napi::Value Worksheet::WriteURL(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_write_url(worksheet,
                      info[0].As<Napi::Number>(),
                      info[1].As<Napi::Number>().Uint32Value(),
                      info[2].As<Napi::String>().Utf8Value().c_str(),
                      Format::Get(info[3]));
  return env.Undefined();
}

Napi::Value Worksheet::WriteFormulaNum(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_write_formula_num(worksheet,
                              info[0].As<Napi::Number>(),
                              info[1].As<Napi::Number>().Uint32Value(),
                              info[2].As<Napi::String>().Utf8Value().c_str(),
                              Format::Get(info[3]),
                              info[4].As<Napi::Number>());
  return env.Undefined();
}

Napi::Value Worksheet::WriteFormulaStr(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_write_formula_str(worksheet,
                              info[0].As<Napi::Number>(),
                              info[1].As<Napi::Number>().Uint32Value(),
                              info[2].As<Napi::String>().Utf8Value().c_str(),
                              Format::Get(info[3]),
                              info[4].As<Napi::String>().Utf8Value().c_str());
  return env.Undefined();
}

Napi::Value Worksheet::Autofilter(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_autofilter(worksheet,
                       info[0].As<Napi::Number>(),
                       info[1].As<Napi::Number>().Uint32Value(),
                       info[2].As<Napi::Number>(),
                       info[3].As<Napi::Number>().Uint32Value());
  return env.Undefined();
}

Napi::Value Worksheet::DataValidationCell(const Napi::CallbackInfo& info) {
  auto env = info.Env();
  worksheet_data_validation_cell(
      worksheet,
      info[0].As<Napi::Number>(),
      info[1].As<Napi::Number>().Uint32Value(),
      info[2].As<Napi::External<lxw_data_validation>>().Data());
  return env.Undefined();
}

Napi::Value Worksheet::DataValidationRange(const Napi::CallbackInfo& info) {
  auto env = info.Env();

  worksheet_data_validation_range(worksheet,
                                  info[0].As<Napi::Number>(),
                                  info[1].As<Napi::Number>().Uint32Value(),
                                  info[2].As<Napi::Number>(),
                                  info[3].As<Napi::Number>().Uint32Value(),
                                  DataValidation(info[4]));

  return env.Undefined();
}

Napi::Number callDateMethod(const Napi::Object& date, const char* method) {
  return date.Get(method).As<Napi::Function>().Call({}).As<Napi::Number>();
}

lxw_datetime convertDate(const Napi::Object& date) {
  return {
      .year =
          static_cast<int>(callDateMethod(date, "getFullYear").Uint32Value()),
      .month =
          static_cast<int>(callDateMethod(date, "getMonth").Uint32Value() + 1),
      .day = static_cast<int>(callDateMethod(date, "getDate").Uint32Value()),
      .hour = static_cast<int>(callDateMethod(date, "getHours").Uint32Value()),
      .min = static_cast<int>(callDateMethod(date, "getMinutes").Uint32Value()),
      .sec = callDateMethod(date, "getSeconds").Uint32Value() +
             callDateMethod(date, "getMilliseconds").DoubleValue() / 1000.0,
  };
}

DataValidation::DataValidation(const Napi::Value& value) {
  const auto obj = value.As<Napi::Object>();

  if (!obj.Get("validate").IsUndefined()) {
    data_validation.validate =
        obj.Get("validate").As<Napi::Number>().Uint32Value();
  }
  if (!obj.Get("criteria").IsUndefined()) {
    data_validation.criteria =
        obj.Get("criteria").As<Napi::Number>().Uint32Value();
  }
  if (!obj.Get("ignoreBlank").IsUndefined()) {
    data_validation.ignore_blank =
        obj.Get("ignoreBlank").As<Napi::Boolean>().Value();
  }
  if (!obj.Get("showInput").IsUndefined()) {
    data_validation.show_input =
        obj.Get("showInput").As<Napi::Boolean>().Value();
  }
  if (!obj.Get("showError").IsUndefined()) {
    data_validation.show_error =
        obj.Get("showError").As<Napi::Boolean>().Value();
  }
  if (!obj.Get("errorType").IsUndefined()) {
    data_validation.error_type =
        obj.Get("errorType").As<Napi::Number>().Uint32Value();
  }
  if (!obj.Get("dropdown").IsUndefined()) {
    data_validation.dropdown = obj.Get("dropdown").As<Napi::Boolean>().Value();
  }

  if (!obj.Get("valueNumber").IsUndefined()) {
    data_validation.value_number =
        obj.Get("valueNumber").As<Napi::Number>().DoubleValue();
  }
  if (!obj.Get("valueFormula").IsUndefined()) {
    valueFormula = obj.Get("valueFormula").As<Napi::String>();
    data_validation.value_formula = valueFormula.c_str();
  }
  if (!obj.Get("valueList").IsUndefined()) {
    auto array = obj.Get("valueList").As<Napi::Array>();
    auto length = array.Length();

    valueList = std::move(std::make_unique<const char*[]>(length + 1));
    for (uint32_t i = 0; i < length; i++) {
      valueListVec.push_back(array.Get(i).As<Napi::String>());
    }
    for (uint32_t i = 0; i < length; i++) {
      valueList[i] = valueListVec[i].c_str();
    }
    valueList[length] = nullptr;

    data_validation.value_list = valueList.get();
  }
  if (!obj.Get("valueDatetime").IsUndefined()) {
    data_validation.value_datetime =
        convertDate(obj.Get("valueDatetime").As<Napi::Object>());
  }

  if (!obj.Get("minimumNumber").IsUndefined()) {
    data_validation.minimum_number =
        obj.Get("minimumNumber").As<Napi::Number>().DoubleValue();
  }
  if (!obj.Get("minimumFormula").IsUndefined()) {
    minimumFormula = obj.Get("minimumFormula").As<Napi::String>();
    data_validation.minimum_formula = minimumFormula.c_str();
  }
  if (!obj.Get("minimumDatetime").IsUndefined()) {
    data_validation.minimum_datetime =
        convertDate(obj.Get("minimumDatetime").As<Napi::Object>());
  }

  if (!obj.Get("maximumNumber").IsUndefined()) {
    data_validation.maximum_number =
        obj.Get("maximumNumber").As<Napi::Number>().DoubleValue();
  }
  if (!obj.Get("maximumFormula").IsUndefined()) {
    maximumFormula = obj.Get("maximumFormula").As<Napi::String>();
    data_validation.maximum_formula = maximumFormula.c_str();
  }
  if (!obj.Get("maximumDatetime").IsUndefined()) {
    data_validation.maximum_datetime =
        convertDate(obj.Get("maximumDatetime").As<Napi::Object>());
  }

  if (!obj.Get("inputTitle").IsUndefined()) {
    inputTitle = obj.Get("inputTitle").As<Napi::String>();
    data_validation.input_title = inputTitle.c_str();
  }
  if (!obj.Get("input_message").IsUndefined()) {
    inputMessage = obj.Get("input_message").As<Napi::String>();
    data_validation.input_message = inputMessage.c_str();
  }
  if (!obj.Get("errorTitle").IsUndefined()) {
    errorTitle = obj.Get("errorTitle").As<Napi::String>();
    data_validation.error_title = errorTitle.c_str();
  }
  if (!obj.Get("error_message").IsUndefined()) {
    errorMessage = obj.Get("error_message").As<Napi::String>();
    data_validation.error_message = errorMessage.c_str();
  }
}
