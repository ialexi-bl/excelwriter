#include <napi.h>
#include <xlsxwriter.h>

#include "chart.h"
#include "format.h"
#include "workbook.h"
#include "worksheet.h"

Napi::Object Init(Napi::Env env, Napi::Object exports) {
  Chart::Init(env, exports);
  Format::Init(env, exports);
  Workbook::Init(env, exports);
  Worksheet::Init(env, exports);

  {
    auto colors = Napi::Object::New(env);
    colors["BLACK_COLOR"] = Napi::Number::New(env, LXW_COLOR_BLACK);
    colors["BLUE_COLOR"] = Napi::Number::New(env, LXW_COLOR_BLUE);
    colors["BROWN_COLOR"] = Napi::Number::New(env, LXW_COLOR_BROWN);
    colors["CYAN_COLOR"] = Napi::Number::New(env, LXW_COLOR_CYAN);
    colors["GRAY_COLOR"] = Napi::Number::New(env, LXW_COLOR_GRAY);
    colors["GREEN_COLOR"] = Napi::Number::New(env, LXW_COLOR_GREEN);
    colors["LIME_COLOR"] = Napi::Number::New(env, LXW_COLOR_LIME);
    colors["MAGENTA_COLOR"] = Napi::Number::New(env, LXW_COLOR_MAGENTA);
    colors["NAVY_COLOR"] = Napi::Number::New(env, LXW_COLOR_NAVY);
    colors["ORANGE_COLOR"] = Napi::Number::New(env, LXW_COLOR_ORANGE);
    colors["PINK_COLOR"] = Napi::Number::New(env, LXW_COLOR_PINK);
    colors["PURPLE_COLOR"] = Napi::Number::New(env, LXW_COLOR_PURPLE);
    colors["RED_COLOR"] = Napi::Number::New(env, LXW_COLOR_RED);
    colors["SILVER_COLOR"] = Napi::Number::New(env, LXW_COLOR_SILVER);
    colors["WHITE_COLOR"] = Napi::Number::New(env, LXW_COLOR_WHITE);
    colors["YELLOW_COLOR"] = Napi::Number::New(env, LXW_COLOR_YELLOW);
    exports["Color"] = colors;
  }

  {
    auto validationTypes = Napi::Object::New(env);
    validationTypes["NONE"] = Napi::Number::New(env, LXW_VALIDATION_TYPE_NONE);
    validationTypes["INTEGER"] =
        Napi::Number::New(env, LXW_VALIDATION_TYPE_INTEGER);
    validationTypes["INTEGER_FORMULA"] =
        Napi::Number::New(env, LXW_VALIDATION_TYPE_INTEGER_FORMULA);
    validationTypes["DECIMAL"] =
        Napi::Number::New(env, LXW_VALIDATION_TYPE_DECIMAL);
    validationTypes["DECIMAL_FORMULA"] =
        Napi::Number::New(env, LXW_VALIDATION_TYPE_DECIMAL_FORMULA);
    validationTypes["LIST"] = Napi::Number::New(env, LXW_VALIDATION_TYPE_LIST);
    validationTypes["LIST_FORMULA"] =
        Napi::Number::New(env, LXW_VALIDATION_TYPE_LIST_FORMULA);
    validationTypes["DATE"] = Napi::Number::New(env, LXW_VALIDATION_TYPE_DATE);
    validationTypes["DATE_FORMULA"] =
        Napi::Number::New(env, LXW_VALIDATION_TYPE_DATE_FORMULA);
    validationTypes["DATE_NUMBER"] =
        Napi::Number::New(env, LXW_VALIDATION_TYPE_DATE_NUMBER);
    validationTypes["TIME"] = Napi::Number::New(env, LXW_VALIDATION_TYPE_TIME);
    validationTypes["TIME_FORMULA"] =
        Napi::Number::New(env, LXW_VALIDATION_TYPE_TIME_FORMULA);
    validationTypes["TIME_NUMBER"] =
        Napi::Number::New(env, LXW_VALIDATION_TYPE_TIME_NUMBER);
    validationTypes["LENGTH"] =
        Napi::Number::New(env, LXW_VALIDATION_TYPE_LENGTH);
    validationTypes["LENGTH_FORMULA"] =
        Napi::Number::New(env, LXW_VALIDATION_TYPE_LENGTH_FORMULA);
    validationTypes["CUSTOM_FORMULA"] =
        Napi::Number::New(env, LXW_VALIDATION_TYPE_CUSTOM_FORMULA);
    validationTypes["ANY"] = Napi::Number::New(env, LXW_VALIDATION_TYPE_ANY);
    exports["ValidationType"] = validationTypes;
  }

  {
    auto validationCriteria = Napi::Object::New(env);
    validationCriteria["NONE"] =
        Napi::Number::New(env, LXW_VALIDATION_CRITERIA_NONE);
    validationCriteria["BETWEEN"] =
        Napi::Number::New(env, LXW_VALIDATION_CRITERIA_BETWEEN);
    validationCriteria["NOT_BETWEEN"] =
        Napi::Number::New(env, LXW_VALIDATION_CRITERIA_NOT_BETWEEN);
    validationCriteria["EQUAL_TO"] =
        Napi::Number::New(env, LXW_VALIDATION_CRITERIA_EQUAL_TO);
    validationCriteria["NOT_EQUAL_TO"] =
        Napi::Number::New(env, LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO);
    validationCriteria["GREATER_THAN"] =
        Napi::Number::New(env, LXW_VALIDATION_CRITERIA_GREATER_THAN);
    validationCriteria["LESS_THAN"] =
        Napi::Number::New(env, LXW_VALIDATION_CRITERIA_LESS_THAN);
    validationCriteria["GREATER_THAN_OR_EQUAL_TO"] = Napi::Number::New(
        env, LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO);
    validationCriteria["LESS_THAN_OR_EQUAL_TO"] =
        Napi::Number::New(env, LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO);
    exports["ValidationCriteria"] = validationCriteria;
  }

  {
    auto validationErrorTypes = Napi::Object::New(env);
    validationErrorTypes["STOP"] =
        Napi::Number::New(env, LXW_VALIDATION_ERROR_TYPE_STOP);
    validationErrorTypes["WARNING"] =
        Napi::Number::New(env, LXW_VALIDATION_ERROR_TYPE_WARNING);
    validationErrorTypes["INFORMATION"] =
        Napi::Number::New(env, LXW_VALIDATION_ERROR_TYPE_INFORMATION);
    exports["ValidationErrorType"] = validationErrorTypes;
  }

  return exports;
}

NODE_API_MODULE(NODE_GYP_MODULE_NAME, Init)
