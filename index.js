/** @type {import('.')} */ // @ts-expect-error
const excelwriter = require("./build/Release/xlsxwriter");

module.exports.Chart = excelwriter.Chart;
module.exports.Color = excelwriter.Color;
module.exports.Format = excelwriter.Format;
module.exports.Workbook = excelwriter.Workbook;
module.exports.ValidationType = excelwriter.ValidationType;
module.exports.ValidationCriteria = excelwriter.ValidationCriteria;
module.exports.ValidationErrorType = excelwriter.ValidationErrorType;
