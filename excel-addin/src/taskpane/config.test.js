import assert from "node:assert/strict";
import test from "node:test";

import { getFilterOffset, parseConfig } from "./config.js";

test("parseConfig reads the named-range config structure by key", () => {
  const config = parseConfig([
    ["Section", "Type", "Key", "Description", "Value", "Value Type"],
    ["1. Timeline", "Assumption - constant", "model.last_actuals_date", "Last actuals date", "31/Mar/26", "Date"],
    ["1. Timeline", "Assumption - constant", "model.model_end_date", "Model end date", "30/Apr/28", "Date"],
    ["1. Timeline", "Assumption - constant", "model.financial_year_end_month", "Financial year end month", 4, "Number"],
    ["2. Payroll", "Assumption - constant", "payroll.benefits.medical.domestic", "Medical - Domestic", 2464, "Number"],
    ["2. Payroll", "Assumption - constant", "payroll.benefits.medical.international", "Medical - International", 2162, "Number"],
    ["2. Payroll", "Assumption - constant", "payroll.benefits.401k.domestic", "401k - Domestic", 432, "Number"],
    ["2. Payroll", "Assumption - constant", "payroll.benefits.401k.international", "401k - International", 501, "Number"],
    ["2. Payroll", "Assumption - constant", "payroll.benefits.other.domestic", "Other Benefits - Domestic", 157, "Number"],
    ["2. Payroll", "Assumption - constant", "payroll.benefits.other.international", "Other Benefits - International", 20, "Number"],
    ["2. Payroll", "Input Range", "payroll.filter_column", "Filter column", "PayrollData!R:R", "Reference - Column"],
    ["2. Payroll", "Input Range", "payroll.data_range", "Input data by Employees", "PayrollData!B5:R1531", "Reference - Cell Range"],
    ["2. Payroll", "Input Range Header", "payroll.headers_range", "Header for Input data by Employees", "PayrollData!B4:R4", "Reference - Cell Range"],
    ["2. Payroll", "Output Range", "payroll.output.headcount", "Headcount by Departments by Month", "HCA_Output!E4", "Reference - Starting Cell"],
    ["2. Payroll", "Output Range", "payroll.output.base_salary_total", "Base Salary by Departments by Month", "HCA_Output!E17", "Reference - Starting Cell"],
    ["2. Payroll", "Output Range", "payroll.output.base_salary_domestic", "Base Salary Domestic", "HCA_Output!E30", "Reference - Starting Cell"],
    ["2. Payroll", "Output Range", "payroll.output.base_salary_international", "Base Salary International", "HCA_Output!E44", "Reference - Starting Cell"],
    ["2. Payroll", "Output Range", "payroll.output.base_salary_cogs", "Base Salary COGS", "HCA_Output!E57", "Reference - Starting Cell"],
    ["2. Payroll", "Output Range", "payroll.output.medical", "Medical", "HCA_Output!E70", "Reference - Starting Cell"],
    ["2. Payroll", "Output Range", "payroll.output.401k", "401k", "HCA_Output!E83", "Reference - Starting Cell"],
    ["2. Payroll", "Output Range", "payroll.output.other_benefits", "Other Benefits", "HCA_Output!E96", "Reference - Starting Cell"],
  ]);

  assert.equal(config.payroll.dataLoadSheet, "PayrollData");
  assert.equal(config.payroll.cellRange, "B5:R1531");
  assert.equal(config.payroll.headers, "B4:R4");
  assert.equal(config.payroll.filterColumn, "R");
  assert.equal(config.output.sheet, "HCA_Output");
  assert.equal(config.output.headcountStartCell, "E4");
  assert.equal(config.output.otherBenefitsStartCell, "E96");
  assert.equal(config.assumptions.benefits.otherBenefits.international, 20);
  assert.equal(config.model.calculationMonths, 25);
  assert.equal(config.model.periods[0].date, "2026-04-30");
  assert.equal(config.model.periods.at(-1).date, "2028-04-30");
});

test("getFilterOffset handles a full column filter reference", () => {
  assert.equal(getFilterOffset("B5:R1531", "R"), 16);
});
