import assert from "node:assert/strict";
import test from "node:test";

import {
  buildClientLogContext,
  normalizePeriodEndDate,
  parseLoadDetailValue,
} from "./load-detail.js";

test("normalizePeriodEndDate accepts an Excel date serial", () => {
  assert.equal(normalizePeriodEndDate(46263), "2026-08-31");
});

test("normalizePeriodEndDate accepts an ISO date string and returns month end", () => {
  assert.equal(normalizePeriodEndDate("2026-05-15"), "2026-05-31");
});

test("parseLoadDetailValue returns zero when backend has no stored row", () => {
  assert.equal(parseLoadDetailValue({ value: null }), 0);
});

test("buildClientLogContext captures diagnostic fields without payroll values", () => {
  assert.deepEqual(
    buildClientLogContext(
      " payroll.output.401k ",
      "30/Apr/26",
      " EX18 ",
      "user@example.com"
    ),
    {
      outputKey: "payroll.output.401k",
      periodType: "string",
      periodRaw: "30/Apr/26",
      unitId: "EX18",
      userKeyOverrideProvided: true,
    }
  );
});
