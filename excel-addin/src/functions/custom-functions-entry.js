import { loadDetail } from "./load-detail.js";

if (globalThis.CustomFunctions?.associate) {
  globalThis.CustomFunctions.associate("LOAD_DETAIL", loadDetail);
}
