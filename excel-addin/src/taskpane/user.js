export function normalizeUserKey(value) {
  return String(value ?? "").trim().toLowerCase();
}
