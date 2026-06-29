type ErrorLike = {
  message?: unknown;
  details?: unknown;
  hint?: unknown;
  code?: unknown;
  status?: unknown;
  statusText?: unknown;
};

function isErrorLike(error: unknown): error is ErrorLike {
  return Boolean(error && typeof error === "object");
}

export function formatPosError(error: unknown, fallback = "Something went wrong.") {
  if (error instanceof Error) return error.message || fallback;
  if (typeof error === "string") return error || fallback;
  if (!isErrorLike(error)) return fallback;

  const parts = [error.message, error.details, error.hint, error.code ? `Code: ${error.code}` : null]
    .filter((part): part is string => typeof part === "string" && part.trim().length > 0)
    .map((part) => part.trim());

  return parts.length ? parts.join(" ") : fallback;
}

export function logPosError(context: string, error: unknown) {
  console.error(`[CJNET POS] ${context}`, error);
}
