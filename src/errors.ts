/**
 * Extracts the diagnostic fields worth logging from an unknown throw.
 * `err.message` alone drops the response body and the provider error code --
 * exactly the information needed when a token exchange or Graph call fails.
 */
export function describeError(err: unknown): Record<string, unknown> {
  const e = err as any;

  const described: Record<string, unknown> = {
    kind: e?.constructor?.name ?? typeof err,
    message: e?.message,
  };

  // GraphError (@microsoft/teams.graph)
  if (e?.statusCode !== undefined) described.statusCode = e.statusCode;
  if (e?.code !== undefined) described.code = e.code;
  if (e?.body !== undefined) described.body = e.body;

  // AxiosError
  if (e?.response) {
    described.httpStatus = e.response.status;
    described.responseData = e.response.data;
  }

  // Anthropic SDK errors put the status on `status` (not `statusCode`), the
  // parsed body on `error`, and a request id that Anthropic support asks for.
  if (described.httpStatus === undefined && e?.status !== undefined) {
    described.httpStatus = e.status;
  }
  if (e?.error !== undefined) described.errorBody = e.error;
  if (e?.request_id !== undefined) described.requestId = e.request_id;

  // Graph and the Bot Framework token service both hide the real reason here.
  const inner = e?.body?.error?.innerError ?? e?.response?.data?.error?.innerError;
  if (inner) described.innerError = inner;

  if (e?.source) described.source = e.source;
  if (e?.stack) described.stack = e.stack;

  return described;
}
