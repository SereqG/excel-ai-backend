import logging

from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt

logger = logging.getLogger(__name__)


@csrf_exempt
def pipeline_execution(request):
    """
    Simple endpoint that logs the incoming request body.

    WARNING: Logging request bodies can leak secrets/PII into logs.
    """
    if request.method != "POST":
        return JsonResponse({"detail": "Method not allowed"}, status=405)

    raw_body = request.body or b""
    body_text = raw_body.decode("utf-8", errors="replace")

    # Avoid dumping huge payloads into logs.
    max_log_chars = 10_000
    truncated = body_text[:max_log_chars]
    was_truncated = len(body_text) > max_log_chars

    # Print to server console (stdout) as requested.
    print(
        f"pipeline execution: content_type={request.content_type} "
        f"bytes={len(raw_body)} truncated={was_truncated} body={truncated}"
    )

    logger.info(
        "pipeline execution: content_type=%s bytes=%d truncated=%s body=%s",
        request.content_type,
        len(raw_body),
        was_truncated,
        truncated,
    )

    return JsonResponse(
        {
            "endpoint": "pipeline execution",
            "received_bytes": len(raw_body),
            "logged": True,
            "truncated_in_logs": was_truncated,
        }
    )
