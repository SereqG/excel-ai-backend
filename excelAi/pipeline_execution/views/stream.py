from __future__ import annotations

import time

from django.http import JsonResponse, StreamingHttpResponse

from file_manager.exceptions import AuthenticationError
from file_manager.services import check_authentication

from ..models import PipelineJob
from .helpers.stream_helpers import (
    connected_event,
    failed_event,
    maybe_emit_job_status_change,
    maybe_emit_task_state_change,
    succeeded_event,
)


def pipeline_stream(request, job_id):
    """
    Stream pipeline progress as Server-Sent Events (SSE).

    Auth: X-Clerk-User-Id header, enforced by job ownership.
    """
    try:
        clerk_user_id = check_authentication(request)
    except AuthenticationError as e:
        return JsonResponse({"error": str(e)}, status=401)

    try:
        job = PipelineJob.objects.select_related("file").get(job_id=job_id)
    except PipelineJob.DoesNotExist:
        return JsonResponse({"error": "Job not found"}, status=404)

    if job.clerk_user_id != clerk_user_id:
        # Avoid leaking existence across users.
        return JsonResponse({"error": "Job not found"}, status=404)

    download_url = f"/api/pipeline/execution/{job_id}/download/"

    def event_stream():
        last_task_state = None
        last_task_meta = None
        last_job_status = None

        yield connected_event(job=job)

        while True:
            # Prefer DB status (works even when Celery runs eagerly without storing results).
            job.refresh_from_db(fields=["status", "error", "celery_task_id", "output_file"])
            maybe_job_event, last_job_status = maybe_emit_job_status_change(
                job=job,
                last_job_status=last_job_status,
            )
            if maybe_job_event:
                yield maybe_job_event

            if job.status == PipelineJob.Status.SUCCEEDED:
                yield succeeded_event(job=job, download_url=download_url)
                break

            if job.status == PipelineJob.Status.FAILED:
                yield failed_event(job=job)
                break

            # If we have a task id and a result backend, also stream PROGRESS meta.
            maybe_task_event, last_task_state, last_task_meta = maybe_emit_task_state_change(
                job=job,
                last_task_state=last_task_state,
                last_task_meta=last_task_meta,
            )
            if maybe_task_event:
                yield maybe_task_event

            time.sleep(0.5)

    resp = StreamingHttpResponse(event_stream(), content_type="text/event-stream")
    resp["Cache-Control"] = "no-cache"
    resp["X-Accel-Buffering"] = "no"
    return resp

