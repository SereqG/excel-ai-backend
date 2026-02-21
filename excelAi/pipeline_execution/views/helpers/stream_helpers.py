from __future__ import annotations

from celery.result import AsyncResult

from ...models import PipelineJob
from ..sse import sse_event


def connected_event(*, job: PipelineJob) -> str:
    return sse_event(
        event="connected",
        data={
            "job_id": str(job.job_id),
            "task_id": job.celery_task_id,
            "status": job.status,
        },
    )


def _job_status_event(*, job: PipelineJob) -> str:
    return sse_event(
        event="job_status",
        data={"job_id": str(job.job_id), "status": job.status, "error": job.error},
    )


def succeeded_event(*, job: PipelineJob, download_url: str) -> str:
    return sse_event(
        event="succeeded",
        data={
            "job_id": str(job.job_id),
            "status": job.status,
            "download_url": download_url,
        },
    )


def failed_event(*, job: PipelineJob) -> str:
    return sse_event(
        event="failed",
        data={
            "job_id": str(job.job_id),
            "status": job.status,
            "error": job.error,
        },
    )


def _task_state_event(*, job: PipelineJob, state: str, meta: dict | None) -> str:
    return sse_event(
        event="task_state",
        data={
            "job_id": str(job.job_id),
            "task_id": job.celery_task_id,
            "state": state,
            "meta": meta,
        },
    )


def _get_task_snapshot(task_id: str) -> tuple[str, dict | None]:
    res = AsyncResult(task_id)
    task_state = res.state
    task_meta = res.info if isinstance(res.info, dict) else None
    return task_state, task_meta


def maybe_emit_job_status_change(
    *,
    job: PipelineJob,
    last_job_status: str | None,
) -> tuple[str | None, str | None]:
    """
    Returns: (event_or_none, new_last_job_status)
    """
    if job.status == last_job_status:
        return None, last_job_status
    return _job_status_event(job=job), job.status


def maybe_emit_task_state_change(
    *,
    job: PipelineJob,
    last_task_state: str | None,
    last_task_meta: dict | None,
) -> tuple[str | None, str | None, dict | None]:
    """
    Returns: (event_or_none, new_last_task_state, new_last_task_meta)
    """
    if not job.celery_task_id:
        return None, last_task_state, last_task_meta

    task_state, task_meta = _get_task_snapshot(job.celery_task_id)
    if task_state == last_task_state and task_meta == last_task_meta:
        return None, last_task_state, last_task_meta

    return _task_state_event(job=job, state=task_state, meta=task_meta), task_state, task_meta

