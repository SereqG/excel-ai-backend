from __future__ import annotations

import io
import os
from pathlib import Path

from celery import shared_task
from django.core.files.base import ContentFile
from django.utils import timezone

from .models import PipelineJob
from .services.validation import validate_pipeline_operations
from .services.operations import apply_drop_column, apply_rename_column
from .services.workbook import extract_selected_sheet_workbook


@shared_task(bind=True)
def execute_pipeline(self, job_id: str):
    """
    Execute a stored pipeline job.

    Loads the original workbook from ProcessedFile.original_file.path, extracts the selected sheet,
    applies operations sequentially in-memory, and only writes output_file if the full pipeline succeeds.
    """
    job: PipelineJob | None = None
    try:
        job = PipelineJob.objects.select_related("file").get(job_id=job_id)
        file_record = job.file

        # Mark running early for visibility.
        task_id = getattr(self.request, "id", None) or job.celery_task_id
        job.mark_running(task_id=task_id)

        ops = validate_pipeline_operations(job.pipeline_operations)

        wb = extract_selected_sheet_workbook(
            original_path=file_record.original_file.path,
            selected_sheet=file_record.selected_sheet,
        )
        ws = wb[file_record.selected_sheet]

        header_row_idx = 1
        total = len(ops)

        for i, op in enumerate(ops):
            self.update_state(
                state="PROGRESS",
                meta={
                    "index": i + 1,
                    "total": total,
                    "op": {"id": op["id"], "operationId": op["operationId"]},
                    "phase": "start",
                },
            )

            operation_id = op["operationId"]
            params = op["params"]

            if operation_id == "rename_column":
                apply_rename_column(
                    ws,
                    header_row_idx=header_row_idx,
                    selected_sheet=file_record.selected_sheet,
                    column_id=params["columnId"],
                    new_name=params["newName"],
                )
            elif operation_id == "drop_column":
                apply_drop_column(
                    ws,
                    header_row_idx=header_row_idx,
                    selected_sheet=file_record.selected_sheet,
                    column_ids=params["columnIds"],
                )
            else:
                # Should be unreachable due to strict validation
                raise ValueError(f"Unsupported operationId: {operation_id}")

            self.update_state(
                state="PROGRESS",
                meta={
                    "index": i + 1,
                    "total": total,
                    "op": {"id": op["id"], "operationId": op["operationId"]},
                    "phase": "done",
                },
            )

        # Save output only after all ops succeed.
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        original_name = os.path.basename(file_record.original_file.name)
        base = Path(original_name).stem or str(file_record.file_id)
        timestamp = timezone.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{base}_{timestamp}.xlsx"

        job.output_file.save(output_filename, ContentFile(buf.read()), save=False)
        job.error = None
        job.save(update_fields=["output_file", "error", "updated_at"])
        job.mark_succeeded()

        return {"job_id": str(job.job_id), "output_file": job.output_file.name}

    except Exception as exc:
        if job is not None:
            job.mark_failed(str(exc))
        raise

