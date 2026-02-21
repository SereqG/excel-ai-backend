from django.db import models
import uuid
from django.utils import timezone


class PipelineJob(models.Model):
    """
    Immutable pipeline execution job.

    Stores the pipeline_operations as received so execution can be:
    - immutable once started
    - auditable
    - securely enforced by ownership (clerk_user_id)
    """

    class Status(models.TextChoices):
        PENDING = "PENDING", "Pending"
        RUNNING = "RUNNING", "Running"
        SUCCEEDED = "SUCCEEDED", "Succeeded"
        FAILED = "FAILED", "Failed"

    job_id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    file = models.ForeignKey(
        "file_manager.ProcessedFile",
        on_delete=models.CASCADE,
        related_name="pipeline_jobs",
    )
    clerk_user_id = models.CharField(max_length=255, db_index=True)

    pipeline_operations = models.JSONField()

    status = models.CharField(
        max_length=20,
        choices=Status.choices,
        default=Status.PENDING,
        db_index=True,
    )
    error = models.TextField(null=True, blank=True)

    celery_task_id = models.CharField(max_length=255, null=True, blank=True, db_index=True)

    # Stored under MEDIA_ROOT/processed/...
    output_file = models.FileField(upload_to="processed/", null=True, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    started_at = models.DateTimeField(null=True, blank=True)
    finished_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        ordering = ["-created_at"]
        indexes = [
            models.Index(fields=["clerk_user_id", "created_at"]),
            models.Index(fields=["status", "created_at"]),
        ]

    def mark_running(self, task_id: str | None = None) -> None:
        self.status = self.Status.RUNNING
        self.started_at = timezone.now()
        if task_id:
            self.celery_task_id = task_id
        self.save(update_fields=["status", "started_at", "celery_task_id", "updated_at"])

    def mark_succeeded(self) -> None:
        self.status = self.Status.SUCCEEDED
        self.finished_at = timezone.now()
        self.save(update_fields=["status", "finished_at", "updated_at"])

    def mark_failed(self, error: str) -> None:
        self.status = self.Status.FAILED
        self.error = error
        self.finished_at = timezone.now()
        self.save(update_fields=["status", "error", "finished_at", "updated_at"])
