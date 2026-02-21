from django.contrib import admin

from .models import PipelineJob


@admin.register(PipelineJob)
class PipelineJobAdmin(admin.ModelAdmin):
    list_display = (
        "job_id",
        "clerk_user_id",
        "file",
        "status",
        "created_at",
        "started_at",
        "finished_at",
    )
    list_filter = ("status", "created_at")
    search_fields = ("job_id", "clerk_user_id", "celery_task_id")
    readonly_fields = ("created_at", "updated_at", "started_at", "finished_at")
