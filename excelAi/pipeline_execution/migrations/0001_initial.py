import uuid

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    initial = True

    dependencies = [
        ("file_manager", "0002_initial"),
    ]

    operations = [
        migrations.CreateModel(
            name="PipelineJob",
            fields=[
                ("job_id", models.UUIDField(default=uuid.uuid4, editable=False, primary_key=True, serialize=False)),
                ("clerk_user_id", models.CharField(db_index=True, max_length=255)),
                ("pipeline_operations", models.JSONField()),
                (
                    "status",
                    models.CharField(
                        choices=[("PENDING", "Pending"), ("RUNNING", "Running"), ("SUCCEEDED", "Succeeded"), ("FAILED", "Failed")],
                        db_index=True,
                        default="PENDING",
                        max_length=20,
                    ),
                ),
                ("error", models.TextField(blank=True, null=True)),
                ("celery_task_id", models.CharField(blank=True, db_index=True, max_length=255, null=True)),
                ("output_file", models.FileField(blank=True, null=True, upload_to="processed/")),
                ("created_at", models.DateTimeField(auto_now_add=True)),
                ("updated_at", models.DateTimeField(auto_now=True)),
                ("started_at", models.DateTimeField(blank=True, null=True)),
                ("finished_at", models.DateTimeField(blank=True, null=True)),
                (
                    "file",
                    models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name="pipeline_jobs", to="file_manager.processedfile"),
                ),
            ],
            options={
                "ordering": ["-created_at"],
                "indexes": [
                    models.Index(fields=["clerk_user_id", "created_at"], name="pipeline_ex_clerk_u_a4f0c0_idx"),
                    models.Index(fields=["status", "created_at"], name="pipeline_ex_status__2a4a67_idx"),
                ],
            },
        ),
    ]

