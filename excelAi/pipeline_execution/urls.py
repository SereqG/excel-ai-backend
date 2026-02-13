from django.urls import path

from .views import (
    PipelineDownloadView,
    PipelineExecuteView,
    PipelineStatusView,
    pipeline_stream,
)

urlpatterns = [
    path("execution/", PipelineExecuteView.as_view(), name="pipeline-execute"),
    path("execution/<uuid:job_id>/status/", PipelineStatusView.as_view(), name="pipeline-status"),
    path("execution/<uuid:job_id>/stream/", pipeline_stream, name="pipeline-stream"),
    path("execution/<uuid:job_id>/download/", PipelineDownloadView.as_view(), name="pipeline-download"),
]

