from django.urls import path

from .views import pipeline_execution

urlpatterns = [
    path("execution/", pipeline_execution, name="pipeline execution"),
]

