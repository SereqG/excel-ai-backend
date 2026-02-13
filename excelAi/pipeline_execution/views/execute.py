from rest_framework import status
from rest_framework.response import Response
from rest_framework.views import APIView

from file_manager.exceptions import AuthenticationError, FileNotFoundError, RequestValidationError
from file_manager.services import check_authentication, get_file_by_id

from ..models import PipelineJob
from ..services.errors import PipelineValidationError
from ..services.validation import validate_pipeline_operations
from ..tasks import execute_pipeline


class PipelineExecuteView(APIView):
    """
    Start a pipeline execution job.

    POST body:
    - file_id: UUID (ProcessedFile.file_id)
    - pipeline_operations: list (may be empty)
    """

    def post(self, request):
        try:
            clerk_user_id = check_authentication(request)

            file_id = request.data.get("file_id")
            pipeline_operations = request.data.get("pipeline_operations", [])

            if not file_id:
                raise RequestValidationError("file_id is required")

            validated_ops = validate_pipeline_operations(pipeline_operations)

            file_record = get_file_by_id(str(file_id), clerk_user_id)

            job = PipelineJob.objects.create(
                file=file_record,
                clerk_user_id=clerk_user_id,
                pipeline_operations=validated_ops,
                status=PipelineJob.Status.PENDING,
            )

            async_result = execute_pipeline.delay(str(job.job_id))
            job.celery_task_id = async_result.id
            job.save(update_fields=["celery_task_id", "updated_at"])

            job_id = str(job.job_id)
            return Response(
                {
                    "job_id": job_id,
                    "task_id": async_result.id,
                    "stream_url": f"/api/pipeline/execution/{job_id}/stream/",
                    "status_url": f"/api/pipeline/execution/{job_id}/status/",
                    "download_url": f"/api/pipeline/execution/{job_id}/download/",
                },
                status=status.HTTP_202_ACCEPTED,
            )

        except AuthenticationError as e:
            return Response({"error": str(e)}, status=status.HTTP_401_UNAUTHORIZED)
        except (RequestValidationError, PipelineValidationError) as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        except FileNotFoundError as e:
            return Response({"error": str(e)}, status=status.HTTP_404_NOT_FOUND)
        except Exception as e:
            return Response(
                {"error": f"Error starting pipeline execution: {str(e)}"},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR,
            )

