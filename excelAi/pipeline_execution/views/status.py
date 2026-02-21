from __future__ import annotations

from rest_framework import status
from rest_framework.response import Response
from rest_framework.views import APIView

from file_manager.exceptions import AuthenticationError, FileNotFoundError
from file_manager.services import check_authentication

from ..models import PipelineJob


class PipelineStatusView(APIView):
    """
    Get status for a pipeline job.
    """

    def get(self, request, job_id):
        try:
            clerk_user_id = check_authentication(request)
            job = PipelineJob.objects.get(job_id=job_id)
            if job.clerk_user_id != clerk_user_id:
                raise FileNotFoundError("Job not found")

            payload = {
                "job_id": str(job.job_id),
                "status": job.status,
                "error": job.error,
                "task_id": job.celery_task_id,
                "has_output": bool(job.output_file),
            }
            if job.status == PipelineJob.Status.SUCCEEDED:
                payload["download_url"] = f"/api/pipeline/execution/{job_id}/download/"

            return Response(payload, status=status.HTTP_200_OK)

        except AuthenticationError as e:
            return Response({"error": str(e)}, status=status.HTTP_401_UNAUTHORIZED)
        except PipelineJob.DoesNotExist:
            return Response({"error": "Job not found"}, status=status.HTTP_404_NOT_FOUND)
        except FileNotFoundError as e:
            return Response({"error": str(e)}, status=status.HTTP_404_NOT_FOUND)
        except Exception as e:
            return Response(
                {"error": f"Error retrieving job status: {str(e)}"},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR,
            )

