from __future__ import annotations

import os

from django.http import FileResponse
from rest_framework import status
from rest_framework.response import Response
from rest_framework.views import APIView

from file_manager.exceptions import AuthenticationError, FileNotFoundError
from file_manager.services import check_authentication

from ..models import PipelineJob


class PipelineDownloadView(APIView):
    """
    Download the resulting .xlsx for a successful job.
    """

    def get(self, request, job_id):
        try:
            clerk_user_id = check_authentication(request)
            job = PipelineJob.objects.select_related("file").get(job_id=job_id)
            if job.clerk_user_id != clerk_user_id:
                raise FileNotFoundError("Job not found")

            if job.status != PipelineJob.Status.SUCCEEDED or not job.output_file:
                return Response(
                    {"error": f"Job is not ready for download. Current status: {job.status}"},
                    status=status.HTTP_400_BAD_REQUEST,
                )

            filename = os.path.basename(job.output_file.name)
            resp = FileResponse(
                job.output_file.open("rb"),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            resp["Content-Disposition"] = f'attachment; filename="{filename}"'
            return resp

        except AuthenticationError as e:
            return Response({"error": str(e)}, status=status.HTTP_401_UNAUTHORIZED)
        except PipelineJob.DoesNotExist:
            return Response({"error": "Job not found"}, status=status.HTTP_404_NOT_FOUND)
        except FileNotFoundError as e:
            return Response({"error": str(e)}, status=status.HTTP_404_NOT_FOUND)
        except Exception as e:
            return Response(
                {"error": f"Error downloading file: {str(e)}"},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR,
            )

