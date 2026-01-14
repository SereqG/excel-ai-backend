"""
Middleware for Clerk authentication integration.
"""
from django.utils.deprecation import MiddlewareMixin


class ClerkUserMiddleware(MiddlewareMixin):
    """
    Middleware to extract Clerk user ID from request headers.
    
    Assumes Clerk user ID is passed in the 'X-Clerk-User-Id' header.
    You may need to adjust this based on your actual Clerk integration.
    """
    
    def process_request(self, request):
        """
        Extract Clerk user ID from request headers and attach to request.
        
        The user ID can come from:
        - X-Clerk-User-Id header (custom header)
        - Authorization header (JWT token - would need decoding)
        - Or other method based on your Clerk setup
        
        For now, we'll check X-Clerk-User-Id header.
        If you're using JWT tokens, you'll need to decode the token here.
        """
        # Extract from custom header (adjust based on your frontend setup)
        clerk_user_id = request.META.get('HTTP_X_CLERK_USER_ID')
        
        # If not in header, try to get from request data (for testing)
        if not clerk_user_id:
            # This is a fallback - in production, you should always use headers
            clerk_user_id = getattr(request, 'data', {}).get('clerk_user_id')
        
        # Attach to request object
        request.clerk_user_id = clerk_user_id
        
        return None
