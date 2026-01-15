"""
Middleware for Clerk authentication integration.
"""
from typing import Callable


def clerk_user_middleware(get_response: Callable) -> Callable:
    """
    Middleware to extract Clerk user ID from request headers.
    
    Assumes Clerk user ID is passed in the 'X-Clerk-User-Id' header.
    You may need to adjust this based on your actual Clerk integration.
    
    Args:
        get_response: The next middleware or view in the chain
        
    Returns:
        Middleware function that processes requests
    """
    def middleware(request):
        """
        Extract Clerk user ID from request headers and attach to request.
        
        The user ID can come from:
        - X-Clerk-User-Id header (custom header)
        - Authorization header (JWT token - would need decoding)
        - Or other method based on your Clerk setup
        
        For now, we'll check X-Clerk-User-Id header.
        If you're using JWT tokens, you'll need to decode the token here.
        """
        clerk_user_id = request.META.get('HTTP_X_CLERK_USER_ID')
        
        if not clerk_user_id:
            clerk_user_id = getattr(request, 'data', {}).get('clerk_user_id')
        
        request.clerk_user_id = clerk_user_id
        
        response = get_response(request)
        
        return response
    
    return middleware
