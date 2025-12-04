from django.shortcuts import redirect
from django.urls import reverse
from django.contrib import messages
from django.utils.deprecation import MiddlewareMixin


class SubscriberSessionMiddleware(MiddlewareMixin):
    """
    Middleware to enforce subscriber session validation for protected views.
    Ensures users have selected and validated a subscriber before accessing
    data cleaning functionality.
    """
    
    # URLs that don't require subscriber validation
    EXEMPT_URLS = [
        '/acctmgt/login/',
        '/acctmgt/logout/',
        '/acctmgt/register/',
        '/acctmgt/subscriber-selection/',
        '/acctmgt/clear-subscriber/',
        '/admin/',
        '/static/',
        '/media/',
    ]
    
    def process_request(self, request):
        """
        Check if the user needs subscriber binding before accessing protected views
        """
        # Skip middleware for exempt URLs
        if any(request.path.startswith(url) for url in self.EXEMPT_URLS):
            return None
        
        # Skip middleware for unauthenticated users (they'll be handled by LoginRequiredMixin)
        if not request.user.is_authenticated:
            return None
        
        # Check if user is bound to a subscriber (new binding system)
        from .models import UserProfile
        try:
            user_profile = UserProfile.get_or_create_profile(request.user)
            
            if not user_profile.is_bound:
                # User is authenticated but not bound to any subscriber
                messages.warning(
                    request,
                    'Please complete the one-time binding to your organization before accessing the application.'
                )
                return redirect('acctmgt:subscriber_selection')
            
            # Get bound subscriber information
            bound_subscriber = user_profile.get_bound_subscriber()
            if not bound_subscriber:
                # Bound but subscriber doesn't exist - data integrity issue
                messages.error(
                    request,
                    'Unable to retrieve your organization information. Please contact support.'
                )
                return redirect('acctmgt:subscriber_selection')
            
            # Add subscriber info to request for easy access in views
            request.subscriber_id = bound_subscriber.subscriber_id
            request.subscriber_name = bound_subscriber.subscriber_name
            
        except Exception as e:
            # Handle any database or model errors gracefully
            messages.error(
                request,
                'Error retrieving your organization information. Please try again.'
            )
            return redirect('acctmgt:subscriber_selection')
        
        return None
    
    def process_response(self, request, response):
        """
        Add subscriber context to response headers for debugging (optional)
        """
        if hasattr(request, 'subscriber_id') and hasattr(request, 'subscriber_name'):
            response['X-Subscriber-ID'] = str(request.subscriber_id)
            response['X-Subscriber-Name'] = request.subscriber_name
        
        return response


class SubscriberDataFilterMixin:
    """
    Mixin to automatically filter data by subscriber in views.
    Use this mixin in views that need to filter data by the current subscriber.
    """
    
    def get_subscriber_id(self):
        """
        Get the current subscriber ID from the request
        """
        return getattr(self.request, 'subscriber_id', None)
    
    def get_subscriber_name(self):
        """
        Get the current subscriber name from the request
        """
        return getattr(self.request, 'subscriber_name', None)
    
    def filter_queryset_by_subscriber(self, queryset, subscriber_field='subscriber_id'):
        """
        Filter a queryset by the current subscriber
        
        Args:
            queryset: The queryset to filter
            subscriber_field: The field name to filter by (default: 'subscriber_id')
        
        Returns:
            Filtered queryset
        """
        subscriber_id = self.get_subscriber_id()
        if subscriber_id:
            filter_kwargs = {subscriber_field: subscriber_id}
            return queryset.filter(**filter_kwargs)
        return queryset.none()  # Return empty queryset if no subscriber
    
    def get_context_data(self, **kwargs):
        """
        Add subscriber information to template context
        """
        context = super().get_context_data(**kwargs)
        context['current_subscriber_id'] = self.get_subscriber_id()
        context['current_subscriber_name'] = self.get_subscriber_name()
        return context


def get_current_subscriber(request):
    """
    Utility function to get current subscriber information from request
    
    Returns:
        dict: {'id': subscriber_id, 'name': subscriber_name} or None
    """
    subscriber_id = getattr(request, 'subscriber_id', None)
    subscriber_name = getattr(request, 'subscriber_name', None)
    
    if subscriber_id and subscriber_name:
        return {
            'id': subscriber_id,
            'name': subscriber_name
        }
    return None


def require_subscriber_session(view_func):
    """
    Decorator to ensure a view requires subscriber session validation
    
    Usage:
        @require_subscriber_session
        def my_view(request):
            # This view will only be accessible if user has validated subscriber
            pass
    """
    def wrapper(request, *args, **kwargs):
        subscriber_info = get_current_subscriber(request)
        if not subscriber_info:
            return redirect('acctmgt:subscriber_selection')
        return view_func(request, *args, **kwargs)
    return wrapper