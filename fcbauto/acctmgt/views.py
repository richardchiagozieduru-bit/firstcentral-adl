from django.shortcuts import render, redirect
from django.contrib.auth import login, authenticate
from django.contrib.auth.views import LoginView
from django.contrib import messages
from django.urls import reverse_lazy
from django.views.generic import CreateView, FormView
from django.contrib.auth.mixins import LoginRequiredMixin
from django.utils import timezone
from django.utils.decorators import method_decorator
from .forms import CustomUserCreationForm, CustomAuthenticationForm, SubscriberSelectionForm

from .models import UserProfile


class CustomLoginView(LoginView):
    """Custom login view with enhanced styling and error handling."""
    form_class = CustomAuthenticationForm
    template_name = 'acctmgt/auth_unified.html'
    redirect_authenticated_user = True
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        # Pre-fill username if provided in GET parameter (e.g., after registration)
        context['username_prefill'] = self.request.GET.get('username', '')
        return context
    
    def get_success_url(self):
        # Check if user is already bound to a subscriber
        user_profile = UserProfile.get_or_create_profile(self.request.user)
        
        if user_profile.is_bound:
            # User is already bound, redirect directly to dashboard
            messages.success(
                self.request,
                f'Welcome back! You are working with {user_profile.get_bound_subscriber().subscriber_name}.'
            )
            return reverse_lazy('auto:dashboard')
        else:
            # User is not bound, redirect to subscriber selection for one-time binding
            return reverse_lazy('acctmgt:subscriber_selection')
    
    def form_invalid(self, form):
        messages.error(self.request, 'Invalid username or password. Please try again.')
        # Don't redirect - instead render the template with the login section shown
        context = self.get_context_data(form=form)
        context['show_section'] = 'login'  # Keep user on login section on failed auth
        return self.render_to_response(context)


class RegisterView(CreateView):
    """User registration view with automatic redirect to login page."""
    form_class = CustomUserCreationForm
    template_name = 'acctmgt/auth_unified.html'
    
    def get_success_url(self):
        # Not used - we override form_valid to handle redirect manually
        username = self.object.username
        return f"{reverse_lazy('acctmgt:login')}?username={username}"
    
    def form_valid(self, form):
        response = super().form_valid(form)
        messages.success(
            self.request, 
            'Account created successfully! Please sign in with your credentials.'
        )
        # Return the login section without redirecting
        context = self.get_context_data(form=form)
        context['show_section'] = 'login'
        context['username_prefill'] = self.object.username
        return self.render_to_response(context)
    
    def form_invalid(self, form):
        messages.error(
            self.request, 
            'Please correct the errors below and try again.'
        )
        # Stay on register section on form errors
        context = self.get_context_data(form=form)
        context['show_section'] = 'register'
        return self.render_to_response(context)
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        # Tell the template to show the register section when there are form errors
        context['show_section'] = 'register'
        return context
    
    def dispatch(self, request, *args, **kwargs):
        # Redirect authenticated users to the main app
        if request.user.is_authenticated:
            return redirect('acctmgt:subscriber_selection')
        return super().dispatch(request, *args, **kwargs)



class SubscriberSelectionView(LoginRequiredMixin, FormView):
    """
    View for subscriber selection and one-time token binding.
    Users who are not yet bound to a subscriber will use this view
    to permanently bind their account to a subscriber using a token.
    """
    form_class = SubscriberSelectionForm
    template_name = 'acctmgt/subscriber_selection.html'
    success_url = reverse_lazy('auto:dashboard')
    
    def dispatch(self, request, *args, **kwargs):
        # Check if user is already bound to a subscriber
        user_profile = UserProfile.get_or_create_profile(request.user)
        
        if user_profile.is_bound:
            # User is already bound, redirect to dashboard
            messages.info(
                request,
                f'You are already bound to {user_profile.get_bound_subscriber().subscriber_name}. '
                f'Contact an administrator if you need to change your subscriber binding.'
            )
            return redirect('auto:dashboard')
        
        return super().dispatch(request, *args, **kwargs)
    
    # Note: Session-based validation methods removed as we now use permanent binding
    
    def get_form_kwargs(self):
        """
        Pass the current user to the form
        """
        kwargs = super().get_form_kwargs()
        kwargs['user'] = self.request.user
        return kwargs
    
    def form_valid(self, form):
        # Get the validated subscriber and token from the form
        validated_subscriber = form.cleaned_data['validated_subscriber']
        validated_token = form.cleaned_data['validated_token']
        
        # Get user profile
        user_profile = UserProfile.get_or_create_profile(self.request.user)
        
        # Get client IP for audit purposes
        client_ip = self.get_client_ip()
        
        # Perform one-time binding
        binding_success = user_profile.bind_to_subscriber(
            subscriber_id=validated_subscriber.subscriber_id,
            token_string=validated_token.token,
            ip_address=client_ip,
            method='token'
        )
        
        if not binding_success:
            messages.error(
                self.request,
                'Failed to bind your account to the subscriber. Please try again or contact support.'
            )
            return self.form_invalid(form)
        
        # Mark token as used for binding
        validated_token.mark_used_for_binding(self.request.user, client_ip)
        
        # Success message
        messages.success(
            self.request,
            f'Welcome! Your account has been permanently bound to {validated_subscriber.subscriber_name}. '
            f'You can now access the data cleaning application directly after login.'
        )
        
        return super().form_valid(form)
    
    def get_client_ip(self):
        """
        Get client IP address for audit purposes
        """
        x_forwarded_for = self.request.META.get('HTTP_X_FORWARDED_FOR')
        if x_forwarded_for:
            ip = x_forwarded_for.split(',')[0]
        else:
            ip = self.request.META.get('REMOTE_ADDR')
        return ip
    
    def form_invalid(self, form):
        messages.error(
            self.request,
            'Authentication failed. Please check your subscriber selection and token.'
        )
        return super().form_invalid(form)
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['user'] = self.request.user
        return context


def clear_subscriber_session(request):
    """
    View to clear subscriber session data (logout from subscriber context)
    """
    session_keys = ['subscriber_id', 'subscriber_name', 'subscriber_token', 'token_validated_at']
    for key in session_keys:
        request.session.pop(key, None)
    
    messages.info(request, 'Subscriber session cleared. Please authenticate again.')
    return redirect('acctmgt:subscriber_selection')
