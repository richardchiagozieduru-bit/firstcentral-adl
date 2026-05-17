from django import forms
from django.contrib.auth.forms import AuthenticationForm, UserCreationForm
from django.contrib.auth.models import User
from django.core.exceptions import ValidationError
from django.utils import timezone
from auto.models import Subscriber, SubscriberToken, UserSubscriberPermission
from .models import UserProfile


class CustomAuthenticationForm(AuthenticationForm):
    """
    Custom authentication form with enhanced styling and validation
    """
    username = forms.CharField(
        max_length=254,
        widget=forms.TextInput(attrs={
            'class': 'form-control',
            'placeholder': 'Enter your username',
            'autofocus': True
        })
    )
    
    password = forms.CharField(
        label="Password",
        strip=False,
        widget=forms.PasswordInput(attrs={
            'class': 'form-control',
            'placeholder': 'Enter your password'
        })
    )
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Add Bootstrap classes to all fields
        for field_name, field in self.fields.items():
            field.widget.attrs.update({'class': 'form-control'})


class CustomUserCreationForm(UserCreationForm):
    """
    Custom user creation form with styling and email collection
    """
    email = forms.EmailField(
        required=True,
        widget=forms.EmailInput(attrs={
            'class': 'form-control',
            'placeholder': 'Enter your email address'
        }),
        help_text='Reports will be sent to this email after processing.'
    )
    
    class Meta:
        model = User
        fields = ('username', 'email', 'password1', 'password2')
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Add Bootstrap classes to all fields
        for field_name, field in self.fields.items():
            field.widget.attrs.update({
                'class': 'form-control',
            })
        
        # Update placeholders
        self.fields['username'].widget.attrs.update({
            'placeholder': 'Enter your username'
        })
        self.fields['password1'].widget.attrs.update({
            'placeholder': 'Enter your password'
        })
        self.fields['password2'].widget.attrs.update({
            'placeholder': 'Confirm your password'
        })
    
    def clean_email(self):
        email = self.cleaned_data.get('email')
        if User.objects.filter(email=email).exists():
            raise ValidationError('An account with this email already exists.')
        return email


class SubscriberSelectionForm(forms.Form):
    """
    Form for one-time subscriber binding using tokens.
    This form is only used by users who are not yet bound to a subscriber.
    After successful validation, the user will be permanently bound to the selected subscriber.
    """
    subscriber_name = forms.ChoiceField(
        label='Select Your Organization',
        choices=[],
        widget=forms.Select(attrs={
            'class': 'form-control',
            'required': True
        }),
        help_text='Choose the organization you want to bind your account to'
    )
    
    subscriber_token = forms.CharField(
        label='One-Time Binding Token',
        max_length=32,
        widget=forms.PasswordInput(attrs={
            'class': 'form-control',
            'placeholder': 'Enter your one-time binding token',
            'required': True
        }),
        help_text='Enter the token provided by your organization for account binding'
    )
    
    def __init__(self, *args, **kwargs):
        self.user = kwargs.pop('user', None)
        super().__init__(*args, **kwargs)
        
        # Populate all available subscribers for one-time binding
        # This form is only shown to unbound users, so they can choose any organization
        try:
            choices = [('', 'Select your organization...')]
            
            # Show all available subscribers
            subscribers = Subscriber.objects.all().order_by('subscriber_name')
            choices.extend([
                (subscriber.subscriber_name, subscriber.subscriber_name)
                for subscriber in subscribers
            ])
                
            self.fields['subscriber_name'].choices = choices
            
            # Update help text for one-time binding
            self.fields['subscriber_name'].help_text = (
                'Select the organization you want to permanently bind your account to. '
                'This is a one-time choice that cannot be changed without administrator assistance.'
            )
            
            self.fields['subscriber_token'].help_text = (
                'Enter the one-time binding token provided by your organization. '
                'This token will be consumed after successful binding.'
            )
                    
        except Exception as e:
            # Handle database connection issues gracefully
            self.fields['subscriber_name'].choices = [('', 'Error loading organizations')]
    
    def clean(self):
        cleaned_data = super().clean()
        subscriber_name = cleaned_data.get('subscriber_name')
        subscriber_token = cleaned_data.get('subscriber_token')
        
        if not subscriber_name or not subscriber_token:
            return cleaned_data
        
        try:
            # Verify user is not already bound
            if self.user and self.user.is_authenticated:
                user_profile = UserProfile.get_or_create_profile(self.user)
                if user_profile.is_bound:
                    raise ValidationError(
                        f'Your account is already bound to {user_profile.get_bound_subscriber().subscriber_name}. '
                        'Contact an administrator if you need to change your binding.'
                    )
            
            # Get the subscriber by name
            subscriber = Subscriber.objects.get(subscriber_name=subscriber_name)
            
            # Validate the token specifically for binding
            token_obj = SubscriberToken.validate_token_for_binding(
                subscriber_token, 
                subscriber.subscriber_id
            )
            
            if not token_obj:
                raise ValidationError(
                    'Invalid or already used token for the selected organization. '
                    'Please check your token or contact your administrator for a new one.'
                )
            
            # Check if token has expired
            if token_obj.expiry_date and token_obj.expiry_date < timezone.now():
                raise ValidationError(
                    'This token has expired. Please contact your administrator for a new token.'
                )
            
            # Store validated objects for use in the view
            cleaned_data['validated_subscriber'] = subscriber
            cleaned_data['validated_token'] = token_obj
            
        except Subscriber.DoesNotExist:
            raise ValidationError(
                'Selected organization not found. Please refresh the page and try again.'
            )
        except Exception as e:
            raise ValidationError(
                f'An error occurred during validation: {str(e)}'
            )
        
        return cleaned_data
    
    def clean_subscriber_name(self):
        subscriber_name = self.cleaned_data.get('subscriber_name')
        if not subscriber_name:
            raise ValidationError('Please select a subscriber.')
        return subscriber_name
    
    def clean_subscriber_token(self):
        subscriber_token = self.cleaned_data.get('subscriber_token')
        if not subscriber_token:
            raise ValidationError('Please enter a validation token.')
        
        # Basic token format validation
        if len(subscriber_token) < 8:
            raise ValidationError('Token must be at least 8 characters long.')
        
        return subscriber_token.strip()

