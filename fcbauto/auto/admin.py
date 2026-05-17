from django.contrib import admin
from django.contrib.auth.models import User
from django.utils.html import format_html
from django.urls import reverse
from django.utils.safestring import mark_safe
from django.contrib import messages
from django.http import HttpResponse
from django.template.response import TemplateResponse
from django.shortcuts import render
from django import forms
from .models import Subscriber, SubscriberToken, UserSubscriberPermission, Feedback


class SubscriberMultipleChoiceField(forms.ModelMultipleChoiceField):
    """
    Custom field that handles float subscriber IDs from MSSQL database
    """
    def validate(self, value):
        """Override validation to handle float primary keys"""
        if self.required and not value:
            raise forms.ValidationError(self.error_messages['required'], code='required')
        
        # Convert float IDs to match the actual database values
        if value:
            # Get all subscriber objects and create a mapping
            all_subscribers = list(self.queryset.all())
            pk_map = {}
            
            # Create mapping for both int and float representations
            for subscriber in all_subscribers:
                pk_value = subscriber.pk
                if pk_value is not None:  # Check for None values
                    pk_map[str(pk_value)] = subscriber
                    try:
                        pk_map[str(int(float(pk_value)))] = subscriber  # Handle "44.0" -> "44"
                    except (ValueError, TypeError):
                        pass  # Skip invalid values
            
            # Validate each selected value
            for pk in value:
                if str(pk) not in pk_map:
                    raise forms.ValidationError(
                        self.error_messages['invalid_choice'],
                        code='invalid_choice',
                        params={'value': pk},
                    )
    
    def _check_values(self, value):
        """Override to handle float/int conversion"""
        # Get all valid PKs in both formats
        all_subscribers = list(self.queryset.all())
        valid_pks = set()
        
        for subscriber in all_subscribers:
             pk_value = subscriber.pk
             if pk_value is not None:  # Check for None values
                 valid_pks.add(str(pk_value))
                 try:
                     valid_pks.add(str(int(float(pk_value))))  # Add integer version
                 except (ValueError, TypeError):
                     pass  # Skip invalid values
        
        # Filter to only valid objects
        result = []
        for pk in value:
            if str(pk) in valid_pks:
                 # Find the actual subscriber object
                 for subscriber in all_subscribers:
                     if subscriber.pk is not None:
                         try:
                             if str(subscriber.pk) == str(pk) or str(int(float(subscriber.pk))) == str(pk):
                                 result.append(subscriber)
                                 break
                         except (ValueError, TypeError):
                             continue  # Skip invalid values
        
        return result


class BulkTokenGenerationForm(forms.Form):
    """
    Form for bulk token generation
    """
    subscribers = SubscriberMultipleChoiceField(
        queryset=Subscriber.objects.none(),  # Will be set in __init__
        widget=forms.CheckboxSelectMultiple,
        required=True,
        help_text="Select subscribers to generate tokens for"
    )
    tokens_per_subscriber = forms.IntegerField(
        min_value=1,
        max_value=50,
        initial=1,
        help_text="Number of tokens to generate per subscriber (max 50)"
    )
    expiry_days = forms.IntegerField(
        min_value=0,
        max_value=365,
        initial=30,
        required=False,
        help_text="Token expiry in days (0 = no expiry)"
    )
    is_active = forms.BooleanField(
        initial=True,
        required=False,
        help_text="Generate tokens as active"
    )
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        try:
            # Set the queryset for subscribers with error handling
            self.fields['subscribers'].queryset = Subscriber.objects.all()
        except Exception as e:
            # If database connection fails, provide empty queryset
            self.fields['subscribers'].queryset = Subscriber.objects.none()
            self.fields['subscribers'].help_text = f"Error loading subscribers: {str(e)}"


@admin.register(Subscriber)
class SubscriberAdmin(admin.ModelAdmin):
    """
    Admin interface for Subscriber model (read-only since it's an external table)
    """
    list_display = ('subscriber_id', 'subscriber_name')
    list_filter = ('subscriber_name',)
    search_fields = ('subscriber_name', 'subscriber_id')
    readonly_fields = ('subscriber_id', 'subscriber_name')
    ordering = ('subscriber_name',)
    
    def has_add_permission(self, request):
        return False  # External table, no adding
    
    def has_delete_permission(self, request, obj=None):
        return False  # External table, no deleting
    
    def has_change_permission(self, request, obj=None):
        return False  # External table, no editing




@admin.register(SubscriberToken)
class SubscriberTokenAdmin(admin.ModelAdmin):
    """
    Admin interface for SubscriberToken model with bulk generation capabilities
    """
    list_display = (
        'token_display', 'subscriber_info', 'is_active', 
        'usage_count', 'created_at', 'expiry_date', 'created_by'
    )
    list_filter = (
        'is_active', 'created_at', 'expiry_date', 'created_by'
    )
    search_fields = ('token', 'subscriber_id')
    readonly_fields = ('token', 'created_at', 'last_used_at', 'usage_count')
    ordering = ('-created_at',)
    date_hierarchy = 'created_at'
    actions = ['bulk_generate_tokens', 'deactivate_selected_tokens', 'activate_selected_tokens']
    
    fieldsets = (
        ('Token Information', {
            'fields': ('token', 'subscriber_id', 'is_active')
        }),
        ('Usage & Expiry', {
            'fields': ('usage_count', 'last_used_at', 'expiry_date')
        }),
        ('Metadata', {
            'fields': ('created_at', 'created_by'),
            'classes': ('collapse',)
        })
    )
    
    def get_urls(self):
        """Add custom URLs for bulk operations"""
        from django.urls import path
        urls = super().get_urls()
        custom_urls = [
            path('bulk-generate/', self.admin_site.admin_view(self.bulk_generate_view), name='auto_subscribertoken_bulk_generate'),
        ]
        return custom_urls + urls
    
    def bulk_generate_view(self, request):
        """Custom view for bulk token generation"""
        if request.method == 'POST':
            form = BulkTokenGenerationForm(request.POST)
            if form.is_valid():
                return self._process_bulk_generation(request, form)
        else:
            form = BulkTokenGenerationForm()
        
        context = {
            'form': form,
            'title': 'Bulk Generate Subscriber Tokens',
            'opts': self.model._meta,
            'has_change_permission': self.has_change_permission(request),
        }
        return render(request, 'admin/auto/bulk_token_generation.html', context)
    
    def _process_bulk_generation(self, request, form):
        """Process the bulk token generation"""
        from django.utils import timezone
        from datetime import timedelta
        
        subscribers = form.cleaned_data['subscribers']
        tokens_per_subscriber = form.cleaned_data['tokens_per_subscriber']
        expiry_days = form.cleaned_data['expiry_days']
        is_active = form.cleaned_data['is_active']
        
        # Calculate expiry date
        expiry_date = None
        if expiry_days > 0:
            expiry_date = timezone.now() + timedelta(days=expiry_days)
        
        total_tokens_created = 0
        created_tokens = []
        
        try:
            for subscriber in subscribers:
                for i in range(tokens_per_subscriber):
                    token = SubscriberToken.objects.create(
                        subscriber_id=int(subscriber.subscriber_id),  # Convert to int
                        is_active=is_active,
                        expiry_date=expiry_date,
                        created_by=request.user
                    )
                    created_tokens.append(token)
                    total_tokens_created += 1
            
            # Success message
            messages.success(
                request,
                f'Successfully generated {total_tokens_created} tokens for {len(subscribers)} subscribers.'
            )
            
            # Redirect to changelist with filter to show newly created tokens
            from django.shortcuts import redirect
            return redirect('admin:auto_subscribertoken_changelist')
            
        except Exception as e:
            messages.error(request, f'Error generating tokens: {str(e)}')
            form = BulkTokenGenerationForm(request.POST)
            context = {
                'form': form,
                'title': 'Bulk Generate Subscriber Tokens',
                'opts': self.model._meta,
                'has_change_permission': self.has_change_permission(request),
            }
            return render(request, 'admin/auto/bulk_token_generation.html', context)
    
    def bulk_generate_tokens(self, request, queryset):
        """Admin action to redirect to bulk generation page"""
        from django.shortcuts import redirect
        return redirect('admin:auto_subscribertoken_bulk_generate')
    bulk_generate_tokens.short_description = "Generate tokens in bulk"
    
    def deactivate_selected_tokens(self, request, queryset):
        """Admin action to deactivate selected tokens"""
        updated = queryset.update(is_active=False)
        messages.success(request, f'Successfully deactivated {updated} tokens.')
    deactivate_selected_tokens.short_description = "Deactivate selected tokens"
    
    def activate_selected_tokens(self, request, queryset):
        """Admin action to activate selected tokens"""
        updated = queryset.update(is_active=True)
        messages.success(request, f'Successfully activated {updated} tokens.')
    activate_selected_tokens.short_description = "Activate selected tokens"
    
    def token_display(self, obj):
        """Display truncated token for security"""
        return f"{obj.token[:8]}...{obj.token[-4:]}"
    token_display.short_description = 'Token'
    
    def subscriber_info(self, obj):
        """Display subscriber information"""
        subscriber = obj.get_subscriber()
        if subscriber:
            return f"{subscriber.subscriber_name} (ID: {obj.subscriber_id})"
        return f"Subscriber ID: {obj.subscriber_id} (Not Found)"
    subscriber_info.short_description = 'Subscriber'
    
    def save_model(self, request, obj, form, change):
        """Set created_by to current user if not set"""
        if not change and not obj.created_by:
            obj.created_by = request.user
        super().save_model(request, obj, form, change)
    
    def changelist_view(self, request, extra_context=None):
        """Add bulk generation button to changelist"""
        extra_context = extra_context or {}
        extra_context['bulk_generate_url'] = 'bulk-generate/'
        return super().changelist_view(request, extra_context)


@admin.register(UserSubscriberPermission)
class UserSubscriberPermissionAdmin(admin.ModelAdmin):
    """
    Admin interface for UserSubscriberPermission model
    """
    list_display = (
        'user', 'subscriber_info', 'permission_level', 
        'is_active', 'granted_at', 'granted_by'
    )
    list_filter = (
        'permission_level', 'is_active', 'granted_at', 'granted_by'
    )
    search_fields = ('user__username', 'user__email', 'subscriber_id')
    ordering = ('-granted_at',)
    date_hierarchy = 'granted_at'
    
    fieldsets = (
        ('Permission Assignment', {
            'fields': ('user', 'subscriber_id', 'permission_level', 'is_active')
        }),
        ('Metadata', {
            'fields': ('granted_by', 'granted_at'),
            'classes': ('collapse',)
        })
    )
    
    readonly_fields = ('granted_at',)
    
    def subscriber_info(self, obj):
        """Display subscriber information"""
        subscriber = obj.get_subscriber()
        if subscriber:
            return f"{subscriber.subscriber_name} (ID: {obj.subscriber_id})"
        return f"Subscriber ID: {obj.subscriber_id} (Not Found)"
    subscriber_info.short_description = 'Subscriber'
    
    def save_model(self, request, obj, form, change):
        """Set granted_by to current user if not set"""
        if not change and not obj.granted_by:
            obj.granted_by = request.user
        super().save_model(request, obj, form, change)


@admin.register(Feedback)
class FeedbackAdmin(admin.ModelAdmin):
    """
    Admin interface for viewing and managing user feedback
    """
    list_display = (
        'created_at', 'rating_display', 'category', 'user', 
        'message_preview', 'is_reviewed'
    )
    list_filter = ('category', 'rating', 'is_reviewed', 'created_at')
    search_fields = ('message', 'user__username', 'user__email')
    ordering = ('-created_at',)
    date_hierarchy = 'created_at'
    readonly_fields = ('user', 'rating', 'category', 'message', 'page_url', 'created_at')
    actions = ['mark_as_reviewed']
    
    fieldsets = (
        ('Feedback Details', {
            'fields': ('user', 'rating', 'category', 'message', 'page_url', 'created_at')
        }),
        ('Admin Review', {
            'fields': ('is_reviewed', 'reviewed_at', 'admin_notes'),
        })
    )
    
    def rating_display(self, obj):
        """Display rating as stars"""
        return '★' * obj.rating + '☆' * (5 - obj.rating)
    rating_display.short_description = 'Rating'
    
    def message_preview(self, obj):
        """Display truncated message"""
        return obj.message[:50] + '...' if len(obj.message) > 50 else obj.message
    message_preview.short_description = 'Message'
    
    def mark_as_reviewed(self, request, queryset):
        """Mark selected feedback as reviewed"""
        from django.utils import timezone
        updated = queryset.update(is_reviewed=True, reviewed_at=timezone.now())
        messages.success(request, f'Marked {updated} feedback entries as reviewed.')
    mark_as_reviewed.short_description = "Mark selected as reviewed"
    
    def has_add_permission(self, request):
        return False  # Feedback is submitted by users, not created in admin
    
    def has_delete_permission(self, request, obj=None):
        return True  # Allow deletion for cleanup
