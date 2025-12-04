from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from django.contrib.auth.models import User
from django.utils.html import format_html
from django.urls import reverse
from django.contrib import messages
from django.utils import timezone
from .models import UserProfile, UnbindingAuditLog


class UserProfileInline(admin.StackedInline):
    model = UserProfile
    can_delete = False
    verbose_name_plural = 'User Profile'
    fk_name = 'user'
    readonly_fields = ('binding_token', 'bound_at', 'bound_by_ip')
    
    fieldsets = (
        ('Subscriber Binding', {
            'fields': (
                'is_bound', 'bound_subscriber_id', 'binding_method',
                'binding_token', 'bound_at', 'bound_by_ip'
            ),
            'description': 'Permanent binding between user and subscriber'
        }),
    )


class CustomUserAdmin(UserAdmin):
    inlines = (UserProfileInline, )
    list_display = ('username', 'email', 'first_name', 'last_name', 'is_staff', 'get_bound_subscriber')
    list_filter = ('is_staff', 'is_superuser', 'is_active', 'profile__is_bound')
    search_fields = ('username', 'email', 'first_name', 'last_name', 'profile__bound_subscriber_id')
    actions = ['unbind_selected_users', 'show_binding_details']
    
    def get_bound_subscriber(self, obj):
        try:
            if hasattr(obj, 'profile') and obj.profile.is_bound:
                subscriber = obj.profile.get_bound_subscriber()
                if subscriber:
                    return format_html(
                        '<span style="color: green;">{}</span>',
                        subscriber.subscriber_name
                    )
                return format_html(
                    '<span style="color: orange;">ID:{}</span>',
                    obj.profile.bound_subscriber_id
                )
            return format_html('<span style="color: red;">Not Bound</span>')
        except Exception:
            return format_html('<span style="color: red;">Error</span>')
    
    get_bound_subscriber.short_description = 'Bound Subscriber'
    get_bound_subscriber.admin_order_field = 'profile__bound_subscriber_id'
    
    def get_inline_instances(self, request, obj=None):
        if not obj:
            return []
        return super().get_inline_instances(request, obj)
    
    def unbind_selected_users(self, request, queryset):
        """
        Admin action to unbind selected users from their subscribers
        """
        bound_users = []
        unbound_users = []
        
        for user in queryset:
            try:
                profile = UserProfile.get_or_create_profile(user)
                if profile.is_bound:
                    subscriber = profile.get_bound_subscriber()
                    subscriber_name = subscriber.subscriber_name if subscriber else f"ID:{profile.bound_subscriber_id}"
                    
                    # Perform unbinding with admin user tracking
                    success = profile.unbind_subscriber(reason=f'admin_unbind_by_{request.user.username}')
                    
                    if success:
                        # Update the audit log to include who performed the unbinding
                        latest_log = UnbindingAuditLog.objects.filter(user=user).order_by('-unbound_at').first()
                        if latest_log:
                            latest_log.unbound_by = request.user
                            latest_log.save()
                        
                        bound_users.append(f"{user.username} (from {subscriber_name})")
                    else:
                        messages.error(request, f'Failed to unbind user {user.username}')
                else:
                    unbound_users.append(user.username)
            except Exception as e:
                messages.error(request, f'Error unbinding {user.username}: {str(e)}')
        
        if bound_users:
            messages.success(
                request, 
                f'Successfully unbound {len(bound_users)} users: {", ".join(bound_users)}'
            )
        
        if unbound_users:
            messages.warning(
                request,
                f'{len(unbound_users)} users were already unbound: {", ".join(unbound_users)}'
            )
    
    unbind_selected_users.short_description = "Unbind selected users from their subscribers"
    
    def show_binding_details(self, request, queryset):
        """
        Admin action to show detailed binding information for selected users
        """
        details = []
        
        for user in queryset:
            try:
                profile = UserProfile.get_or_create_profile(user)
                if profile.is_bound:
                    subscriber = profile.get_bound_subscriber()
                    subscriber_name = subscriber.subscriber_name if subscriber else f"ID:{profile.bound_subscriber_id}"
                    
                    bound_date = profile.bound_at.strftime('%Y-%m-%d %H:%M:%S') if profile.bound_at else 'Unknown'
                    method = profile.get_binding_method_display()
                    token = profile.binding_token or 'N/A'
                    
                    details.append(
                        f"{user.username}: Bound to {subscriber_name} on {bound_date} via {method} (Token: {token})"
                    )
                else:
                    details.append(f"{user.username}: Not bound to any subscriber")
            except Exception as e:
                details.append(f"{user.username}: Error retrieving binding info - {str(e)}")
        
        if details:
            message = "Binding Details:\n" + "\n".join(details)
            messages.info(request, message)
        else:
            messages.warning(request, "No binding details found for selected users")
    
    show_binding_details.short_description = "Show binding details for selected users"


@admin.register(UnbindingAuditLog)
class UnbindingAuditLogAdmin(admin.ModelAdmin):
    list_display = ('user', 'previous_subscriber_id', 'unbinding_reason', 'unbound_at', 'unbound_by')
    list_filter = ('unbound_at', 'unbinding_reason')
    search_fields = ('user__username', 'previous_subscriber_id', 'unbinding_reason')
    readonly_fields = ('user', 'previous_subscriber_id', 'previous_binding_token', 
                      'previous_bound_at', 'unbinding_reason', 'unbound_at', 'unbound_by')
    
    fieldsets = (
        ('User Information', {
            'fields': ('user',)
        }),
        ('Previous Binding', {
            'fields': ('previous_subscriber_id', 'previous_binding_token', 'previous_bound_at')
        }),
        ('Unbinding Details', {
            'fields': ('unbinding_reason', 'unbound_at', 'unbound_by')
        }),
    )
    
    def has_add_permission(self, request):
        return False
    
    def has_change_permission(self, request, obj=None):
        return False


@admin.register(UserProfile)
class UserProfileAdmin(admin.ModelAdmin):
    """
    Dedicated admin interface for UserProfile to manage bindings
    """
    list_display = ('user', 'is_bound', 'get_subscriber_name', 'binding_method', 'bound_at')
    list_filter = ('is_bound', 'binding_method', 'bound_at')
    search_fields = ('user__username', 'user__email', 'bound_subscriber_id')
    readonly_fields = ('binding_token', 'bound_at', 'bound_by_ip')
    actions = ['unbind_selected_profiles']
    
    fieldsets = (
        ('User Information', {
            'fields': ('user',)
        }),
        ('Binding Status', {
            'fields': ('is_bound', 'bound_subscriber_id', 'binding_method')
        }),
        ('Binding Details', {
            'fields': ('binding_token', 'bound_at', 'bound_by_ip'),
            'classes': ('collapse',)
        }),
    )
    
    def get_subscriber_name(self, obj):
        if obj.is_bound:
            subscriber = obj.get_bound_subscriber()
            if subscriber:
                return format_html(
                    '<span style="color: green;">{}</span>',
                    subscriber.subscriber_name
                )
            return format_html(
                '<span style="color: orange;">ID:{}</span>',
                obj.bound_subscriber_id
            )
        return format_html('<span style="color: red;">Not Bound</span>')
    
    get_subscriber_name.short_description = 'Subscriber'
    get_subscriber_name.admin_order_field = 'bound_subscriber_id'
    
    def unbind_selected_profiles(self, request, queryset):
        """
        Admin action to unbind selected user profiles
        """
        unbound_count = 0
        already_unbound = 0
        
        for profile in queryset:
            if profile.is_bound:
                subscriber = profile.get_bound_subscriber()
                subscriber_name = subscriber.subscriber_name if subscriber else f"ID:{profile.bound_subscriber_id}"
                
                success = profile.unbind_subscriber(reason=f'admin_unbind_by_{request.user.username}')
                
                if success:
                    # Update the audit log to include who performed the unbinding
                    latest_log = UnbindingAuditLog.objects.filter(user=profile.user).order_by('-unbound_at').first()
                    if latest_log:
                        latest_log.unbound_by = request.user
                        latest_log.save()
                    
                    unbound_count += 1
            else:
                already_unbound += 1
        
        if unbound_count > 0:
            messages.success(request, f'Successfully unbound {unbound_count} users.')
        
        if already_unbound > 0:
            messages.warning(request, f'{already_unbound} users were already unbound.')
    
    unbind_selected_profiles.short_description = "Unbind selected user profiles"


# Unregister the default UserAdmin and register our custom one
admin.site.unregister(User)
admin.site.register(User, CustomUserAdmin)
