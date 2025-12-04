from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
from auto.models import Subscriber, SubscriberToken

class UserProfile(models.Model):
    """
    Extended user profile to store permanent subscriber binding.
    Implements one-time token binding where users are permanently
    associated with a subscriber after initial token validation.
    """
    user = models.OneToOneField(
        User, 
        on_delete=models.CASCADE, 
        related_name='profile'
    )
    
    # Permanent subscriber binding
    bound_subscriber_id = models.IntegerField(
        null=True, 
        blank=True,
        help_text="Permanently bound subscriber ID after one-time token validation",
        db_index=True
    )
    
    # Binding metadata
    binding_token = models.CharField(
        max_length=32,
        null=True,
        blank=True,
        help_text="Token used for initial binding (for audit purposes)"
    )
    
    bound_at = models.DateTimeField(
        null=True,
        blank=True,
        help_text="Timestamp when user was bound to subscriber"
    )
    
    bound_by_ip = models.GenericIPAddressField(
        null=True,
        blank=True,
        help_text="IP address from which binding was performed"
    )
    
    # Status flags
    is_bound = models.BooleanField(
        default=False,
        db_index=True,
        help_text="Whether user is permanently bound to a subscriber"
    )
    
    binding_method = models.CharField(
        max_length=20,
        choices=[
            ('token', 'Token Validation'),
            ('admin', 'Admin Assignment'),
            ('migration', 'Data Migration')
        ],
        default='token',
        help_text="How the binding was established"
    )
    
    # Timestamps
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    class Meta:
        db_table = 'user_profiles'
        verbose_name = 'User Profile'
        verbose_name_plural = 'User Profiles'
        indexes = [
            models.Index(fields=['bound_subscriber_id', 'is_bound']),
            models.Index(fields=['user', 'is_bound']),
        ]
    
    def __str__(self):
        if self.is_bound:
            subscriber = self.get_bound_subscriber()
            subscriber_name = subscriber.subscriber_name if subscriber else f"ID:{self.bound_subscriber_id}"
            return f"{self.user.username} → {subscriber_name}"
        return f"{self.user.username} (Unbound)"
    
    def get_bound_subscriber(self):
        """
        Get the bound subscriber object.
        Returns None if not bound or subscriber doesn't exist.
        """
        if not self.is_bound or not self.bound_subscriber_id:
            return None
        
        try:
            return Subscriber.objects.get(subscriber_id=self.bound_subscriber_id)
        except Subscriber.DoesNotExist:
            return None
    
    def bind_to_subscriber(self, subscriber_id, token_string, ip_address=None, method='token'):
        """
        Permanently bind user to a subscriber.
        
        Args:
            subscriber_id: ID of the subscriber to bind to
            token_string: Token used for binding
            ip_address: IP address of the binding request
            method: How the binding was established
        
        Returns:
            bool: True if binding was successful
        """
        if self.is_bound:
            return False  # Already bound
        
        # Verify subscriber exists
        try:
            subscriber = Subscriber.objects.get(subscriber_id=subscriber_id)
        except Subscriber.DoesNotExist:
            return False
        
        # Perform binding
        self.bound_subscriber_id = subscriber_id
        self.binding_token = token_string
        self.bound_at = timezone.now()
        self.bound_by_ip = ip_address
        self.is_bound = True
        self.binding_method = method
        self.save()
        
        return True
    
    def unbind_subscriber(self, reason='admin_action'):
        """
        Remove subscriber binding (admin function).
        
        Args:
            reason: Reason for unbinding
        
        Returns:
            bool: True if unbinding was successful
        """
        if not self.is_bound:
            return False
        
        # Create audit log before unbinding
        UnbindingAuditLog.objects.create(
            user=self.user,
            previous_subscriber_id=self.bound_subscriber_id,
            previous_binding_token=self.binding_token,
            previous_bound_at=self.bound_at,
            unbinding_reason=reason,
            unbound_at=timezone.now()
        )
        
        # Clear binding
        self.bound_subscriber_id = None
        self.binding_token = None
        self.bound_at = None
        self.bound_by_ip = None
        self.is_bound = False
        self.save()
        
        return True
    
    @classmethod
    def get_or_create_profile(cls, user):
        """
        Get or create user profile for a given user.
        
        Args:
            user: Django User instance
        
        Returns:
            UserProfile instance
        """
        profile, created = cls.objects.get_or_create(
            user=user,
            defaults={
                'is_bound': False
            }
        )
        return profile
    
    @classmethod
    def get_bound_users_for_subscriber(cls, subscriber_id):
        """
        Get all users bound to a specific subscriber.
        
        Args:
            subscriber_id: ID of the subscriber
        
        Returns:
            QuerySet of UserProfile instances
        """
        return cls.objects.filter(
            bound_subscriber_id=subscriber_id,
            is_bound=True
        ).select_related('user')


class UnbindingAuditLog(models.Model):
    """
    Audit log for tracking when users are unbound from subscribers.
    Important for security and compliance.
    """
    user = models.ForeignKey(
        User,
        on_delete=models.CASCADE,
        related_name='unbinding_logs'
    )
    
    # Previous binding information
    previous_subscriber_id = models.IntegerField(
        help_text="Subscriber ID user was previously bound to"
    )
    
    previous_binding_token = models.CharField(
        max_length=32,
        null=True,
        blank=True,
        help_text="Token that was used for original binding"
    )
    
    previous_bound_at = models.DateTimeField(
        null=True,
        blank=True,
        help_text="When the original binding occurred"
    )
    
    # Unbinding information
    unbinding_reason = models.CharField(
        max_length=100,
        help_text="Reason for unbinding"
    )
    
    unbound_at = models.DateTimeField(
        help_text="When the unbinding occurred"
    )
    
    unbound_by = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='performed_unbindings',
        help_text="Admin user who performed the unbinding"
    )
    
    class Meta:
        db_table = 'unbinding_audit_logs'
        verbose_name = 'Unbinding Audit Log'
        verbose_name_plural = 'Unbinding Audit Logs'
        ordering = ['-unbound_at']
    
    def __str__(self):
        return f"{self.user.username} unbound from subscriber {self.previous_subscriber_id} at {self.unbound_at}"
