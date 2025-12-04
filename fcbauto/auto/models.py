from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
import secrets


class Subscriber(models.Model):
    """
    Model to represent subscribers in the existing MSSQL Subscribers database.
    This connects to the Sheet1 table with SubscriberID and SubscriberName columns.
    """
    subscriber_id = models.AutoField(primary_key=True, db_column='SubscriberID')
    subscriber_name = models.CharField(max_length=255, db_column='SubscriberName')
    
    class Meta:
        db_table = 'Sheet1'
        managed = False  # Django won't manage this existing table
    
    def __str__(self):
        return self.subscriber_name or f"Subscriber {self.subscriber_id}"
    
    @classmethod
    def get_subscriber_choices(cls):
        """
        Return a list of tuples for form choices (id, name)
        """
        return [(sub.subscriber_id, sub.subscriber_name) for sub in cls.objects.all()]
    
    @classmethod
    def validate_subscriber(cls, subscriber_name, subscriber_id):
        """
        Validate if the subscriber name and ID match
        """
        try:
            subscriber = cls.objects.get(
                subscriber_name=subscriber_name,
                subscriber_id=subscriber_id
            )
            return subscriber
        except cls.DoesNotExist:
            return None


class SubscriberToken(models.Model):
    """
    Model to manage secure tokens for subscriber access.
    Replaces raw subscriber_id exposure with secure, random tokens.
    Uses integer field for subscriber_id - NO foreign key to avoid constraint issues.
    
    Supports one-time binding where tokens can be marked as used after
    permanently binding a user to a subscriber.
    """
    subscriber_id = models.IntegerField(
        help_text="Reference to subscriber ID from external table",
        db_index=True  # Add index for performance
    )
    token = models.CharField(max_length=32, unique=True, help_text="URL-safe random string")
    is_active = models.BooleanField(default=True, db_index=True)
    created_at = models.DateTimeField(auto_now_add=True)
    expiry_date = models.DateTimeField(null=True, blank=True, help_text="Optional expiration date")
    last_used_at = models.DateTimeField(null=True, blank=True)
    usage_count = models.PositiveIntegerField(default=0)
    created_by = models.ForeignKey(
        User, 
        on_delete=models.SET_NULL, 
        null=True, 
        blank=True, 
        related_name='created_tokens'
    )
    
    # One-time binding fields
    is_used_for_binding = models.BooleanField(
        default=False,
        db_index=True,
        help_text="Whether this token has been used for permanent user-subscriber binding"
    )
    bound_user = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='binding_tokens',
        help_text="User who used this token for permanent binding"
    )
    binding_date = models.DateTimeField(
        null=True,
        blank=True,
        help_text="When this token was used for binding"
    )
    binding_ip = models.GenericIPAddressField(
        null=True,
        blank=True,
        help_text="IP address from which binding was performed"
    )
    
    class Meta:
        db_table = 'subscriber_tokens'
        verbose_name = 'Subscriber Token'
        verbose_name_plural = 'Subscriber Tokens'
        ordering = ['-created_at']
        indexes = [
            models.Index(fields=['subscriber_id', 'is_active']),
            models.Index(fields=['token', 'is_active']),
        ]
    
    def save(self, *args, **kwargs):
        """Generate secure token if not provided"""
        if not self.token:
            self.token = secrets.token_urlsafe(16)
        super().save(*args, **kwargs)
    
    def __str__(self):
        try:
            subscriber = self.get_subscriber()
            subscriber_name = subscriber.subscriber_name if subscriber else f"ID:{self.subscriber_id}"
            return f"Token for {subscriber_name} ({'Active' if self.is_active else 'Inactive'})"
        except:
            return f"Token for Subscriber ID:{self.subscriber_id} ({'Active' if self.is_active else 'Inactive'})"
    
    def get_subscriber(self):
        """
        Get the related subscriber object manually (no FK relationship)
        """
        try:
            return Subscriber.objects.get(subscriber_id=self.subscriber_id)
        except Subscriber.DoesNotExist:
            return None
    
    def is_valid(self):
        """
        Check if token is valid (active and not expired)
        """
        if not self.is_active:
            return False
        if self.expiry_date and timezone.now() > self.expiry_date:
            return False
        return True
    
    def mark_used(self):
        """
        Mark token as used (increment usage count and update last used time)
        """
        self.usage_count += 1
        self.last_used_at = timezone.now()
        self.save(update_fields=['usage_count', 'last_used_at'])
    
    @classmethod
    def validate_token(cls, token_string, subscriber_id=None):
        """
        Validate a token string and optionally check subscriber association
        Returns the token object if valid, None otherwise
        """
        try:
            token_obj = cls.objects.get(token=token_string, is_active=True)
            if not token_obj.is_valid():
                return None
            if subscriber_id and token_obj.subscriber_id != subscriber_id:
                return None
            return token_obj
        except cls.DoesNotExist:
            return None
    
    @classmethod
    def get_tokens_for_subscriber(cls, subscriber_id):
        """
        Get all active tokens for a specific subscriber
        """
        return cls.objects.filter(subscriber_id=subscriber_id, is_active=True)
    
    def can_be_used_for_binding(self):
        """
        Check if this token can be used for one-time binding.
        Returns True if token is valid and hasn't been used for binding yet.
        """
        return self.is_valid() and not self.is_used_for_binding
    
    def mark_used_for_binding(self, user, ip_address=None):
        """
        Mark this token as used for permanent user-subscriber binding.
        
        Args:
            user: User who used this token for binding
            ip_address: IP address from which binding was performed
        
        Returns:
            bool: True if successfully marked as used
        """
        if self.is_used_for_binding:
            return False  # Already used for binding
        
        if not self.is_valid():
            return False  # Token is not valid
        
        # Mark as used for binding
        self.is_used_for_binding = True
        self.bound_user = user
        self.binding_date = timezone.now()
        self.binding_ip = ip_address
        
        # Also update regular usage tracking
        self.usage_count += 1
        self.last_used_at = timezone.now()
        
        self.save(update_fields=[
            'is_used_for_binding', 'bound_user', 'binding_date', 'binding_ip',
            'usage_count', 'last_used_at'
        ])
        
        return True
    
    @classmethod
    def validate_token_for_binding(cls, token_string, subscriber_id=None):
        """
        Validate a token specifically for one-time binding.
        Returns the token object if valid for binding, None otherwise.
        
        Args:
            token_string: The token to validate
            subscriber_id: Optional subscriber ID to check against
        
        Returns:
            SubscriberToken instance if valid for binding, None otherwise
        """
        try:
            token_obj = cls.objects.get(token=token_string, is_active=True)
            
            # Check if token can be used for binding
            if not token_obj.can_be_used_for_binding():
                return None
            
            # Check subscriber association if provided
            if subscriber_id and token_obj.subscriber_id != subscriber_id:
                return None
            
            return token_obj
        except cls.DoesNotExist:
            return None
    
    @classmethod
    def get_available_tokens_for_subscriber(cls, subscriber_id):
        """
        Get all tokens available for binding for a specific subscriber.
        
        Args:
            subscriber_id: ID of the subscriber
        
        Returns:
            QuerySet of available tokens
        """
        return cls.objects.filter(
            subscriber_id=subscriber_id,
            is_active=True,
            is_used_for_binding=False
        ).filter(
            models.Q(expiry_date__isnull=True) | 
            models.Q(expiry_date__gt=timezone.now())
        )
    
    @classmethod
    def get_binding_history_for_user(cls, user):
        """
        Get all tokens that were used by a specific user for binding.
        
        Args:
            user: User instance
        
        Returns:
            QuerySet of tokens used for binding by this user
        """
        return cls.objects.filter(
            bound_user=user,
            is_used_for_binding=True
        ).order_by('-binding_date')


class UserSubscriberPermission(models.Model):
    """
    Model to manage user-subscriber relationships with permission levels.
    Supports both one-to-one and one-to-many user-subscriber scenarios.
    Uses integer field for subscriber_id - NO foreign key to avoid constraint issues.
    """
    PERMISSION_CHOICES = [
        ('read', 'Read Only'),
        ('write', 'Read/Write'),
        ('admin', 'Administrator'),
    ]
    
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='subscriber_permissions')
    subscriber_id = models.IntegerField(
        help_text="Reference to subscriber ID from external table",
        db_index=True  # Add index for performance
    )
    permission_level = models.CharField(max_length=20, choices=PERMISSION_CHOICES, default='read')
    granted_by = models.ForeignKey(
        User, 
        on_delete=models.SET_NULL, 
        null=True, 
        blank=True, 
        related_name='granted_permissions'
    )
    granted_at = models.DateTimeField(auto_now_add=True)
    is_active = models.BooleanField(default=True, db_index=True)
    
    class Meta:
        db_table = 'user_subscriber_permissions'
        verbose_name = 'User Subscriber Permission'
        verbose_name_plural = 'User Subscriber Permissions'
        unique_together = ['user', 'subscriber_id']
        ordering = ['-granted_at']
        indexes = [
            models.Index(fields=['user', 'subscriber_id', 'is_active']),
            models.Index(fields=['subscriber_id', 'is_active']),
        ]
    
    def __str__(self):
        subscriber = self.get_subscriber()
        subscriber_name = subscriber.subscriber_name if subscriber else f"ID:{self.subscriber_id}"
        return f"{self.user.username} - {subscriber_name} ({self.permission_level})"
    
    def get_subscriber(self):
        """
        Get the related subscriber object manually (no FK relationship)
        """
        try:
            return Subscriber.objects.get(subscriber_id=self.subscriber_id)
        except Subscriber.DoesNotExist:
            return None
    
    @classmethod
    def get_user_subscribers(cls, user):
        """
        Get all subscribers a user has access to with their permission details
        Returns a queryset of permission objects
        """
        return cls.objects.filter(user=user, is_active=True).select_related('user')
    
    @classmethod
    def get_subscriber_choices_for_user(cls, user):
        """
        Get subscriber choices formatted for form dropdowns
        """
        user_subscribers = cls.get_user_subscribers(user)
        return [
            (sub['subscriber'].subscriber_id, sub['subscriber'].subscriber_name) 
            for sub in user_subscribers
        ]
    
    @classmethod
    def has_permission(cls, user, subscriber_id, required_permission='read'):
        """
        Check if user has required permission level for a subscriber
        Permission hierarchy: admin > write > read
        """
        permission_hierarchy = {'read': 1, 'write': 2, 'admin': 3}
        required_level = permission_hierarchy.get(required_permission, 1)
        
        try:
            user_permission = cls.objects.get(
                user=user, 
                subscriber_id=subscriber_id, 
                is_active=True
            )
            user_level = permission_hierarchy.get(user_permission.permission_level, 0)
            return user_level >= required_level
        except cls.DoesNotExist:
            return False
    
    @classmethod
    def get_users_for_subscriber(cls, subscriber_id):
        """
        Get all users who have access to a specific subscriber
        """
        return cls.objects.filter(
            subscriber_id=subscriber_id, 
            is_active=True
        ).select_related('user')


class SubscriberUtils:
    """
    Utility class for subscriber-related operations
    """
    
    @staticmethod
    def get_subscriber_by_id(subscriber_id):
        """
        Get subscriber by ID with error handling
        """
        try:
            return Subscriber.objects.get(subscriber_id=subscriber_id)
        except Subscriber.DoesNotExist:
            return None
    
    @staticmethod
    def validate_user_subscriber_access(user, subscriber_id, token_string):
        """
        Comprehensive validation of user access to subscriber with token
        Returns tuple: (is_valid, subscriber_obj, token_obj, permission_obj)
        """
        # Validate token
        token_obj = SubscriberToken.validate_token(token_string, subscriber_id)
        if not token_obj:
            return False, None, None, None
        
        # Get subscriber
        subscriber = SubscriberUtils.get_subscriber_by_id(subscriber_id)
        if not subscriber:
            return False, None, token_obj, None
        
        # Check user permissions
        permission_obj = None
        if user and user.is_authenticated:
            try:
                permission_obj = UserSubscriberPermission.objects.get(
                    user=user, subscriber_id=subscriber_id, is_active=True
                )
            except UserSubscriberPermission.DoesNotExist:
                pass
        
        return True, subscriber, token_obj, permission_obj
    
    @staticmethod
    def get_user_context_info(user):
        """
        Get comprehensive context information for a user
        """
        if not user or not user.is_authenticated:
            return {}
        
        user_permissions = UserSubscriberPermission.get_user_subscribers(user)
        subscribers_info = []
        
        for perm in user_permissions:
            subscriber = perm.get_subscriber()
            if subscriber:
                subscribers_info.append({
                    'subscriber_id': subscriber.subscriber_id,
                    'subscriber_name': subscriber.subscriber_name,
                    'permission_level': perm.permission_level,
                    'granted_at': perm.granted_at
                })
        
        context = {
            'user': user,
            'subscribers': subscribers_info,
            'subscriber_count': len(subscribers_info)
        }
        
        return context


class UploadSession(models.Model):
    """
    Model to track file upload sessions for dashboard analytics and real-time progress tracking
    """
    STATUS_CHOICES = [
        ('pending', 'Pending'),
        ('uploading', 'Uploading'),
        ('processing', 'Processing'),
        ('awaiting_verification', 'Awaiting Verification'),
        ('finalizing', 'Finalizing'),
        ('completed', 'Completed'),
        ('failed', 'Failed'),
    ]
    
    STAGE_CHOICES = [
        ('pending', 'Pending'),
        ('upload_validation', 'File Upload & Validation'),
        ('data_mapping', 'Data Mapping & Preparation'),
        ('data_cleaning', 'Data Cleaning & Transformation'),
        ('credit_matching', 'Credit Matching & Classification'),
        ('awaiting_verification', 'Awaiting Human Verification'),
        ('post_verification', 'Finalization After Verification'),
        ('output_generation', 'Output Generation'),
        ('completed', 'Completed'),
    ]
    
    user = models.ForeignKey(
        User, 
        on_delete=models.CASCADE, 
        related_name='upload_sessions'
    )
    subscriber_id = models.IntegerField(
        help_text="Reference to subscriber ID from external table",
        db_index=True
    )
    filename = models.CharField(max_length=255)
    original_filename = models.CharField(max_length=255)
    file_size = models.PositiveIntegerField(help_text="File size in bytes")
    status = models.CharField(max_length=30, choices=STATUS_CHOICES, default='pending')
    
    # Progress tracking fields
    processing_stage = models.CharField(
        max_length=50, 
        choices=STAGE_CHOICES, 
        default='pending',
        help_text="Current processing stage for detailed progress tracking"
    )
    progress_percentage = models.IntegerField(
        default=0, 
        help_text="Overall progress percentage (0-100)"
    )
    current_message = models.TextField(
        blank=True, 
        help_text="Current processing message for real-time updates"
    )
    activity_log = models.JSONField(
        default=list, 
        blank=True,
        help_text="Timestamped log of processing activities for audit trail"
    )
    
    # Processing metrics
    individual_records = models.PositiveIntegerField(default=0)
    corporate_records = models.PositiveIntegerField(default=0)
    total_records = models.PositiveIntegerField(default=0)
    individual_credit_matched = models.PositiveIntegerField(default=0, help_text="Number of individual records with credit matches")
    corporate_credit_matched = models.PositiveIntegerField(default=0, help_text="Number of corporate records with credit matches")
    unmatched_credit_records = models.PositiveIntegerField(default=0, help_text="Number of credit records that could not be matched to any borrower")
    processing_time = models.FloatField(null=True, blank=True, help_text="Processing time in seconds")
    
    # Verification tracking fields
    has_verification_candidates = models.BooleanField(
        default=False,
        help_text="Whether this upload has records requiring human verification"
    )
    commercial_candidates_count = models.IntegerField(
        default=0,
        help_text="Number of potential commercial entities found in consumer records"
    )
    consumer_candidates_count = models.IntegerField(
        default=0,
        help_text="Number of potential consumer entities found in commercial records"
    )
    verification_completed_at = models.DateTimeField(
        null=True, 
        blank=True,
        help_text="Timestamp when human verification was completed"
    )
    
    # Timestamps
    uploaded_at = models.DateTimeField(auto_now_add=True)
    processing_started_at = models.DateTimeField(null=True, blank=True)
    completed_at = models.DateTimeField(null=True, blank=True)
    
    # Error tracking
    error_message = models.TextField(blank=True, null=True)
    
    # File paths for downloads
    individual_file_path = models.CharField(max_length=500, blank=True, null=True)
    corporate_file_path = models.CharField(max_length=500, blank=True, null=True)
    individual_txt_path = models.CharField(max_length=500, blank=True, null=True)
    corporate_txt_path = models.CharField(max_length=500, blank=True, null=True)
    
    # Output format tracking (for smart format selection)
    individual_output_format = models.CharField(
        max_length=10,
        blank=True,
        null=True,
        help_text="Output format for individual sheet: 'xlsx' or 'txt'"
    )
    corporate_output_format = models.CharField(
        max_length=10,
        blank=True,
        null=True,
        help_text="Output format for corporate sheet: 'xlsx' or 'txt'"
    )
    
    # Original file archival for dispute resolution and compliance
    original_file_path = models.CharField(
        max_length=500, 
        blank=True, 
        null=True,
        help_text="Path to the original unmodified uploaded file for dispute resolution and audit purposes"
    )
    
    # Processing data storage (for async workflow and session recovery)
    processing_data = models.TextField(
        blank=True,
        null=True,
        help_text="JSON storage for processing data when using async workflow"
    )
    
    # Reporting period fields (user-selected date for output filename generation)
    reporting_month = models.IntegerField(
        null=True,
        blank=True,
        help_text="User-selected reporting month (1-12) for output filename generation"
    )
    reporting_year = models.IntegerField(
        null=True,
        blank=True,
        help_text="User-selected reporting year for output filename generation"
    )
    
    class Meta:
        db_table = 'upload_sessions'
        verbose_name = 'Upload Session'
        verbose_name_plural = 'Upload Sessions'
        ordering = ['-uploaded_at']
        indexes = [
            models.Index(fields=['user', 'uploaded_at']),
            models.Index(fields=['subscriber_id', 'uploaded_at']),
            models.Index(fields=['status', 'uploaded_at']),
        ]
    
    def __str__(self):
        return f"{self.original_filename} - {self.user.username} ({self.status})"
    
    def get_subscriber(self):
        """Get the related subscriber object manually"""
        try:
            return Subscriber.objects.get(subscriber_id=self.subscriber_id)
        except Subscriber.DoesNotExist:
            return None
    
    def mark_processing_started(self):
        """Mark the upload as processing started"""
        self.status = 'processing'
        self.processing_stage = 'upload_validation'
        self.progress_percentage = 5
        self.processing_started_at = timezone.now()
        self.add_activity_log('Processing started')
        self.save(update_fields=['status', 'processing_stage', 'progress_percentage', 'processing_started_at'])
    
    def update_progress(self, stage, percentage, message=''):
        """
        Update processing progress with stage, percentage, and optional message
        
        Args:
            stage: Processing stage from STAGE_CHOICES
            percentage: Progress percentage (0-100)
            message: Optional message describing current activity
        """
        self.processing_stage = stage
        self.progress_percentage = percentage
        if message:
            self.current_message = message
            self.add_activity_log(message)
        
        # Update status based on stage
        if stage == 'awaiting_verification':
            self.status = 'awaiting_verification'
        elif stage == 'post_verification':
            self.status = 'finalizing'
        elif stage == 'completed':
            self.status = 'completed'
        elif stage != 'pending':
            self.status = 'processing'
        
        self.save(update_fields=['processing_stage', 'progress_percentage', 'current_message', 'status'])
    
    def add_activity_log(self, message):
        """
        Add a timestamped entry to the activity log
        
        Args:
            message: Activity message to log
        """
        from datetime import datetime
        if not isinstance(self.activity_log, list):
            self.activity_log = []
        
        log_entry = {
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'message': message,
            'stage': self.processing_stage,
            'progress': self.progress_percentage
        }
        self.activity_log.append(log_entry)
        self.save(update_fields=['activity_log'])
    
    def mark_awaiting_verification(self, commercial_count=0, consumer_count=0):
        """
        Mark the upload as awaiting human verification
        
        Args:
            commercial_count: Number of potential commercial entities in consumer records
            consumer_count: Number of potential consumer entities in commercial records
        """
        self.status = 'awaiting_verification'
        self.processing_stage = 'awaiting_verification'
        self.progress_percentage = 80
        self.has_verification_candidates = True
        self.commercial_candidates_count = commercial_count
        self.consumer_candidates_count = consumer_count
        self.current_message = f'Verification required: {commercial_count} commercial candidates, {consumer_count} consumer candidates'
        self.add_activity_log(self.current_message)
        self.save(update_fields=[
            'status', 'processing_stage', 'progress_percentage', 
            'has_verification_candidates', 'commercial_candidates_count', 
            'consumer_candidates_count', 'current_message'
        ])
    
    def mark_verification_completed(self):
        """Mark that human verification has been completed"""
        self.verification_completed_at = timezone.now()
        self.status = 'finalizing'
        self.processing_stage = 'post_verification'
        self.progress_percentage = 85
        self.current_message = 'Verification completed, finalizing processing'
        self.add_activity_log('Human verification completed')
        self.save(update_fields=[
            'verification_completed_at', 'status', 'processing_stage', 
            'progress_percentage', 'current_message'
        ])
    
    def mark_completed(self, individual_count=0, corporate_count=0, processing_time=None):
        """Mark the upload as completed with metrics"""
        self.status = 'completed'
        self.processing_stage = 'completed'
        self.progress_percentage = 100
        self.completed_at = timezone.now()
        self.individual_records = individual_count
        self.corporate_records = corporate_count
        self.total_records = individual_count + corporate_count
        if processing_time:
            self.processing_time = processing_time
        self.current_message = 'Processing completed successfully'
        self.add_activity_log(f'Completed: {self.total_records} total records processed')
        self.save(update_fields=[
            'status', 'processing_stage', 'progress_percentage', 'completed_at', 
            'individual_records', 'corporate_records', 'total_records', 
            'processing_time', 'current_message'
        ])
    
    def mark_failed(self, error_message=None):
        """Mark the upload as failed"""
        self.status = 'failed'
        self.progress_percentage = 0
        self.completed_at = timezone.now()
        if error_message:
            self.error_message = error_message
            self.current_message = f'Processing failed: {error_message[:200]}'
            self.add_activity_log(f'ERROR: {error_message[:200]}')
        self.save(update_fields=['status', 'progress_percentage', 'completed_at', 'error_message', 'current_message'])
    
    @classmethod
    def get_user_stats(cls, user, days=None):
        """Get upload statistics for a user"""
        queryset = cls.objects.filter(user=user)
        
        if days:
            from datetime import timedelta
            start_date = timezone.now() - timedelta(days=days)
            queryset = queryset.filter(uploaded_at__gte=start_date)
        
        stats = {
            'total_uploads': queryset.count(),
            'successful_uploads': queryset.filter(status='completed').count(),
            'failed_uploads': queryset.filter(status='failed').count(),
            'pending_uploads': queryset.filter(status__in=['pending', 'processing']).count(),
            'total_records_processed': sum(queryset.filter(status='completed').values_list('total_records', flat=True)),
            'individual_records': sum(queryset.filter(status='completed').values_list('individual_records', flat=True)),
            'corporate_records': sum(queryset.filter(status='completed').values_list('corporate_records', flat=True)),
            'unmatched_credit_records': sum(queryset.filter(status='completed').values_list('unmatched_credit_records', flat=True)),
        }
        
        # Calculate success rate
        if stats['total_uploads'] > 0:
            stats['success_rate'] = (stats['successful_uploads'] / stats['total_uploads']) * 100
        else:
            stats['success_rate'] = 0
        
        return stats
    
    @classmethod
    def get_recent_uploads(cls, user, limit=10):
        """Get recent uploads for a user"""
        return cls.objects.filter(user=user).order_by('-uploaded_at')[:limit]