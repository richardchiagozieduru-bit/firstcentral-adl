from django.core.management.base import BaseCommand, CommandError
from django.contrib.auth.models import User
from django.db import models
from acctmgt.models import UserProfile
from auto.models import Subscriber
from django.utils import timezone

class Command(BaseCommand):
    help = 'Manage user-subscriber bindings for administrators'
    
    def add_arguments(self, parser):
        parser.add_argument(
            '--list-bindings',
            action='store_true',
            help='List all user bindings',
        )
        parser.add_argument(
            '--list-unbound',
            action='store_true',
            help='List all unbound users',
        )
        parser.add_argument(
            '--unbind-user',
            type=str,
            help='Unbind a specific user (provide username)',
        )
        parser.add_argument(
            '--bind-user',
            type=str,
            help='Bind a user to a subscriber (provide username)',
        )
        parser.add_argument(
            '--subscriber-id',
            type=int,
            help='Subscriber ID for binding operations',
        )
        parser.add_argument(
            '--reason',
            type=str,
            default='admin_action',
            help='Reason for unbinding (default: admin_action)',
        )
        parser.add_argument(
            '--stats',
            action='store_true',
            help='Show binding statistics',
        )
    
    def handle(self, *args, **options):
        if options['list_bindings']:
            self.list_bindings()
        elif options['list_unbound']:
            self.list_unbound_users()
        elif options['unbind_user']:
            self.unbind_user(options['unbind_user'], options['reason'])
        elif options['bind_user']:
            if not options['subscriber_id']:
                raise CommandError('--subscriber-id is required when binding a user')
            self.bind_user(options['bind_user'], options['subscriber_id'])
        elif options['stats']:
            self.show_stats()
        else:
            self.stdout.write(
                self.style.WARNING(
                    'Please specify an action: --list-bindings, --list-unbound, --unbind-user, --bind-user, or --stats'
                )
            )
    
    def list_bindings(self):
        """List all user bindings"""
        self.stdout.write(self.style.SUCCESS('\n=== USER BINDINGS ==='))
        
        bound_profiles = UserProfile.objects.filter(is_bound=True).select_related('user')
        
        if not bound_profiles.exists():
            self.stdout.write(self.style.WARNING('No bound users found.'))
            return
        
        self.stdout.write(f'\nFound {bound_profiles.count()} bound users:\n')
        
        for profile in bound_profiles:
            subscriber = profile.get_bound_subscriber()
            subscriber_name = subscriber.subscriber_name if subscriber else f"ID:{profile.bound_subscriber_id}"
            
            self.stdout.write(
                f"• {profile.user.username} → {subscriber_name}"
            )
            self.stdout.write(
                f"  Bound: {profile.bound_at.strftime('%Y-%m-%d %H:%M:%S') if profile.bound_at else 'Unknown'}"
            )
            self.stdout.write(
                f"  Method: {profile.get_binding_method_display()}"
            )
            self.stdout.write(
                f"  Token: {profile.binding_token or 'N/A'}"
            )
            self.stdout.write('')
    
    def list_unbound_users(self):
        """List all unbound users"""
        self.stdout.write(self.style.SUCCESS('\n=== UNBOUND USERS ==='))
        
        # Get all users who don't have a bound profile
        unbound_users = User.objects.filter(
            models.Q(profile__isnull=True) | models.Q(profile__is_bound=False)
        )
        
        if not unbound_users.exists():
            self.stdout.write(self.style.WARNING('No unbound users found.'))
            return
        
        self.stdout.write(f'\nFound {unbound_users.count()} unbound users:\n')
        
        for user in unbound_users:
            profile = getattr(user, 'profile', None)
            if profile:
                created_date = profile.created_at.strftime('%Y-%m-%d %H:%M:%S')
                self.stdout.write(f"• {user.username} (profile created: {created_date})")
            else:
                self.stdout.write(f"• {user.username} (no profile)")
    
    def unbind_user(self, username, reason):
        """Unbind a specific user"""
        try:
            user = User.objects.get(username=username)
            profile = UserProfile.get_or_create_profile(user)
            
            if not profile.is_bound:
                self.stdout.write(
                    self.style.WARNING(f'User "{username}" is not bound to any subscriber.')
                )
                return
            
            subscriber = profile.get_bound_subscriber()
            subscriber_name = subscriber.subscriber_name if subscriber else f"ID:{profile.bound_subscriber_id}"
            
            # Confirm the action
            self.stdout.write(
                f'About to unbind user "{username}" from "{subscriber_name}"'
            )
            self.stdout.write(f'Reason: {reason}')
            
            confirm = input('Are you sure? (yes/no): ')
            if confirm.lower() != 'yes':
                self.stdout.write(self.style.WARNING('Operation cancelled.'))
                return
            
            # Perform unbinding
            success = profile.unbind_subscriber(reason)
            
            if success:
                self.stdout.write(
                    self.style.SUCCESS(
                        f'Successfully unbound user "{username}" from "{subscriber_name}"'
                    )
                )
            else:
                self.stdout.write(
                    self.style.ERROR(f'Failed to unbind user "{username}"')
                )
                
        except User.DoesNotExist:
            raise CommandError(f'User "{username}" does not exist')
        except Exception as e:
            raise CommandError(f'Error unbinding user: {e}')
    
    def bind_user(self, username, subscriber_id):
        """Bind a user to a subscriber (admin assignment)"""
        try:
            user = User.objects.get(username=username)
            profile = UserProfile.get_or_create_profile(user)
            
            if profile.is_bound:
                current_subscriber = profile.get_bound_subscriber()
                current_name = current_subscriber.subscriber_name if current_subscriber else f"ID:{profile.bound_subscriber_id}"
                self.stdout.write(
                    self.style.WARNING(
                        f'User "{username}" is already bound to "{current_name}". '
                        f'Use --unbind-user first if you want to change the binding.'
                    )
                )
                return
            
            # Verify subscriber exists
            try:
                subscriber = Subscriber.objects.get(subscriber_id=subscriber_id)
            except Subscriber.DoesNotExist:
                raise CommandError(f'Subscriber with ID {subscriber_id} does not exist')
            
            # Confirm the action
            self.stdout.write(
                f'About to bind user "{username}" to "{subscriber.subscriber_name}" (ID: {subscriber_id})'
            )
            
            confirm = input('Are you sure? (yes/no): ')
            if confirm.lower() != 'yes':
                self.stdout.write(self.style.WARNING('Operation cancelled.'))
                return
            
            # Perform binding
            success = profile.bind_to_subscriber(
                subscriber_id=subscriber_id,
                token_string='admin_assigned',
                ip_address=None,
                method='admin'
            )
            
            if success:
                self.stdout.write(
                    self.style.SUCCESS(
                        f'Successfully bound user "{username}" to "{subscriber.subscriber_name}"'
                    )
                )
            else:
                self.stdout.write(
                    self.style.ERROR(f'Failed to bind user "{username}"')
                )
                
        except User.DoesNotExist:
            raise CommandError(f'User "{username}" does not exist')
        except Exception as e:
            raise CommandError(f'Error binding user: {e}')
    
    def show_stats(self):
        """Show binding statistics"""
        self.stdout.write(self.style.SUCCESS('\n=== BINDING STATISTICS ==='))
        
        total_users = User.objects.count()
        bound_users = UserProfile.objects.filter(is_bound=True).count()
        unbound_users = total_users - bound_users
        
        self.stdout.write(f'\nTotal Users: {total_users}')
        self.stdout.write(f'Bound Users: {bound_users}')
        self.stdout.write(f'Unbound Users: {unbound_users}')
        
        if total_users > 0:
            bound_percentage = (bound_users / total_users) * 100
            self.stdout.write(f'Binding Rate: {bound_percentage:.1f}%')
        
        # Binding methods breakdown
        self.stdout.write('\nBinding Methods:')
        for method, display in UserProfile._meta.get_field('binding_method').choices:
            count = UserProfile.objects.filter(binding_method=method, is_bound=True).count()
            if count > 0:
                self.stdout.write(f'  {display}: {count}')
        
        # Recent bindings
        recent_bindings = UserProfile.objects.filter(
            is_bound=True,
            bound_at__gte=timezone.now() - timezone.timedelta(days=7)
        ).count()
        
        self.stdout.write(f'\nRecent Bindings (last 7 days): {recent_bindings}')
        
        # Subscribers with most users
        self.stdout.write('\nTop Subscribers by User Count:')
        from django.db import models
        subscriber_counts = UserProfile.objects.filter(is_bound=True).values(
            'bound_subscriber_id'
        ).annotate(
            user_count=models.Count('user')
        ).order_by('-user_count')[:5]
        
        for item in subscriber_counts:
            try:
                subscriber = Subscriber.objects.get(subscriber_id=item['bound_subscriber_id'])
                name = subscriber.subscriber_name
            except Subscriber.DoesNotExist:
                name = f"ID:{item['bound_subscriber_id']}"
            
            self.stdout.write(f'  {name}: {item["user_count"]} users')