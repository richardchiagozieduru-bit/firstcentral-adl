from .models import UserProfile

def subscriber_context(request):
    """
    Context processor to add bound subscriber information to all templates.
    This makes subscriber information available globally across the application.
    """
    context = {}
    
    if request.user.is_authenticated:
        try:
            user_profile = UserProfile.get_or_create_profile(request.user)
            if user_profile.is_bound:
                bound_subscriber = user_profile.get_bound_subscriber()
                if bound_subscriber:
                    context.update({
                        'bound_subscriber': bound_subscriber,
                        'subscriber_name': bound_subscriber.subscriber_name,
                        'subscriber_id': bound_subscriber.subscriber_id,
                        'user_is_bound': True,
                        'binding_date': user_profile.bound_at,
                        'binding_method': user_profile.get_binding_method_display(),
                    })
                else:
                    context['user_is_bound'] = False
            else:
                context['user_is_bound'] = False
        except Exception:
            # Handle any database or model errors gracefully
            context['user_is_bound'] = False
    else:
        context['user_is_bound'] = False
    
    return context