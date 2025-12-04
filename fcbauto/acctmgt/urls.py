from django.urls import path
from django.contrib.auth import views as auth_views
from django.urls import reverse_lazy
from . import views

app_name = 'acctmgt'

urlpatterns = [
    # Registration
    path('register/', views.RegisterView.as_view(), name='register'),
    
    # Login/Logout
    path('login/', views.CustomLoginView.as_view(), name='login'),
    
    path('logout/', auth_views.LogoutView.as_view(
        # Ensure users land on the Login section after logout
        next_page='/acctmgt/login/?section=login'
    ), name='logout'),
    
    path('subscriber-selection/', views.SubscriberSelectionView.as_view(), name='subscriber_selection'),
    path('clear-subscriber/', views.clear_subscriber_session, name='clear_subscriber'),
]
