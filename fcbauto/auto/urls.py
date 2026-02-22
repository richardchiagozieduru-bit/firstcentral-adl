from django.urls import path
from . import views

app_name = 'auto'  

urlpatterns=[
    path('', views.home, name='home'),  
    path('upload/', views.upload_file, name='upload'),  
    path('dashboard/', views.dashboard, name='dashboard'),
    path('api/dashboard/', views.dashboard_api, name='dashboard_api'),
    path('api/engagement/', views.engagement_metrics_api, name='engagement_api'),
    path('verify-split/', views.verify_split_decision, name='verify_split_decision'),
    path('export-report/<int:session_id>/', views.generate_excel_report, name='export_report'),
    path('api/delete-upload/<int:upload_id>/', views.delete_upload, name='delete_upload'),
    path('api/cancel-upload/<int:upload_id>/', views.cancel_upload, name='cancel_upload'),
    
    # Progress tracking endpoints
    path('progress/<int:session_id>/', views.progress_tracking, name='progress_tracking'),
    path('api/progress/<int:session_id>/', views.progress_api, name='progress_api'),
    
    # Feedback
    path('api/feedback/', views.submit_feedback, name='submit_feedback'),
]
