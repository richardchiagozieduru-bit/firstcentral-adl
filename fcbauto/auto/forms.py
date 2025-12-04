from django import forms
import calendar
from datetime import datetime


def get_default_month():
    """Get default month (previous month)"""
    now = datetime.now()
    return now.month - 1 if now.month > 1 else 12


def get_default_year():
    """Get default year (current year, or previous if January)"""
    now = datetime.now()
    return now.year if now.month > 1 else now.year - 1


class ExcelUploadForm(forms.Form):
    file = forms.FileField(
        label='Upload a file',
        widget=forms.FileInput(attrs={
            'accept': '.xlsx,.xls',
            'class': 'form-control'
        })
    )
    
    # Month picker with smart default (previous month)
    month = forms.ChoiceField(
        label='Reporting Month',
        choices=[(i, calendar.month_name[i]) for i in range(1, 13)],
        required=True,
        initial=get_default_month,
        widget=forms.Select(attrs={'class': 'form-select'}),
        help_text='Select the reporting period month (defaults to previous month)'
    )
    
    # Year picker with smart default
    year = forms.ChoiceField(
        label='Reporting Year',
        choices=[(y, str(y)) for y in range(datetime.now().year - 5, datetime.now().year + 1)],
        required=True,
        initial=get_default_year,
        widget=forms.Select(attrs={'class': 'form-select'}),
        help_text='Select the reporting period year'
    )

