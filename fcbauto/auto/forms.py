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


class MultipleFileInput(forms.ClearableFileInput):
    """Custom widget that allows multiple file selection"""
    allow_multiple_selected = True


class MultipleFileField(forms.FileField):
    """Custom field that handles multiple file uploads"""
    
    def __init__(self, *args, **kwargs):
        kwargs.setdefault("widget", MultipleFileInput())
        super().__init__(*args, **kwargs)

    def clean(self, data, initial=None):
        single_file_clean = super().clean
        if isinstance(data, (list, tuple)):
            result = [single_file_clean(d, initial) for d in data]
        else:
            result = [single_file_clean(data, initial)]
        return result


class ExcelUploadForm(forms.Form):
    file = MultipleFileField(
        label='Upload files',
        widget=MultipleFileInput(attrs={
            'accept': '.xlsx,.xls,.xlsb,.xlsm,.csv,.txt',
            'class': 'form-control',
            'multiple': True
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
    
    # Optional subscriber_id for multi-subscriber users (hidden for regular users)
    subscriber_id = forms.IntegerField(
        required=False,
        widget=forms.HiddenInput()
    )
    
    def clean_file(self):
        """
        Validate uploaded files are genuine Excel or CSV files.

        - Excel files are validated with magic bytes detection.
        - CSV files are validated as plain text (no magic bytes).
        - Mixing CSV and Excel files in the same upload is not allowed.
        """
        uploaded_files = self.cleaned_data.get('file')

        if not uploaded_files:
            return uploaded_files

        # Ensure it's a list
        if not isinstance(uploaded_files, list):
            uploaded_files = [uploaded_files]

        from .file_validators import validate_excel_file_type, validate_csv_file

        # Detect whether any file is CSV or Excel
        csv_count = sum(1 for f in uploaded_files if f and f.name.lower().endswith(('.csv', '.txt')))
        excel_count = len(uploaded_files) - csv_count

        if csv_count > 0 and excel_count > 0:
            raise forms.ValidationError(
                "Please upload either CSV files or Excel files in a single upload — not a mix of both."
            )

        validated_files = []
        for uploaded_file in uploaded_files:
            if uploaded_file:
                if uploaded_file.name.lower().endswith(('.csv', '.txt')):
                    is_valid, error_message = validate_csv_file(uploaded_file)
                else:
                    is_valid, error_message = validate_excel_file_type(uploaded_file)

                if not is_valid:
                    raise forms.ValidationError(f"{uploaded_file.name}: {error_message}")

                # Reset file pointer after validation
                uploaded_file.seek(0)
                validated_files.append(uploaded_file)

        return validated_files

