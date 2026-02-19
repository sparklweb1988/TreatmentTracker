from django import forms
from .models import Refill, Facility
from django.utils import timezone
from datetime import timedelta


# -----------------------
# Refill Form (existing)
# -----------------------





class RefillForm(forms.ModelForm):
    class Meta:
        model = Refill
        fields = [
            'facility',
            'unique_id',
            'art_start_date',                  # ART Start Date
            'vl_sample_collection_date',       # Viral Load Sample Collection
            'vl_result',                        # Viral Load Result
            'last_pickup_date',
            'sex',
            'months_of_refill_days',           # now decimal
            'current_regimen',
            'case_manager',
            'remark',
        ]

        widgets = {
            'facility': forms.Select(attrs={'class': 'form-select'}),
            'unique_id': forms.TextInput(attrs={'class': 'form-control'}),
            'art_start_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'vl_sample_collection_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'vl_result': forms.NumberInput(attrs={'class': 'form-control', 'placeholder': 'copies/ml'}),
            'last_pickup_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'sex': forms.Select(attrs={'class': 'form-select'}),
            'months_of_refill_days': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.1'}),
            'current_regimen': forms.TextInput(attrs={'class': 'form-control'}),
            'case_manager': forms.TextInput(attrs={'class': 'form-control'}),
            'remark': forms.Textarea(attrs={'class': 'form-control', 'rows': 3, 'placeholder': 'Optional'}),
        }

    def clean_vl_result(self):
        vl = self.cleaned_data.get('vl_result')
        if vl is not None and vl < 0:
            raise forms.ValidationError("Viral Load cannot be negative.")
        return vl

    def save(self, commit=True):
        """
        Override save to automatically calculate:
        - VL eligibility (≥180 days on ART)
        - VL status (Done / Eligible - Pending / Not Eligible)
        """
        instance = super().save(commit=False)

        today = timezone.now().date()

        # Calculate VL eligibility
        instance.vl_eligible = False
        if instance.art_start_date and (today - instance.art_start_date).days >= 180 and today.year >= 2026:
            instance.vl_eligible = True

        # Determine current quarter
        def get_quarter(date):
            if not date:
                return None
            month = date.month
            if month in [1, 2, 3]:
                return "Q1"
            elif month in [4, 5, 6]:
                return "Q2"
            elif month in [7, 8, 9]:
                return "Q3"
            else:
                return "Q4"

        current_quarter = get_quarter(today)

        # Calculate VL status
        if not instance.vl_eligible:
            instance.vl_status = "Not Eligible"
        else:
            sample_date = instance.vl_sample_collection_date
            if sample_date and get_quarter(sample_date) == current_quarter and sample_date.year == today.year:
                instance.vl_status = "Done"
            else:
                instance.vl_status = "Eligible - Pending"

        if commit:
            instance.save()
        return instance



class UploadExcelForm(forms.Form):
    file = forms.FileField(
        widget=forms.ClearableFileInput(attrs={'class': 'form-control'})
    )

