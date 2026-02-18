from django import forms
from .models import Refill, Facility

# -----------------------
# Refill Form (existing)
# -----------------------


class RefillForm(forms.ModelForm):
    class Meta:
        model = Refill
        fields = [
            'facility',
            'unique_id',
            'last_pickup_date',
            'sex',
            'months_of_refill_days',  # ✅ now decimal
            'current_regimen',
            'case_manager',
            'remark',
        ]
        widgets = {
            'last_pickup_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'facility': forms.Select(attrs={'class': 'form-select'}),
            'sex': forms.Select(attrs={'class': 'form-select'}),
            # ✅ step="0.1" to allow decimals
            'months_of_refill_days': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.1'}),
            'remark': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
        }


class UploadExcelForm(forms.Form):
    file = forms.FileField(
        widget=forms.ClearableFileInput(attrs={'class': 'form-control'})
    )

