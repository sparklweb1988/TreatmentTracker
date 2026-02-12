from django import forms
from .models import Refill, Facility



class RefillForm(forms.ModelForm):
    class Meta:
        model = Refill
        fields = [
            'facility',
            'unique_id',
            'last_pickup_date',
            'sex',
            'months_of_refill_days',
            'current_regimen',
            'case_manager',
            'remark',
        ]
        widgets = {
            'last_pickup_date': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'facility': forms.Select(attrs={'class': 'form-select'}),
            'sex': forms.Select(attrs={'class': 'form-select'}),
            'months_of_refill_days': forms.Select(attrs={'class': 'form-select'}),
            'remark': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Unique ID dropdown - get unique values in Python
        unique_ids = list({ref.unique_id for ref in Refill.objects.all()})
        self.fields['unique_id'] = forms.ChoiceField(
            choices=[('', 'Select Unique ID')] + [(uid, uid) for uid in unique_ids],
            widget=forms.Select(attrs={'class': 'form-select'})
        )

        # Current Regimen dropdown - unique values
        regimens = list({ref.current_regimen for ref in Refill.objects.all() if ref.current_regimen})
        self.fields['current_regimen'] = forms.ChoiceField(
            choices=[('', 'Select Regimen')] + [(r, r) for r in regimens],
            widget=forms.Select(attrs={'class': 'form-select'})
        )

        # Case Manager dropdown - unique values
        case_managers = list({ref.case_manager for ref in Refill.objects.all() if ref.case_manager})
        self.fields['case_manager'] = forms.ChoiceField(
            choices=[('', 'Select Case Manager')] + [(c, c) for c in case_managers],
            widget=forms.Select(attrs={'class': 'form-select'})
        )






class UploadExcelForm(forms.Form):
    facility = forms.ModelChoiceField(
        queryset=Facility.objects.all(),
        widget=forms.Select(attrs={'class': 'form-select'})
    )
    file = forms.FileField(
        widget=forms.FileInput(attrs={'class': 'form-control'})
    )
