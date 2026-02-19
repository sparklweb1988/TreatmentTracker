from django.db import models
from datetime import timedelta
from django.utils import timezone
from decimal import Decimal
from django.db.models import F, Q


class Facility(models.Model):
    name = models.CharField(max_length=255, unique=True)
    code = models.CharField(max_length=50, unique=True)
    location = models.CharField(max_length=255, blank=True, null=True)

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']

    def __str__(self):
        return self.name









class Refill(models.Model):

    # Choices for gender
    SEX_CHOICES = (
        ('Male', 'Male'),
        ('Female', 'Female'),
    )

    # Choices for ART status
    STATUS_CHOICES = (
        ('Active', 'Active'),
        ('Active Restart', 'Active Restart'),
        ('Inactive', 'Inactive'),
    )

    # ===================== FIELDS =====================
    facility = models.ForeignKey(
        'Facility', on_delete=models.CASCADE, related_name="refills"
    )
    unique_id = models.CharField(max_length=100)
    last_pickup_date = models.DateField(null=True, blank=True)
    sex = models.CharField(max_length=10, choices=SEX_CHOICES)
    months_of_refill_days = models.DecimalField(max_digits=4, decimal_places=2)
    current_regimen = models.CharField(max_length=255)
    case_manager = models.CharField(max_length=255)
    remark = models.TextField(blank=True, null=True)
    current_art_status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES,
        default='Active'
    )
    next_appointment = models.DateField(blank=True, null=True)
    expected_iit_date = models.DateField(blank=True, null=True)
    missed_appointment = models.BooleanField(default=False)

    # ===================== VIRAL LOAD FIELDS =====================
    art_start_date = models.DateField(blank=True, null=True)
    vl_sample_collection_date = models.DateField(blank=True, null=True)
    vl_result = models.IntegerField(blank=True, null=True)  # copies/ml

    created_at = models.DateTimeField(auto_now_add=True)

    # ===================== META =====================
    class Meta:
        unique_together = ('facility', 'unique_id')
        ordering = ['-last_pickup_date']
        indexes = [
            models.Index(fields=['facility']),
            models.Index(fields=['next_appointment']),
            models.Index(fields=['last_pickup_date']),
        ]

    # ===================== CALCULATE REFILL DATES =====================
    def calculate_dates(self):
        """
        next_appointment = last_pickup_date + months_of_refill_days
        expected_iit_date = next_appointment + 28 days
        """
        if self.last_pickup_date and self.months_of_refill_days:
            days = float(self.months_of_refill_days) * 30
            self.next_appointment = self.last_pickup_date + timedelta(days=days)
            self.expected_iit_date = self.next_appointment + timedelta(days=28)

    # ===================== SAVE =====================
    def save(self, *args, **kwargs):
        today = timezone.now().date()
        self.calculate_dates()

        if self.next_appointment and self.next_appointment < today:
            self.missed_appointment = True

        super().save(*args, **kwargs)

    # ===================== VIRAL LOAD LOGIC =====================
    @property
    def is_vl_eligible(self):
        """
        Eligible for VL if:
        - On ART >= 180 days
        - No VL sample in current year (adults)
        - Children (<15) get 2 samples per year (6 months apart)
        """
        if not self.art_start_date:
            return False

        today = timezone.now().date()
        days_on_art = (today - self.art_start_date).days
        if days_on_art < 180:
            return False

        # Determine age (approximate using ART start date)
        age_years = (today - self.art_start_date).days // 365
        vl_date = self.vl_sample_collection_date

        if age_years >= 15:  # Adult
            if vl_date and vl_date.year == today.year:
                return False
            return True
        else:  # Child
            if vl_date and (today - vl_date).days < 180:
                return False
            return True
        
        
    @property
    def vl_status(self):
        """
        Return human-readable VL eligibility status for template.
        """
        return "Eligible" if self.is_vl_eligible else "Not Eligible"


    @property
    def is_suppressed(self):
        """
        Suppressed if VL result < 1000
        """
        if self.vl_result is None:
            return None
        return self.vl_result < 1000

    # ===================== QUARTER CALCULATION =====================
    @staticmethod
    def get_quarter(dt):
        if not dt:
            return None
        month = dt.month
        if month in [1,2,3]:
            return "Q1"
        elif month in [4,5,6]:
            return "Q2"
        elif month in [7,8,9]:
            return "Q3"
        else:
            return "Q4"

    @classmethod
    def calculate_quarterly_vl_coverage(cls, year, quarter, facility=None):
        refills = cls.objects.filter(
            current_art_status__in=['Active', 'Active Restart'],
            art_start_date__isnull=False
        )

        if facility:
            refills = refills.filter(facility=facility)

        # Eligible: ≥180 days on ART, ART start before quarter end
        quarter_start_month = {"Q1":1, "Q2":4, "Q3":7, "Q4":10}[quarter]
        quarter_start = timezone.datetime(year, quarter_start_month, 1).date()
        quarter_end = (timezone.datetime(year, quarter_start_month+3, 1).date() - timedelta(days=1)
                       if quarter in ["Q1","Q2","Q3"]
                       else timezone.datetime(year+1, 1, 1).date() - timedelta(days=1))

        denominator_qs = refills.filter(
            art_start_date__lte=quarter_end,
        )
        denominator = denominator_qs.count()

        numerator = denominator_qs.filter(
            vl_sample_collection_date__range=[quarter_start, quarter_end]
        ).count()

        coverage = (numerator / denominator * 100) if denominator > 0 else 0
        return {"denominator": denominator, "numerator": numerator, "coverage": round(coverage,1)}

    def __str__(self):
        return f"{self.unique_id} - {self.facility.name}"
