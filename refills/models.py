from django.db import models
from datetime import timedelta
from django.utils import timezone
from decimal import Decimal

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

    # Choices for refill status
    STATUS_CHOICES = (
        ('Active', 'Active'),
        ('Active Restart', 'Active Restart'),
        ('Inactive', 'Inactive'),
    )

    # Fields
    facility = models.ForeignKey(
        'Facility',
        on_delete=models.CASCADE,
        related_name="refills"
    )
    unique_id = models.CharField(max_length=100)
    last_pickup_date = models.DateField()
    sex = models.CharField(max_length=10, choices=SEX_CHOICES)
    
    # ✅ Changed from IntegerField to DecimalField to allow decimals
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
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        unique_together = ('facility', 'unique_id')
        ordering = ['-last_pickup_date']
        indexes = [
            models.Index(fields=['facility']),
            models.Index(fields=['next_appointment']),
            models.Index(fields=['last_pickup_date']),
        ]

    def calculate_dates(self):
        """
        Calculate the next appointment and expected IIT.
        - next_appointment = last_pickup_date + months_of_refill_days
        - expected_iit_date = next_appointment + 28 days
        """
        if self.last_pickup_date and self.months_of_refill_days:
            days = float(self.months_of_refill_days) * 30  # ✅ handle decimals
            self.next_appointment = self.last_pickup_date + timedelta(days=days)
            self.expected_iit_date = self.next_appointment + timedelta(days=28)

    def save(self, *args, **kwargs):
        today = timezone.now().date()
        self.calculate_dates()

        if self.next_appointment and self.next_appointment < today:
            self.missed_appointment = True

        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.unique_id} - {self.facility.name}"
