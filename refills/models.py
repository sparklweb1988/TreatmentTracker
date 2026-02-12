from django.db import models
from datetime import timedelta
from django.utils import timezone

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

    # Choices for refill duration
    REFILL_DAY_CHOICES = (
        (30, "30 Days"),
        (60, "60 Days"),
        (90, "90 Days"),
        (180, "180 Days"),
    )

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

    # Fields related to the refill
    facility = models.ForeignKey(
        'Facility',
        on_delete=models.CASCADE,
        related_name="refills"
    )
    unique_id = models.CharField(max_length=100)
    last_pickup_date = models.DateField()
    sex = models.CharField(max_length=10, choices=SEX_CHOICES)
    months_of_refill_days = models.IntegerField(
        choices=REFILL_DAY_CHOICES
    )
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

    # Missed Appointment Flag (New Field)
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
        Calculate the next appointment and IIT date based on the refill days.
        """
        if self.last_pickup_date and self.months_of_refill_days:
            self.next_appointment = self.last_pickup_date + timedelta(
                days=self.months_of_refill_days
            )
            self.expected_iit_date = self.next_appointment + timedelta(days=1)

    def save(self, *args, **kwargs):
        """
        Save the refill, and calculate dates if applicable.
        If the next appointment has passed and the appointment is missed, set missed_appointment to True.
        """
        today = timezone.now().date()  # Get today's date

        # Check if the next appointment has passed, and mark as missed if needed
        if self.next_appointment and self.next_appointment < today:
            self.missed_appointment = True

        # Recalculate the appointment dates
        self.calculate_dates()

        # Call the parent save method
        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.unique_id} - {self.facility.name}"

