from django.contrib import admin
from .models import Refill, Facility


@admin.register(Facility)
class FacilityAdmin(admin.ModelAdmin):
    list_display = ("name", "code", "location")


@admin.register(Refill)
class RefillAdmin(admin.ModelAdmin):
    list_display = (
        "unique_id",
        "facility",
        "sex",
        "last_pickup_date",
        "months_of_refill_days",
        "next_appointment",
        "case_manager",
    )
    list_filter = ("facility", "sex", "months_of_refill_days")
    search_fields = ("unique_id",)
