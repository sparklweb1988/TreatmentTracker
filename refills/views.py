from django.shortcuts import render, redirect, get_object_or_404
from django.utils import timezone
from django.db import transaction
from datetime import timedelta
from django.core.exceptions import ValidationError
from .forms import RefillForm, UploadExcelForm
from .models import Refill, Facility
import pandas as pd
from django.utils import timezone
import openpyxl
from django.http import HttpResponse
from django.conf import settings
from django.db.models import F, Q
from django.core.paginator import Paginator
from openpyxl import Workbook
from datetime import datetime
from .forms import UploadExcelForm



# views.py


VALID_REFILL_MONTHS = [1, 2, 3, 6]

def import_refills_from_excel(file):
    """
    Import refill data from an Excel file containing multiple facilities.
    Old data per facility will be deleted before uploading new rows.
    """

    # ================= FILE SIZE CHECK =================
    if file.size > 1073741824:  # 1GB
        raise ValidationError("File size exceeds the maximum allowed limit of 1GB.")

    # ================= SAFE FILE READ =================
    file.seek(0)
    from io import BytesIO
    file_data = file.read()
    df = pd.read_excel(BytesIO(file_data))

    # ================= REQUIRED COLUMNS =================
    required_columns = [
        'Unique Id',
        'Last Pickup Date (yyyy-mm-dd)',
        'Months of ARV Refill',
        'Current ART Regimen',
        'Case Manager',
        'Sex',
        'Current ART Status',
        'Facility Name'
    ]

    for col in required_columns:
        if col not in df.columns:
            raise ValidationError(f"Missing column: {col}")

    # ================= FILTER ACTIVE PATIENTS =================
    df = df[df['Current ART Status'].isin(['Active', 'Active Restart'])]
    if df.empty:
        raise ValidationError("No Active or Active Restart patients found.")

    # ================= TRACK DELETED FACILITIES =================
    deleted_facilities = set()

    with transaction.atomic():
        for _, row in df.iterrows():
            facility_name = str(row['Facility Name']).strip()
            if not facility_name or facility_name.lower() == "nan":
                raise ValidationError(
                    f"Missing Facility Name for Unique Id {row['Unique Id']}"
                )

            try:
                facility_obj = Facility.objects.get(name__iexact=facility_name)
            except Facility.DoesNotExist:
                raise ValidationError(f"Facility '{facility_name}' does not exist.")
            except Facility.MultipleObjectsReturned:
                raise ValidationError(
                    f"Multiple facilities found for '{facility_name}'. Remove duplicates."
                )

            # ================= DELETE OLD DATA ONCE PER FACILITY =================
            if facility_obj.id not in deleted_facilities:
                Refill.objects.filter(facility=facility_obj).delete()
                deleted_facilities.add(facility_obj.id)

            # ================= VALIDATE LAST PICKUP DATE =================
            if pd.isnull(row['Last Pickup Date (yyyy-mm-dd)']):
                raise ValidationError(
                    f"Missing Last Pickup Date for Unique Id {row['Unique Id']}"
                )
            try:
                last_pickup_date = pd.to_datetime(
                    row['Last Pickup Date (yyyy-mm-dd)']
                ).date()
            except Exception:
                raise ValidationError(
                    f"Invalid Last Pickup Date format for Unique Id {row['Unique Id']}"
                )

            # ================= VALIDATE MONTHS =================
            try:
                months = int(row['Months of ARV Refill'])
            except Exception:
                raise ValidationError(
                    f"Invalid Months of ARV Refill for Unique Id {row['Unique Id']}"
                )

            if months not in VALID_REFILL_MONTHS:
                raise ValidationError(
                    f"Invalid refill duration {months} months "
                    f"for Unique Id {row['Unique Id']}. Allowed: 1, 2, 3, 6"
                )

            refill_days = months * 30
            next_appointment = last_pickup_date + timedelta(days=refill_days)

            # ================= CREATE OR UPDATE =================
            Refill.objects.update_or_create(
                facility=facility_obj,
                unique_id=row['Unique Id'],
                defaults={
                    'last_pickup_date': last_pickup_date,
                    'next_appointment': next_appointment,
                    'months_of_refill_days': refill_days,
                    'current_regimen': row['Current ART Regimen'],
                    'case_manager': str(row['Case Manager']).strip(),
                    'sex': str(row['Sex']).strip(),
                    'current_art_status': row['Current ART Status'].strip(),
                }
            )


def upload_excel(request):
    if request.method == 'POST':
        form = UploadExcelForm(request.POST, request.FILES)
        excel_file = request.FILES.get('file')

        if excel_file and excel_file.size > 1073741824:  # 1GB
            return render(request, 'upload.html', {
                'form': form,
                'error': "File size exceeds the 1GB limit."
            })

        if form.is_valid():
            try:
                import_refills_from_excel(excel_file)
                return redirect('refill_list')
            except ValidationError as e:
                return render(request, 'upload.html', {
                    'form': form,
                    'error': str(e)
                })
        else:
            return render(request, 'upload.html', {'form': form})
    else:
        form = UploadExcelForm()

    return render(request, 'upload.html', {'form': form})

# ================================
# DASHBOARD
# ================================





def dashboard(request):
    today = timezone.now().date()
    week_end = today + timedelta(days=7)

    # First and last day of current month
    month_start = today.replace(day=1)
    if today.month == 12:
        month_end = today.replace(year=today.year+1, month=1, day=1) - timedelta(days=1)
    else:
        month_end = today.replace(month=today.month+1, day=1) - timedelta(days=1)

    facility_id = request.GET.get("facility")
    facilities = Facility.objects.all()
    
    refills = Refill.objects.filter(
        current_art_status__in=['Active', 'Active Restart']  # only active clients
    )

    if facility_id:
        refills = refills.filter(facility_id=facility_id)

    # ✅ NEW: Total Missed for Current Month (up to today)
    monthly_missed_total = refills.filter(
        next_appointment__year=today.year,
        next_appointment__month=today.month,
        next_appointment__lt=today
    ).filter(
        Q(last_pickup_date__lt=F('next_appointment')) |
        Q(last_pickup_date__isnull=True)
    ).count()

    # ================= IIT COUNT (≥ 28 DAYS MISSED) =================
    iit_queryset = refills.filter(
        next_appointment__lt=today
    ).filter(
        Q(last_pickup_date__lt=F('next_appointment')) |
        Q(last_pickup_date__isnull=True)
    )

    iit_total = 0
    for refill in iit_queryset:
        if refill.next_appointment:
            days_missed = (today - refill.next_appointment).days
            if days_missed >= 28:
                iit_total += 1

    context = {
        "facilities": facilities,
        "selected_facility": facility_id,

        # Daily
        "daily_expected": refills.filter(next_appointment=today),
        "daily_refills": refills.filter(last_pickup_date=today),

        # Weekly
        "weekly_expected": refills.filter(next_appointment__range=[today, week_end]),
        "weekly_refills": refills.filter(last_pickup_date__range=[today, week_end]),

        # Monthly (only current month)
        "monthly_expected": refills.filter(
            next_appointment__year=today.year,
            next_appointment__month=today.month
        ),
        "monthly_refills": refills.filter(
            last_pickup_date__year=today.year,
            last_pickup_date__month=today.month
        ),

        # Existing Card
        "monthly_missed_total": monthly_missed_total,

        # 🔥 NEW IIT CARD VALUE
        "iit_total": iit_total,

        "today": today
    }

    return render(request, "dashboard.html", context)


# ================================
# CRUD VIEWS
# ================================





from django.utils import timezone







def refill_list(request):
    today = datetime.now().date()
    week_end = today + timedelta(days=7)

    # GET filters
    facility_id = request.GET.get("facility")
    selected_case_manager = request.GET.get("case_manager")
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")

    # Parse start and end date if provided
    start_date_obj = None
    end_date_obj = None
    if start_date:
        try:
            start_date_obj = datetime.strptime(start_date, "%Y-%m-%d").date()
        except ValueError:
            pass
    if end_date:
        try:
            end_date_obj = datetime.strptime(end_date, "%Y-%m-%d").date()
        except ValueError:
            pass

    # All facilities
    facilities = Facility.objects.all()

    # Base queryset
    refills = Refill.objects.all()

    # Filter by facility if provided
    if facility_id:
        try:
            refills = refills.filter(facility_id=int(facility_id))
        except ValueError:
            pass

    # Filter by case manager if provided
    if selected_case_manager:
        refills = refills.filter(case_manager=selected_case_manager)

    # Filter by date range if provided
    if start_date_obj:
        refills = refills.filter(next_appointment__gte=start_date_obj)
    if end_date_obj:
        refills = refills.filter(next_appointment__lte=end_date_obj)

    # Unique case managers for filter dropdown
    case_managers_qs = (
        Refill.objects.exclude(case_manager__isnull=True)
        .exclude(case_manager__exact="")
        .values_list("case_manager", flat=True)
        .distinct()
    )
    case_managers = sorted({cm.strip() for cm in case_managers_qs if cm.strip()})

    # Calculate days missed and missed_appointment flag
    for refill in refills:
        if refill.next_appointment and (not refill.last_pickup_date or refill.last_pickup_date < refill.next_appointment):
            refill.days_missed = (today - refill.next_appointment).days
            refill.missed_appointment = True
        else:
            refill.days_missed = 0
            refill.missed_appointment = False

    # Group refills by period
    daily_expected = refills.filter(next_appointment=today)
    weekly_expected = refills.filter(next_appointment__range=[today, week_end])
    monthly_expected = refills.filter(next_appointment__year=today.year, next_appointment__month=today.month)

    # Pagination per period
    daily_page = Paginator(daily_expected.order_by("next_appointment"), 10)
    weekly_page = Paginator(weekly_expected.order_by("next_appointment"), 10)
    monthly_page = Paginator(monthly_expected.order_by("next_appointment"), 10)

    # Get current page numbers
    daily_number = request.GET.get("daily_page")
    weekly_number = request.GET.get("weekly_page")
    monthly_number = request.GET.get("monthly_page")

    context = {
        "facilities": facilities,
        "selected_facility": facility_id,
        "case_managers": case_managers,
        "selected_case_manager": selected_case_manager,
        "today": today,
        "selected_start_date": start_date,
        "selected_end_date": end_date,
        "periods": [
            {"name": "Daily", "page_obj": daily_page.get_page(daily_number)},
            {"name": "Weekly", "page_obj": weekly_page.get_page(weekly_number)},
            {"name": "Monthly", "page_obj": monthly_page.get_page(monthly_number)},
        ],
    }

    # Excel export
    if 'download' in request.GET:
        return export_refills_to_excel(refills)

    return render(request, "refill_list.html", context)


def export_refills_to_excel(refills):
    today = datetime.now().date()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Expected Refills Data"

    headers = ['Unique ID', 'Facility', 'Sex', 'Current Regimen', 'Case Manager', 'Last Pickup', 'Next Appointment', 'Days Missed']
    ws.append(headers)

    for refill in refills:
        last_pickup_date = refill.last_pickup_date
        refill_days = refill.months_of_refill_days or 0
        next_appointment = last_pickup_date + timedelta(days=refill_days) if last_pickup_date else None

        row = [
            refill.unique_id,
            refill.facility.name if refill.facility else "",
            refill.sex,
            refill.current_regimen,
            refill.case_manager or "",
            last_pickup_date.strftime("%Y-%m-%d") if last_pickup_date else "Never Picked",
            next_appointment.strftime("%Y-%m-%d") if next_appointment else "",
            getattr(refill, "days_missed", 0),
        ]
        ws.append(row)

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename="Expected_Refills_{today}.xlsx"'
    wb.save(response)
    return response





def refill_create(request, unique_id=None):
    if unique_id:
        # Fetch the refill by unique_id if passed
        refill = get_object_or_404(Refill, unique_id=unique_id)
    else:
        refill = None  # New refill if no unique_id is passed

    # Get today's date, ensuring it's a datetime.date object
    today = timezone.now().date()  # `today` is a datetime.date object
    
    if refill:
        # Ensure that next_appointment is a datetime.date (not a datetime.datetime)
        if isinstance(refill.next_appointment, datetime):  # Check if it's a datetime object
            refill_next_appointment = refill.next_appointment.date()  # Get only the date part
        else:
            refill_next_appointment = refill.next_appointment  # It's already a datetime.date

        # Example comparison (this is just for illustration, customize based on your logic)
        if refill_next_appointment < today:
            # Logic if the refill's next appointment is in the past
            print("This refill's next appointment is in the past.")

    if request.method == 'POST':
        if refill:
            form = RefillForm(request.POST, instance=refill)  # Edit the existing refill
        else:
            form = RefillForm(request.POST)  # Create a new refill

        if form.is_valid():
            form.save()
            return redirect('daily_refill_list')  # After saving, redirect to the daily refill list

    else:
        if refill:
            form = RefillForm(instance=refill)  # Pre-fill the form if editing
        else:
            form = RefillForm()  # Empty form for new refill

    return render(request, 'refill_form.html', {'form': form})




def refill_update(request, pk):
    """
    Update an existing refill entry and auto-recalculate next appointment.
    """
    refill = get_object_or_404(Refill, pk=pk)
    form = RefillForm(request.POST or None, instance=refill)

    if form.is_valid():
        refill = form.save(commit=False)

        # Auto recalculate next appointment
        refill.next_appointment = (
            refill.last_pickup_date +
            timedelta(days=refill.months_of_refill_days)
        )

        refill.save()
        return redirect('refill_list')

    return render(request, "refill_form.html", {"form": form})












def refill_add_or_update(request, pk=None):
    today = timezone.now().date()

    # If editing an existing refill, get it by pk
    if pk:
        refill = get_object_or_404(Refill, pk=pk)
    else:
        refill = Refill()
        refill.calculate_dates()  # Pre-calculate next_appointment and IIT for new record

    if request.method == "POST":
        form = RefillForm(request.POST, instance=refill)
        if form.is_valid():
            form.save()
            return redirect("refill_list")
    else:
        form = RefillForm(instance=refill)

    return render(
        request,
        "refill_form.html",
        {"form": form, "today": today}
    )



# ================================
# UPLOAD VIEW
# ================================









# def upload_excel(request):
#     if request.method == 'POST':
#         form = UploadExcelForm(request.POST, request.FILES)
#         excel_file = request.FILES.get('file')

#         # Optional: check max file size (1 GB)
#         MAX_FILE_SIZE = 1073741824
#         if excel_file and excel_file.size > MAX_FILE_SIZE:
#             return render(request, 'upload.html', {
#                 'form': form,
#                 'error': "File size exceeds the 1GB limit."
#             })

#         if form.is_valid():
#             facility = form.cleaned_data['facility']  # None means All
#             try:
#                 import_refills_from_excel(excel_file, facility)
#                 return redirect('refill_list')
#             except ValidationError as e:
#                 return render(request, 'upload.html', {
#                     'form': form,
#                     'error': str(e)
#                 })
#         else:
#             return render(request, 'upload.html', {'form': form})
#     else:
#         form = UploadExcelForm()

#     return render(request, 'upload.html', {'form': form})




def track_refills(request):
    today = timezone.now().date()
    start_of_week = today - timedelta(days=today.weekday())
    start_of_month = today.replace(day=1)

    # Filters
    facility_id = request.GET.get("facility")
    selected_case_manager = request.GET.get("case_manager")
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")

    # Parse start and end date if provided
    if start_date:
        try:
            start_date_obj = datetime.strptime(start_date, "%Y-%m-%d").date()
        except ValueError:
            start_date_obj = None
    else:
        start_date_obj = None

    if end_date:
        try:
            end_date_obj = datetime.strptime(end_date, "%Y-%m-%d").date()
        except ValueError:
            end_date_obj = None
    else:
        end_date_obj = None

    # All facilities
    facilities = Facility.objects.all()

    # Base queryset
    refills = Refill.objects.all()

    # Apply filters
    if facility_id:
        try:
            refills = refills.filter(facility_id=int(facility_id))
        except ValueError:
            pass

    if selected_case_manager:
        refills = refills.filter(case_manager=selected_case_manager)

    # Filter by date range if provided
    if start_date_obj:
        refills = refills.filter(last_pickup_date__gte=start_date_obj)

    if end_date_obj:
        refills = refills.filter(last_pickup_date__lte=end_date_obj)

    # Group by period
    daily_qs = refills.filter(last_pickup_date=today).order_by('-last_pickup_date')
    weekly_qs = refills.filter(last_pickup_date__gte=start_of_week).order_by('-last_pickup_date')
    monthly_qs = refills.filter(last_pickup_date__gte=start_of_month).order_by('-last_pickup_date')

    # Pagination
    daily_paginator = Paginator(daily_qs, 10)
    weekly_paginator = Paginator(weekly_qs, 10)
    monthly_paginator = Paginator(monthly_qs, 10)

    daily_page = request.GET.get("daily_page")
    weekly_page = request.GET.get("weekly_page")
    monthly_page = request.GET.get("monthly_page")

    daily_refills = daily_paginator.get_page(daily_page)
    weekly_refills = weekly_paginator.get_page(weekly_page)
    monthly_refills = monthly_paginator.get_page(monthly_page)

    # Unique case managers for filter
    case_managers_qs = (
        Refill.objects.exclude(case_manager__isnull=True)
        .exclude(case_manager__exact="")
        .values_list("case_manager", flat=True)
        .distinct()
    )
    case_managers = sorted({cm.strip() for cm in case_managers_qs})

    # Prepare list of periods for template
    periods = [
        ('Daily', daily_refills),
        ('Weekly', weekly_refills),
        ('Monthly', monthly_refills),
    ]

    # Excel export
    if 'download' in request.GET:
        return export_track_refills_to_excel(refills)

    context = {
        "facilities": facilities,
        "selected_facility": facility_id,
        "case_managers": case_managers,
        "selected_case_manager": selected_case_manager,
        "today": today,
        "selected_start_date": start_date,
        "selected_end_date": end_date,
        "periods": periods,
    }

    return render(request, "track_refills.html", context)





def export_track_refills_to_excel(refills):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Track Refills Data"

    headers = ['Unique ID', 'Facility', 'Last Pickup Date', 'Refill Days', 'Sex', 'Current Regimen', 'Case Manager', 'Next Appointment']
    ws.append(headers)

    for refill in refills:
        row = [
            refill.unique_id,
            refill.facility.name if refill.facility else "",
            refill.last_pickup_date,
            refill.months_of_refill_days,
            refill.sex,
            refill.current_regimen,
            refill.case_manager or "",
            refill.next_appointment,
        ]
        ws.append(row)

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename="track_refills.xlsx"'
    wb.save(response)
    return response












def daily_refill_list(request):
    today = timezone.now().date()

    # Filter by Facility (optional)
    facility_id = request.GET.get("facility")
    facilities = Facility.objects.all()

    # Only Daily refills: filter by today's date
    refills = Refill.objects.filter(next_appointment=today).order_by('unique_id')

    if facility_id:
        refills = refills.filter(facility_id=facility_id)

    context = {
        "facilities": facilities,
        "selected_facility": facility_id,
        "today": today,  # pass today for overdue highlighting
        "refills": refills,
    }

    return render(request, "daily_refill_list.html", context)















def missed_refills(request):
    today = timezone.now().date()

    # ================= GET FILTER PARAMETERS =================
    facility_id = request.GET.get("facility")
    case_manager = request.GET.get("case_manager")
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")

    # ================= BASE QUERYSET =================
    refills = Refill.objects.filter(
        current_art_status__in=["Active", "Active Restart"]
    ).select_related("facility")

    # ================= FILTERS =================
    if facility_id:
        try:
            refills = refills.filter(facility_id=int(facility_id))
        except ValueError:
            pass

    if case_manager:
        refills = refills.filter(case_manager=case_manager)

    # ================= DATE FILTER =================
    if start_date:
        try:
            start_date_obj = datetime.strptime(start_date, "%Y-%m-%d").date()
            refills = refills.filter(next_appointment__gte=start_date_obj)
        except ValueError:
            pass

    if end_date:
        try:
            end_date_obj = datetime.strptime(end_date, "%Y-%m-%d").date()
            refills = refills.filter(next_appointment__lte=end_date_obj)
        except ValueError:
            pass

    # ================= MISSED REFILLS LOGIC =================
    missed_list = refills.filter(next_appointment__lt=today).filter(
        Q(last_pickup_date__lt=F("next_appointment")) |
        Q(last_pickup_date__isnull=True)
    ).order_by("next_appointment")

    # ================= CALCULATE DAYS MISSED AND IIT STATUS =================
    for refill in missed_list:
        if refill.next_appointment:
            days_missed = (today - refill.next_appointment).days
            refill.days_missed = days_missed

            if days_missed >= 28:
                refill.iit_status = "IIT"
            elif days_missed > 0:
                days_remaining = 28 - days_missed
                refill.iit_status = f"{days_remaining} days to IIT"
            else:
                refill.iit_status = "0"
        else:
            refill.days_missed = 0
            refill.iit_status = "0"

    total_missed = missed_list.count()

    # ================= EXPORT TO EXCEL =================
    if request.GET.get("export") == "excel":
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Missed Refills"

        headers = [
            "Unique ID",
            "Case Manager",
            "Facility",
            "Last Pickup",
            "Next Appointment",
            "Days Missed",
            "IIT Status",
        ]
        worksheet.append(headers)

        for refill in missed_list:
            worksheet.append([
                refill.unique_id,
                refill.case_manager,
                refill.facility.name if refill.facility else "",
                refill.last_pickup_date,
                refill.next_appointment,
                refill.days_missed,
                refill.iit_status,
            ])

        response = HttpResponse(
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        response["Content-Disposition"] = 'attachment; filename="missed_refills.xlsx"'
        workbook.save(response)
        return response

    # ================= PAGINATION =================
    paginator = Paginator(missed_list, 25)
    page_number = request.GET.get("page")
    page_obj = paginator.get_page(page_number)

    # ================= UNIQUE CASE MANAGERS =================
    case_managers_qs = (
        Refill.objects.exclude(case_manager__isnull=True)
        .exclude(case_manager__exact="")
        .values_list("case_manager", flat=True)
        .distinct()
    )
    case_managers = sorted({cm.strip() for cm in case_managers_qs if cm.strip()})

    # ================= CONTEXT =================
    context = {
        "page_obj": page_obj,
        "today": today,
        "total_missed": total_missed,
        "facilities": Facility.objects.all(),
        "case_managers": case_managers,
        "selected_facility": facility_id,
        "selected_case_manager": case_manager,
        "selected_start_date": start_date,
        "selected_end_date": end_date,
    }

    return render(request, "missed_refills.html", context)
