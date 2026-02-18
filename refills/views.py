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
from django.contrib import messages



# views.py

from io import BytesIO



VALID_REFILL_MONTHS = [0.5, 1, 2, 2.8, 3, 4, 5, 6]


def import_refills_from_excel(file):
    df = pd.read_excel(file)

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

    df = df[df['Current ART Status'].isin(['Active', 'Active Restart'])]

    if df.empty:
        raise ValidationError("No Active or Active Restart patients found.")

    df['Facility Name'] = df['Facility Name'].astype(str).str.strip()
    facility_names = df['Facility Name'].unique()

    facilities = Facility.objects.filter(name__in=facility_names)

    if not facilities.exists():
        raise ValidationError("No matching facilities found in database.")

    facility_map = {f.name: f for f in facilities}

    new_refills = []

    for _, row in df.iterrows():

        facility = facility_map.get(row['Facility Name'])

        if not facility:
            raise ValidationError(f"Facility not found: {row['Facility Name']}")

        last_pickup = pd.to_datetime(row['Last Pickup Date (yyyy-mm-dd)']).date()
        months = float(row['Months of ARV Refill'])

        next_appointment = last_pickup + timedelta(days=months * 30)

        new_refills.append(
            Refill(
                facility=facility,
                unique_id=row['Unique Id'],
                last_pickup_date=last_pickup,
                months_of_refill_days=months,
                next_appointment=next_appointment,
                current_regimen=row['Current ART Regimen'],
                case_manager=row['Case Manager'],
                sex=row['Sex'],
                current_art_status=row['Current ART Status'],
            )
        )

    # 🔥 Only now delete and insert inside transaction
    with transaction.atomic():
        facility_ids = facilities.values_list('id', flat=True)
        Refill.objects.filter(facility_id__in=facility_ids).delete()
        Refill.objects.bulk_create(new_refills, batch_size=1000)

    return len(new_refills)





def upload_excel(request):
    if request.method == 'POST':
        form = UploadExcelForm(request.POST, request.FILES)

        if not request.FILES:
            messages.error(request, "No file was uploaded.")
            return redirect('upload_excel')

        if form.is_valid():
            excel_file = form.cleaned_data['file']

            try:
                import_refills_from_excel(excel_file)
                messages.success(request, "Excel uploaded and processed successfully!")
                return redirect('upload_excel')

            except Exception as e:
                messages.error(request, f"Upload failed: {str(e)}")
                return redirect('upload_excel')

        else:
            messages.error(request, "Form validation failed.")
            print("FORM ERRORS:", form.errors)

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

    # =============================
    # GET FILTERS
    # =============================
    facility_id = request.GET.get("facility")
    selected_case_manager = request.GET.get("case_manager")
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")

    # =============================
    # PARSE DATE FILTERS
    # =============================
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

    # =============================
    # LOAD DATA
    # =============================
    facilities = Facility.objects.all()
    refills = Refill.objects.all()

    # =============================
    # APPLY FILTERS
    # =============================
    if facility_id:
        try:
            refills = refills.filter(facility_id=int(facility_id))
        except ValueError:
            pass

    if selected_case_manager:
        refills = refills.filter(case_manager=selected_case_manager)

    if start_date_obj:
        refills = refills.filter(next_appointment__gte=start_date_obj)

    if end_date_obj:
        refills = refills.filter(next_appointment__lte=end_date_obj)

    # =============================
    # CASE MANAGER DROPDOWN LIST
    # =============================
    case_managers_qs = (
        Refill.objects.exclude(case_manager__isnull=True)
        .exclude(case_manager__exact="")
        .values_list("case_manager", flat=True)
        .distinct()
    )

    case_managers = sorted(
        {cm.strip() for cm in case_managers_qs if cm and cm.strip()}
    )

    # =============================
    # CALCULATE MISSED + DAYS MISSED
    # =============================
    for refill in refills:
        if (
            refill.next_appointment
            and refill.next_appointment < today
        ):
            refill.days_missed = (today - refill.next_appointment).days
            refill.missed_appointment = True
        else:
            refill.days_missed = 0
            refill.missed_appointment = False

    # =============================
    # PURE PYTHON RISK PREDICTION
    # =============================

    high_risk_keywords = [
        "transport", "money", "no money", "travel",
        "forgot", "busy", "work", "distance",
        "sick", "hospital", "admitted",
        "defaulted", "stopped", "side effect"
    ]

    medium_risk_keywords = [
        "delay", "reschedule", "family issue",
        "school", "appointment clash",
        "funeral", "religious"
    ]

    for refill in refills:

        score = 0

        # 1️⃣ Past Missed Appointment
        if refill.missed_appointment:
            score += 40

        # 2️⃣ Days Missed Impact
        if refill.days_missed > 30:
            score += 25
        elif refill.days_missed > 7:
            score += 15

        # 3️⃣ Remark Text Analysis
        if refill.remark:
            remark_lower = refill.remark.lower()

            for word in high_risk_keywords:
                if word in remark_lower:
                    score += 20

            for word in medium_risk_keywords:
                if word in remark_lower:
                    score += 10

        # 4️⃣ ART Status Impact
        if refill.current_art_status == "Inactive":
            score += 30
        elif refill.current_art_status == "Active Restart":
            score += 20

        # Cap at 100%
        score = min(score, 100)

        refill.prediction_probability = score

    # =============================
    # GROUP BY PERIOD
    # =============================
    daily_expected = refills.filter(next_appointment=today)
    weekly_expected = refills.filter(
        next_appointment__range=[today, week_end]
    )
    monthly_expected = refills.filter(
        next_appointment__year=today.year,
        next_appointment__month=today.month
    )

    # =============================
    # PAGINATION
    # =============================
    daily_page = Paginator(
        daily_expected.order_by("next_appointment"), 10
    )
    weekly_page = Paginator(
        weekly_expected.order_by("next_appointment"), 10
    )
    monthly_page = Paginator(
        monthly_expected.order_by("next_appointment"), 10
    )

    daily_number = request.GET.get("daily_page")
    weekly_number = request.GET.get("weekly_page")
    monthly_number = request.GET.get("monthly_page")

    # =============================
    # CONTEXT
    # =============================
    context = {
        "facilities": facilities,
        "selected_facility": facility_id,
        "case_managers": case_managers,
        "selected_case_manager": selected_case_manager,
        "today": today,
        "selected_start_date": start_date,
        "selected_end_date": end_date,
        "periods": [
            {
                "name": "Daily",
                "page_obj": daily_page.get_page(daily_number),
            },
            {
                "name": "Weekly",
                "page_obj": weekly_page.get_page(weekly_number),
            },
            {
                "name": "Monthly",
                "page_obj": monthly_page.get_page(monthly_number),
            },
        ],
    }

    # =============================
    # EXCEL EXPORT
    # =============================
    if "download" in request.GET:
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

            iit_date = refill.next_appointment + timedelta(days=28)
            days_to_iit = (iit_date - today).days

            if days_missed >= 28:
                refill.iit_status = "IIT"
            elif days_missed > 0:
                refill.iit_status = f"{days_to_iit} days to IIT"
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

            iit_date = refill.next_appointment + timedelta(days=28)
            days_to_iit = (iit_date - today).days

            if days_missed >= 28:
                refill.iit_status = "IIT"
            elif days_missed > 0:
                refill.iit_status = f"{days_to_iit} days to IIT"
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

        # Make header bold
        from openpyxl.styles import Font
        for col in range(1, len(headers) + 1):
            worksheet.cell(row=1, column=col).font = Font(bold=True)

        for refill in missed_list:
            worksheet.append([
                refill.unique_id,
                refill.case_manager or "",
                refill.facility.name if refill.facility else "",
                refill.last_pickup_date.strftime("%Y-%m-%d") if refill.last_pickup_date else "",
                refill.next_appointment.strftime("%Y-%m-%d") if refill.next_appointment else "",
                refill.days_missed,
                refill.iit_status,
            ])

        # Auto column width
        for column_cells in worksheet.columns:
            length = max(len(str(cell.value)) for cell in column_cells if cell.value)
            worksheet.column_dimensions[column_cells[0].column_letter].width = length + 4

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

 