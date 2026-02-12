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








# Allowed refill months
VALID_REFILL_MONTHS = [1, 2, 3, 6]

# Function to handle importing refills
def import_refills_from_excel(file, facility):
    """
    This function handles uploading and replacing refills from the provided Excel file.
    It first removes existing refills for the provided facility before adding the new ones.
    """

    # Check the file size before processing
    if file.size > settings.FILE_UPLOAD_MAX_MEMORY_SIZE:
        raise ValidationError("File size exceeds the maximum allowed limit of 1GB.")

    # Read the Excel file
    df = pd.read_excel(file)

    required_columns = [
        'Unique Id',
        'Last Pickup Date (yyyy-mm-dd)',
        'Months of ARV Refill',
        'Current ART Regimen',
        'Case Manager',
        'Sex',
        'Current ART Status'
    ]

    # Check if required columns are present
    for column in required_columns:
        if column not in df.columns:
            raise ValidationError(f"Missing column: {column}")

    # Filter only Active and Active Restart patients
    df = df[df['Current ART Status'].isin(['Active', 'Active Restart'])]

    if df.empty:
        raise ValidationError("No Active or Active Restart patients found.")

    # Step 1: Delete previous refills for the given facility to avoid duplicates
    Refill.objects.filter(facility=facility).delete()

    # Step 2: Re-upload new data from the Excel file
    with transaction.atomic():
        for _, row in df.iterrows():
            if pd.isnull(row['Last Pickup Date (yyyy-mm-dd)']):
                raise ValidationError(
                    f"Missing Last Pickup Date for Unique Id {row['Unique Id']}"
                )

            months = int(row['Months of ARV Refill'])
            if months not in VALID_REFILL_MONTHS:
                raise ValidationError(
                    f"Invalid refill duration {months} months "
                    f"for Unique Id {row['Unique Id']}. "
                    f"Allowed: 1, 2, 3, 6"
                )

            # Convert months to refill days
            refill_days = months * 30

            # Properly convert the pickup date
            last_pickup_date = pd.to_datetime(row['Last Pickup Date (yyyy-mm-dd)']).date()

            # Auto calculate next appointment based on the refill days
            next_appointment = last_pickup_date + timedelta(days=refill_days)

            # Create or update the refill entry in the database
            Refill.objects.update_or_create(
                facility=facility,
                unique_id=row['Unique Id'],
                defaults={
                    'last_pickup_date': last_pickup_date,
                    'next_appointment': next_appointment,
                    'months_of_refill_days': refill_days,
                    'current_regimen': row['Current ART Regimen'],
                    'case_manager': row['Case Manager'],
                    'sex': row['Sex'],
                }
            )


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

        "today": today
    }

    return render(request, "dashboard.html", context)


# ================================
# CRUD VIEWS
# ================================





from django.utils import timezone



def refill_list(request):
    today = timezone.now().date()
    week_end = today + timedelta(days=7)

    # Current month range
    month_start = today.replace(day=1)
    if today.month == 12:
        month_end = today.replace(year=today.year + 1, month=1, day=1) - timedelta(days=1)
    else:
        month_end = today.replace(month=today.month + 1, day=1) - timedelta(days=1)

    # Get the first day of the previous month
    last_month_start = (today.replace(day=1) - timedelta(days=1)).replace(day=1)
    last_month_end = today.replace(day=1) - timedelta(days=1)  # Last day of the previous month

    facility_id = request.GET.get("facility")
    facilities = Facility.objects.all()

    refills = Refill.objects.all()

    if facility_id:
        refills = refills.filter(facility_id=facility_id)

    # Adjust for previous month's last pickup date filtering
    previous_month_refills = refills.filter(
        last_pickup_date__range=[last_month_start, last_month_end]
    )

    # Excel Export Functionality
    if 'download' in request.GET:
        return export_refills_to_excel(refills)

    context = {
        "facilities": facilities,
        "selected_facility": facility_id,
        "today": today,
        "daily_expected": refills.filter(next_appointment=today),
        "weekly_expected": refills.filter(next_appointment__range=[today, week_end]),
        "monthly_expected": refills.filter(
            next_appointment__year=today.year,
            next_appointment__month=today.month
        ),
        "previous_month_refills": previous_month_refills,  # Add previous month data
    }

    return render(request, "refill_list.html", context)




def export_refills_to_excel(refills):
    # Get today's date
    today = timezone.now().date()

    # Calculate the last day of the current month
    month_end = today.replace(day=1) + timedelta(days=32)  # Get the first day of next month
    month_end = month_end.replace(day=1) - timedelta(days=1)  # Get the last day of the current month

    # Create a new Excel workbook and sheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Expected Refills Data"

    # Define header row (without 'Last Pickup Date' and 'Refill Days')
    headers = ['Unique ID', 'Facility', 'Sex', 'Current Regimen', 'Case Manager', 'Next Appointment']
    ws.append(headers)

    # Add refill data
    for refill in refills:
        # Calculate the Next Appointment by adding the refill duration (in days) to the Last Pickup Date
        last_pickup_date = refill.last_pickup_date  # This is still used internally to calculate next appointment
        refill_days = refill.months_of_refill_days  # This is in days (e.g., 90 days)
        next_appointment = last_pickup_date + timedelta(days=refill_days)

        # Ensure Next Appointment is within the current month
        if next_appointment > month_end:
            next_appointment = month_end

        # Row to be added to Excel
        row = [
            refill.unique_id,
            refill.facility.name,
            refill.sex,
            refill.current_regimen,
            refill.case_manager,
            next_appointment,  # Only showing the Next Appointment
        ]
        ws.append(row)

    # Create HTTP response and set content type for Excel
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename="Expected_Refills.xlsx"'

    # Save workbook to response
    wb.save(response)
    return response




def refill_create(request, unique_id=None):
    if unique_id:
        # Fetch the refill by unique_id if passed
        refill = get_object_or_404(Refill, unique_id=unique_id)
    else:
        refill = None  # New refill if no unique_id is passed

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



# ================================
# UPLOAD VIEW
# ================================




from .forms import UploadExcelForm


def upload_excel(request):
    if request.method == 'POST':
        form = UploadExcelForm(request.POST, request.FILES)

        # Checking for file size manually (optional, since Django settings already limit it)
        MAX_FILE_SIZE = 1073741824  # 1 GB (in bytes)
        file = request.FILES.get('file')
        
        if file.size > MAX_FILE_SIZE:
            return HttpResponse("File size exceeds the 1GB limit.", status=400)

        if form.is_valid():
            facility = form.cleaned_data['facility']
            excel_file = form.cleaned_data['file']
            
            try:
                # Handle large file uploads efficiently using pandas
                import_refills_from_excel(excel_file, facility)  # Process the file
                return redirect('refill_list')  # Redirect to the refills list after successful upload
            except ValidationError as e:
                return render(request, 'upload.html', {'form': form, 'error': str(e)})
        else:
            return render(request, 'upload.html', {'form': form})
    
    else:
        form = UploadExcelForm()

    return render(request, 'upload.html', {'form': form})









def track_refills(request):
    today = timezone.now().date()

    # Calculate the start of the week (assuming Sunday is the first day of the week)
    start_of_week = today - timedelta(days=today.weekday())

    # Calculate the start of the month (first day of the current month)
    start_of_month = today.replace(day=1)

    # Query refills that happened today, this week, and this month
    # Use exact lookup on the DateField
    daily_refills = Refill.objects.filter(last_pickup_date=today)  # Use 'exact' lookup by default
    weekly_refills = Refill.objects.filter(last_pickup_date__gte=start_of_week)
    monthly_refills = Refill.objects.filter(last_pickup_date__gte=start_of_month)

    # Excel Export Functionality
    if 'download' in request.GET:
        return export_track_refills_to_excel(daily_refills, weekly_refills, monthly_refills)

    context = {
        'daily_refills': daily_refills,
        'weekly_refills': weekly_refills,
        'monthly_refills': monthly_refills,
        'today': today,
    }

    return render(request, 'track_refills.html', context)





def export_track_refills_to_excel(daily_refills, weekly_refills, monthly_refills):
    # Create a new Excel workbook and sheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Track Refills Data"

    # Define header row (without 'Refill Type')
    headers = ['Unique ID', 'Facility', 'Last Pickup Date', 'Refill Days', 'Sex', 'Current Regimen', 'Case Manager', 'Next Appointment']
    ws.append(headers)

    # Add daily refills data (without 'Refill Type')
    for refill in daily_refills:
        row = [
            refill.unique_id,
            refill.facility.name,
            refill.last_pickup_date,
            refill.months_of_refill_days,
            refill.sex,
            refill.current_regimen,
            refill.case_manager,
            refill.next_appointment,
        ]
        ws.append(row)

    # Add weekly refills data (without 'Refill Type')
    for refill in weekly_refills:
        row = [
            refill.unique_id,
            refill.facility.name,
            refill.last_pickup_date,
            refill.months_of_refill_days,
            refill.sex,
            refill.current_regimen,
            refill.case_manager,
            refill.next_appointment,
        ]
        ws.append(row)

    # Add monthly refills data (without 'Refill Type')
    for refill in monthly_refills:
        row = [
            refill.unique_id,
            refill.facility.name,
            refill.last_pickup_date,
            refill.months_of_refill_days,
            refill.sex,
            refill.current_regimen,
            refill.case_manager,
            refill.next_appointment,
        ]
        ws.append(row)

    # Create HTTP response and set content type for Excel
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename="track_refills.xlsx"'

    # Save workbook to response
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
