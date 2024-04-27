import os
from django.conf import settings
from django.http import HttpResponse, HttpResponseNotFound, HttpResponseServerError
from django.shortcuts import render
from django.templatetags.static import static
from .utils import *

def upload_cv(request):
    if request.method == 'POST' and request.FILES.getlist('cv_files'):
        cv_files = request.FILES.getlist('cv_files')
        data_to_write = []

        for cv_file in cv_files:
            if cv_file.name.endswith(('.doc', '.docx')):
                text = convert_doc_to_docx_and_extract_text(cv_file)
            elif cv_file.name.endswith('.pdf'):
                text = extract_text_from_pdf(cv_file)
            elif cv_file.name.endswith('.txt'):
                text = extract_text_from_text(cv_file)
            else:
                text = ""

            email = extract_email(text)
            contact_number = extract_contact_number(text)

            # Extract the first email address and contact number (if any)
            email = email[0] if email else ''
            contact_number = contact_number[0] if contact_number else ''

            # Add the extracted data to the list
            data_to_write.append({'Email': email, 'Contact Number': contact_number})

        try:
            # Generate Excel file
            output_filename = 'excel_file.xlsx'
            output_path = os.path.join(settings.MEDIA_ROOT, output_filename)
            create_excel_file(data_to_write, output_path)

            # Construct the URL for the Excel file
            excel_file_url = os.path.relpath(output_path, settings.MEDIA_ROOT)
            excel_file_url = static(excel_file_url)
        except Exception as e:
            # Handle any errors that occur during Excel file generation
            excel_file_url = None
            print("Error generating Excel file:", e)

        return render(request, 'upload_cv.html', {'excel_file_url': excel_file_url})

    return render(request, 'upload_cv.html')


def download_excel(request, file_name):
    try:
        file_path = os.path.join(settings.MEDIA_ROOT, file_name)
        with open(file_path, 'rb') as excel_file:
            response = HttpResponse(excel_file.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = f'attachment; filename="excel_file.xlsx"'
            return response
    except FileNotFoundError:
        return HttpResponseNotFound("Excel file not found.")
    except Exception as e:
        print("Error downloading Excel file:", e)
        return HttpResponseServerError("An error occurred while downloading the Excel file.")
