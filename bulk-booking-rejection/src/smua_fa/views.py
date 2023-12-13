from django.shortcuts import render
from django import forms
import pandas as pd
import numpy as np
import re
from django.http import HttpResponse
import io
import zipfile

headers = None
success_file, error_file = "Bulk Booking Template(Success).csv", "Bulk Booking Template(Error).csv"

class InputForm(forms.Form):
  input_field = forms.CharField(widget=forms.Textarea)
  file = forms.FileField(label="Upload CSV or Excel Sheet")
  
def separate_files(success_data, error_data):
    success_data_response, error_data_response = None, None
    if len(success_data) > 0:
        success = pd.DataFrame(success_data, columns=headers)
        success_buffer = io.BytesIO()
        success.to_csv(success_buffer, index=False, date_format="%d-%b-%Y")
        success_buffer.seek(0)
        success_data_response = HttpResponse(success_buffer.getvalue(), content_type='application/csv')
        success_data_response['Content-Disposition'] = f'attachment; filename={success_file}'

    if len(error_data) > 0:
        err = pd.DataFrame(error_data, columns=headers)
        err_buffer = io.BytesIO()
        err.to_csv(err_buffer, index=False, date_format="%d-%b-%Y")
        err_buffer.seek(0)
        error_data_response = HttpResponse(err_buffer.getvalue(), content_type='application/csv')
        error_data_response['Content-Disposition'] = f'attachment; filename={error_file}'

    return success_data_response, error_data_response
  
def split_data(original_data, error_msg):
  error_idx = set()
  for error in error_msg:
    match = re.search(r'Row \d+', error)
    if match:
      error_idx.add(int(match.group().split()[1],10)-1)
  
  error_data = []
  success_data = []
  for i in range(len(original_data)):
    if i in error_idx:
      error_data.append(original_data[i])
    else:
      success_data.append(original_data[i])
      
  return success_data, error_data

def manipulate_excel_data(data):
  """
    Function to split Excel data into comma-separated array for easier manipulation. 
    Excel sheet is unlike CSV data, it is not split by comma.
  """
  csv_data = data.to_csv(index=False).strip().splitlines()
  csv_data_arr = [line.split(",") for line in csv_data]
  np_arr = np.asarray(csv_data_arr, dtype="object")

  return np_arr
  

def home(request):
  global headers
  content = None
  input_form = InputForm()
  combined_response = None

  if request.method == "POST":
    input_form = InputForm(request.POST, request.FILES)
    if input_form.is_valid():
      input_value = input_form.cleaned_data["input_field"].split("\n")
        
      file = input_form.cleaned_data["file"]
      file_extension = file.name.split('.')[-1]

      # check file extensions
      if file_extension == 'xlsx':
        df = pd.read_excel(file).fillna('-')
        np_arr = np.asarray(df, dtype='object')
        headers = df.columns
        
        success_data, error_data = split_data(np_arr[1:], input_value)

        success_data_response, error_data_response = separate_files(success_data, error_data)

        combined_response = HttpResponse(content_type='application/zip')
        combined_response['Content-Disposition'] = 'attachment; filename="Bulk_Booking_Files.zip'

        with zipfile.ZipFile(combined_response, 'w') as zip_file:
          if success_data_response != None:
                zip_file.writestr(success_file,success_data_response.content)
          if error_data_response != None:
                zip_file.writestr(error_file,error_data_response.content)

        return combined_response
      else:
        content = "Unsupported file type"
        
      return render(request, 'home.html', {'input_form': input_form, 'content': content, 'input_value': input_value})

  return render(request, 'home.html', {'input_form': input_form})
