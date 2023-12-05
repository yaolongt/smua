from django import forms
from django.shortcuts import render

class InputForm(forms.Form):
    gv_file = forms.FileField(label="Upload GVSession Excel sheet")
    schedule_file = forms.FileField(label="Upload Manage Schedule Excel sheet")
    enrollment_summary_file = forms.FileField(label="Upload Enrollment Summary Excel sheet")

def home (request):
    input_form = InputForm()

    return render(request, 'home.html', {'input_form': input_form})
