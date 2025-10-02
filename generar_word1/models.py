from django.db import models

# Create your models here.
from django import forms
# Create your models here.

class formSubirExcel(forms.Form):
    excel_file = forms.FileField(
        label='Suba el archivo excel generado',        
    )