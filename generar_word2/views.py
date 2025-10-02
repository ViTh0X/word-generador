from django.shortcuts import render

from django.http import HttpResponse

from utilidades.secreto_bancario1 import Secretobancario1
from .models import formSubirExcel
import zipfile
from io import BytesIO
from datetime import datetime

def generar_secreto_bancario2(request):
    dia = datetime.now().day
    mes = datetime.now().month    
    año = datetime.now().year
    mes_texto = ""
    if mes < 10:
        mes_texto = f"0{mes}"
    else:
        mes_texto = str(mes)
    
    dia_texto = ""
    if dia < 10:
        dia_texto = f"0{dia}"
    else:
        dia_texto = str(dia)
    if request.method == 'POST':
        form = formSubirExcel(request.POST, request.FILES)
        if form.is_valid():
            archivo_excel = request.FILES['excel_file']
            correlativo = form.cleaned_data['num_correlativo']
            sb = Secretobancario1(archivo_excel)
            sb.generar_word_2(dia,mes,año,correlativo)
            response = HttpResponse(sb.buffer.read(), content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            response["Content-Disposition"] = f'attachment; filename="Levantamiento secreto bancario {dia_texto}-{mes_texto}.docx"'
            return response
    else:
        form = formSubirExcel()            
    return render(request,'secreto_bancario2/secreto2.html',{'form':form})