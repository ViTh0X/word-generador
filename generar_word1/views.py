from django.shortcuts import render
from django.http import HttpResponse

from utilidades.secreto_bancario1 import Secretobancario1
from .models import formSubirExcel
import zipfile
from io import BytesIO

# Create your views here.
def generar_secreto_bancario1(request):
    if request.method == 'POST':
        form = formSubirExcel(request.POST, request.FILES)
        if form.is_valid():
            archivo_excel = request.FILES['excel_file']
            sb = Secretobancario1(archivo_excel)
            sb.generar_word_1()
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer,"w") as zip_file:
                for filename, buffer in sb.buffers:
                    zip_file.writestr(filename,buffer.getvalue())
            zip_buffer.seek(0)
            response = HttpResponse(zip_buffer.getvalue(), content_type='application/zip')
            response["Content-Disposition"] = 'attachment; filename="archivoword_convertido.zip"'
            return response
    else:
        form = formSubirExcel()            
    return render(request,'secreto_bancario1/secreto1.html',{'form':form})