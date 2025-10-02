from django.shortcuts import render
from django.contrib import messages

import pandas as pd

from io import BytesIO

# Create your views here.
def main(request):    
    return render(request,'main/main.html')

'''def generar_word1(request):
    if request.method == 'POST':
        form = formSubirExcel(request.POST, request.FILES)
        if form.is_valid():
            archivo_excel = request.FILES['excel_file']
            
            
        try:
            df = pd.read_excel(BytesIO(archivo_excel.read()))
        except Exception as e:
            messages.error(request,f'Ocurrio un error {e} al procesar el archivo')
        form  = formSubirExcel()
    else:
        form = formSubirExcel()
    return render(request,'main/main.html',{'form':form})
        
        '''
            
    