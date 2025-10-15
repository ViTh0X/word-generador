from utilidades.creador_word1 import CreadorWord
import pandas as pd
import re


class Secretobancario1(CreadorWord):
    def __init__(self,archivo_excel):
        super().__init__()
        self.archivo = archivo_excel
        self.dataCuadros = {}
        self.buffers = []        
    
    def generar_word_1(self):
        df = pd.read_excel(self.archivo,dtype={"N° Documento Identidad":str})
        df = df.sort_values(by="N° Envío",ascending=True)
        df["N° Oficio de la Autoridad"] = df["N° Oficio de la Autoridad"].fillna("X")
        listaCodEnvio = df["N° Oficio de la Autoridad"].tolist()                    
        listaLimpia = [str(elemento).replace('_x000D_','').replace('\n',' ').replace('/',' ').replace(':',' ').replace('*',' ').replace('?',' ').replace('"',' ').replace('<',' ').replace('>',' ').replace('|',' ') for elemento in listaCodEnvio] 
        ultimoCodigoUsado =""
        posicion = 0
        try:
            for fila, columna  in df.iterrows():
                numeroOficio = columna['N° Oficio de la Autoridad'].replace('_x000D_','').replace('\n',' ').strip()                    
                #***************Nombre del documento se usara como codigo Unico**************
                nombreDocumento = numeroOficio.replace('/',' ').replace(':',' ').replace('*',' ').replace('?',' ').replace('"',' ').replace('<',' ').replace('>',' ').replace('|',' ')
                numeroEnvio = columna["N° Envío"]
                if nombreDocumento != ultimoCodigoUsado:                                            
                    super().__init__()
                    self.crearContenido()
                    self.agregarTitulo("Datos de Identificacion y Ubicacion :",1)                
                    self.agregarTextoNegrita("Número de envio: ")
                    self.agregarTexto(str(numeroEnvio).replace('_x000D_','').replace('\n',' ').strip())                    
                    self.agregarTextoNegrita("Fecha de Envio de Paquete: ") 
                    self.agregarTexto(str(columna['Fecha de Envío']).replace('_x000D_','').replace('\n',' ').strip())
                    self.agregarTextoNegrita("Tipo de solicitud: ")
                    self.agregarTexto(str(columna['Tipo Solicitud']).replace('_x000D_','').replace('\n',' ').strip())
                    self.agregarTextoNegrita("Número de Expediente: ")                     
                    self.agregarTexto(str(columna['N° Expediente SBS']).replace('_x000D_','').replace('\n',' ').strip())
                    self.agregarTextoNegrita("Entidad Solicitante: ")                
                    self.agregarTexto(str(columna['Entidad Solicitante']).replace('_x000D_','').replace('\n',' ').strip())
                    self.agregarTextoNegrita("Nombre de la Autoridad: ")                    
                    self.agregarTexto(str(columna['Nombre de la Autoridad']).replace('_x000D_','').replace('\n',' ').strip())
                    self.agregarTextoNegrita("Número de oficio de la autoridad: ")
                    numeroOficio = columna['N° Oficio de la Autoridad'].replace('_x000D_','').replace('\n',' ').strip()                    
                    #***************Nombre del documento**************
                    nombreDocumento = numeroOficio.replace('/',' ').replace(':',' ').replace('*',' ').replace('?',' ').replace('"',' ').replace('<',' ').replace('>',' ').replace('|',' ')
                    #************************************************
                    self.agregarTexto(str(columna['N° Oficio de la Autoridad']).replace('_x000D_','').replace('\n',' ').strip())
                    self.agregarTextoNegrita("Dirección de la Autoridad: ")                    
                    self.agregarTexto(str(columna['Dirección Autoridad']).replace('_x000D_','').replace('\n',' ').strip())
                    self.agregarTextoNegrita("Delito / Materia: ")                    
                    self.agregarTexto(str(columna['Delito / Materia']).replace('_x000D_','').replace('\n',' ').strip())
                    self.agregarTextoNegrita("Información requerida: ")                    
                    self.agregarTexto(str(columna['Información Requerida']).replace('_x000D_','').replace('\n',' ').strip())
                    self.agregarTextoNegrita("Información adicional: ")
                    texto = str(columna['Información Adicional']).replace('_x000D_',' ')                   
                    texto = re.sub(r'\s+',' ',texto).strip()
                    self.agregarTexto(texto)
                    self.agregarTextoNegrita("N° Expediente/Carpeta Fiscal/Caso: ")                    
                    self.agregarTexto(str(columna['N° Expediente / Carpeta Fiscal / Caso']).replace('_x000D_','').replace('\n',' ').strip())
                    self.agregarTextoNegrita("Prioridad Alta: ")                    
                    self.agregarTexto(str(columna['Prioridad Alta']).replace('_x000D_','').replace('\n',' ').strip())
                    self.agregarTextoNegrita("Plazo de respuesta: ")                    
                    if str(columna['Tipo Plazo Atención']) == "NaN" or  str(columna['Tipo Plazo Atención']) == "nan":
                        tipoPlazoAtencion = " "
                    else:
                        tipoPlazoAtencion = str(columna['Tipo Plazo Atención'])
                    data = f"{columna['Plazo de Atención']} {tipoPlazoAtencion}"
                    self.agregarTexto(str(data).replace('_x000D_','').replace('\n',' ').strip())
                    self.agregarTextoNegrita("Periodo de consulta: ")                    
                    self.agregarTexto(str(columna['Precisa Periodo de Consulta']).replace('_x000D_','').replace('\n',' ').strip())
                    
                    self.agregarTitulo("Personas incluidas en la solicitud : ")
                    self.crearTabla()
                    self.estiloTabla()
                    self.dataCuadros["Nombre"] = str(columna['Nombre sin especificar'])
                    self.dataCuadros["Tipo Documento"] = str(columna['Tipo Documento Identidad'])
                    self.dataCuadros["N° Documento"] = str(columna['N° Documento Identidad'])                    
                    self.dataCuadros["Pais"] = str(columna['Pais'])
                    self.dataCuadros["Observación"] = str(columna['Observación'])
                    self.ingresarDataTabla(self.dataCuadros)
                    print(f"{posicion+1} --- {len(listaLimpia)-1}")
                    if posicion+1 <= len(listaLimpia)-1:                        
                        if str(nombreDocumento) == str(listaLimpia[posicion+1]).strip():
                            ultimoCodigoUsado = nombreDocumento
                            self.dataCuadros = {}
                            posicion +=1
                            continue                
                        else:                            
                            ultimoCodigoUsado = nombreDocumento
                            self.dataCuadros = {}
                            self.agregarEspaciado()
                            #self.guardarDocumento(nombreDocumento)        
                            self.documento.save(self.buffer)
                            self.buffer.seek(0)
                            self.buffers.append((f"{nombreDocumento}.docx",self.buffer))
                            posicion +=1
                    else:
                        print("Ingreso al guardado 2")
                        self.dataCuadros = {}
                        self.agregarEspaciado()
                        #self.guardarDocumento(nombreDocumento)                        
                        self.documento.save(self.buffer)
                        self.buffer.seek(0)
                        self.buffers.append((f"{nombreDocumento}.docx",self.buffer))
                        
                else:
                    self.dataCuadros["Nombre"] = str(columna['Nombre sin especificar'])
                    self.dataCuadros["Tipo Documento"] = str(columna['Tipo Documento Identidad'])
                    self.dataCuadros["N° Documento"] = str(columna['N° Documento Identidad'])
                    self.dataCuadros["Pais"] = str(columna['Pais'])
                    self.dataCuadros["Observación"] = str(columna['Observación'])
                    self.ingresarDataTabla(self.dataCuadros)                    
                    if posicion+1 <= len(listaLimpia)-1:                                
                        if nombreDocumento == str(listaLimpia[posicion+1]).strip():
                            ultimoCodigoUsado = nombreDocumento
                            self.dataCuadros = {}
                            posicion +=1
                            continue                
                        else:   
                            print("Ingreso al guardado 3")
                            ultimoCodigoUsado = nombreDocumento
                            self.dataCuadros = {}
                            self.agregarEspaciado()
                            #self.guardarDocumento(nombreDocumento)       
                            self.documento.save(self.buffer)
                            self.buffer.seek(0)
                            self.buffers.append((f"{nombreDocumento}.docx",self.buffer))                        
                            posicion +=1
                    else:
                        print("Ingreso al guardado 4")
                        self.dataCuadros = {}
                        self.agregarEspaciado()
                        #self.guardarDocumento(nombreDocumento)                        
                        self.documento.save(self.buffer)
                        self.buffer.seek(0)
                        self.buffers.append((f"{nombreDocumento}.docx",self.buffer))
                        
        except Exception as e:
            print(f"Errores es********** \n{e}")
            
    def generar_word_2(self,dia,mes,año,correlativo):
        df = pd.read_excel(self.archivo,dtype={"N° Documento Identidad":str})
        df["N° Oficio de la Autoridad"] = df["N° Oficio de la Autoridad"].fillna("X")
        listaCodEnvio = df["N° Oficio de la Autoridad"].tolist()        
        listaLimpia = [str(elemento).replace('_x000D_','').replace('\n',' ').replace('/',' ').replace(':',' ').replace('*',' ').replace('?',' ').replace('"',' ').replace('<',' ').replace('>',' ').replace('|',' ') for elemento in listaCodEnvio]                        
        ultimoCodigoUsado =""
        posicion = 0
        numero_correlativo = int(correlativo)
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
        try:
            super().__init__()
            self.crear_cabezera()
            self.crear_piepagina()
            for fila, columna  in df.iterrows():
                numeroOficio = columna['N° Oficio de la Autoridad'].replace('_x000D_','').replace('\n',' ').strip()                                    
                nombreDocumento = numeroOficio.replace('/',' ').replace(':',' ').replace('*',' ').replace('?',' ').replace('"',' ').replace('<',' ').replace('>',' ').replace('|',' ')            
                if nombreDocumento != ultimoCodigoUsado:                    
                    try:                        
                        contador = 1                                  
                        self.titulo_secreto_bancario(mes_texto,año,numero_correlativo)
                        self.añadir_parte1(str(columna['Entidad Solicitante']).replace('_x000D_','').replace('\n',' ').strip().upper())
                        self.parrafo_contenido_princial()                    
                        self.agregar_contenidos("ASUNTO	: ","Remite información sobre Levantamiento del Secreto Bancario",1)
                        texto = str(columna['Nombre de la Autoridad']).replace('_x000D_','').replace('\n',' ').strip()
                        self.agregar_contenidos("ATENCIÓN	: ",texto,1)
                        texto = str(columna['N° Oficio de la Autoridad']).replace('_x000D_','').replace('\n',' ').strip()
                        text_añadido = f"Oficio {texto}"
                        self.agregar_contenidos("REFERENCIA	: ",text_añadido,1)
                        texto = str(columna['N° Expediente / Carpeta Fiscal / Caso']).replace('_x000D_','').replace('\n',' ').strip()
                        self.agregar_contenidos("		  ",texto,2)
                        self.escrito_final1()
                        self.escrito_final2()
                        self.tabla_secreto_bancario()                        
                        try:                    
                            self.dataCuadros["Contador"] = str(contador)
                            self.dataCuadros["Nombre"] = str(columna['Nombre sin especificar'])
                            self.dataCuadros["Tipo Documento"] = str(columna['Tipo Documento Identidad'])
                            self.dataCuadros["N° Documento"] = str(columna['N° Documento Identidad'])
                            self.añadir_fila_sb(self.dataCuadros)
                        except Exception as e:
                            print(e)
                        if posicion+1 <= len(listaLimpia)-1:
                            if str(nombreDocumento) == str(listaLimpia[posicion+1]).strip():
                                ultimoCodigoUsado = nombreDocumento
                                self.dataCuadros = {}
                                posicion +=1
                                continue                
                            else:              
                                #print("Ingreso parte final y salto de pagina1")              
                                ultimoCodigoUsado = nombreDocumento
                                self.dataCuadros = {}
                                self.agregar_texto_normal()
                                self.agregar_texto_derecha(dia_texto,mes_texto,año)
                                self.salto_pagina()
                                numero_correlativo += 1
                                # aqui debe colocar sin otro en particular y la fecha y luego
                                #aqui debo añadir salto de pagina               
                                posicion +=1
                        else:
                            #print("Ingreso a cerrar el documento1")              
                            self.dataCuadros = {}                        
                            #self.guardarDocumento(nombreDocumento)                
                            self.agregar_texto_normal()
                            self.agregar_texto_derecha(dia_texto,mes_texto,año)        
                            self.documento.save(self.buffer)
                            self.buffer.seek(0)
                    except Exception as e:
                        print(e)                        
                else:
                    contador += 1                    
                    self.dataCuadros["Contador"] = str(contador)
                    self.dataCuadros["Nombre"] = str(columna['Nombre sin especificar'])
                    self.dataCuadros["Tipo Documento"] = str(columna['Tipo Documento Identidad'])
                    self.dataCuadros["N° Documento"] = str(columna['N° Documento Identidad'])
                    self.añadir_fila_sb(self.dataCuadros)               
                    if posicion+1 <= len(listaLimpia)-1:                                   
                        ndoc = str(listaLimpia[posicion+1]).strip()                        
                        if nombreDocumento == str(listaLimpia[posicion+1]).strip():
                            ultimoCodigoUsado = nombreDocumento
                            self.dataCuadros = {}
                            posicion +=1
                            continue                
                        else:                                           
                            ultimoCodigoUsado = nombreDocumento
                            self.agregar_texto_normal()
                            self.agregar_texto_derecha(dia_texto,mes_texto,año)
                            self.salto_pagina()
                            numero_correlativo += 1                       
                            posicion +=1
                    else:
                        #print("Ingreso a cerrar el documento2")              
                        self.dataCuadros = {}                        
                        #self.guardarDocumento(nombreDocumento)
                        self.agregar_texto_normal()
                        self.agregar_texto_derecha(dia_texto,mes_texto,año)                                                       
                        self.documento.save(self.buffer)
                        self.buffer.seek(0)                                            
                    
        except Exception as e:
            print(f"Error generando word2 {e}")